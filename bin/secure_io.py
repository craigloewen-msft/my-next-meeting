#!/usr/bin/env python3
"""
Safe on-disk handling for the two sensitive files this plugin touches:

  ~/.config/omarchy/next-meeting/config.json  - holds the published ICS URL,
                                                which is a *bearer secret*
                                                (anyone with it can read the
                                                calendar).
  ~/.cache/omarchy/next-meeting/token_cache.bin - MSAL refresh/access tokens.

Both used to be read with a plain ``Path.read_text()`` and written with a
plain ``write_text()`` followed by a ``chmod``. That pattern has four
problems, all of which this module exists to fix:

1. **Symlink following.** ``open(path)`` happily follows a symlink, so
   anything that can create a file at that pathname (a shared /home, a
   careless dotfile-sync tool, another compromised app running as the same
   user) can redirect the read anywhere or make us clobber an arbitrary
   file. Every open here uses ``O_NOFOLLOW`` and is bound to a directory
   descriptor we verified first, so the final path component cannot be
   swapped between the check and the open.

2. **No type/owner checks.** A FIFO at the config path would block the poll
   forever; a file owned by someone else has no business supplying our
   credentials. We ``fstat`` the *descriptor we actually opened* (never the
   pathname a second time) and reject anything that isn't a regular,
   single-linked file owned by us.

3. **Unbounded reads.** ``read_text()`` on a 2 GB file is a 2 GB
   allocation inside a widget poll. Reads are capped, and the cap is
   enforced both against ``st_size`` and against the bytes actually
   streamed, so a file growing underneath us can't slip past.

4. **Non-atomic writes with late hardening.** ``write_text()`` then
   ``chmod(0600)`` leaves a window where the token cache is world-readable,
   and a crash mid-write leaves a truncated cache. Writes go to a fresh
   ``O_EXCL`` temp file created *already* at mode 0600, are fsynced, and
   are then ``os.replace``d into position - so readers only ever see a
   complete file and the permissions are never wrong, not even briefly.

Unsafe files are rejected rather than repaired: if the path is a symlink,
a device node, or owned by another user, that is a situation the user needs
to look at, not something to silently overwrite. Merely *loose permissions*
on a file we already own are tightened in place instead, since that is the
common, benign case (a config file created by hand with a 0022 umask) and
failing there would just be a papercut.
"""
import errno
import json
import os
import secrets
import stat

# Directories we own are kept owner-only; so are the files inside them.
DIR_MODE = 0o700
FILE_MODE = 0o600

# Generous ceilings - these files are a few hundred bytes and a few tens of
# kilobytes respectively in normal use. The point is only to keep a
# pathological file from turning into a pathological allocation.
MAX_CONFIG_BYTES = 64 * 1024
MAX_TOKEN_CACHE_BYTES = 4 * 1024 * 1024

_READ_CHUNK = 64 * 1024


class SecureIOError(Exception):
    """A path failed its safety checks and was refused."""


def _reject(path_desc, reason):
    raise SecureIOError(f"{path_desc}: {reason}")


class PrivateDir:
    """
    An open descriptor on a directory we have verified is ours and private.

    Holding the descriptor (rather than re-deriving the path for every
    operation) is the whole point: once the directory is verified, every
    file open below it is resolved relative to that descriptor via
    ``dir_fd=``. An attacker who swaps the *directory* after our check
    cannot affect us, because we never look the directory up by name again.

    Use as a context manager so the descriptor is always closed::

        with PrivateDir(path) as d:
            data = d.read_bytes("config.json", MAX_CONFIG_BYTES)
    """

    def __init__(self, path, create=True):
        self.path = str(path)
        self.fd = -1
        self._open(create)

    # -- lifecycle ---------------------------------------------------------

    def _open(self, create):
        flags = os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW | os.O_CLOEXEC
        try:
            self.fd = os.open(self.path, flags)
        except FileNotFoundError:
            if not create:
                raise
            # mode= applies to the leaf only; intermediates get the usual
            # umask treatment, which is what we want for ~/.config etc.
            os.makedirs(self.path, mode=DIR_MODE, exist_ok=True)
            self.fd = os.open(self.path, flags)
        except OSError as e:
            if e.errno == errno.ELOOP:
                _reject(self.path, "is a symlink, refusing to follow it")
            if e.errno == errno.ENOTDIR:
                _reject(self.path, "exists but is not a directory")
            raise

        try:
            self._verify()
        except BaseException:
            self.close()
            raise

    def _verify(self):
        st = os.fstat(self.fd)
        if not stat.S_ISDIR(st.st_mode):
            _reject(self.path, "is not a directory")
        if st.st_uid != os.geteuid():
            _reject(self.path, f"is owned by uid {st.st_uid}, not by you")
        # Group/other access on a directory holding credentials is worth
        # closing, and doing so via the descriptor can't be redirected.
        if st.st_mode & 0o077:
            os.fchmod(self.fd, DIR_MODE)

    def close(self):
        if self.fd >= 0:
            os.close(self.fd)
            self.fd = -1

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    # -- reads -------------------------------------------------------------

    def read_bytes(self, name, max_bytes):
        """
        Read ``name`` from this directory, or return None if it isn't there.

        Raises SecureIOError if the file is a symlink, is not a regular
        file, is not ours, has extra hard links, or exceeds ``max_bytes``.
        """
        _check_name(name)
        where = os.path.join(self.path, name)
        try:
            # O_NONBLOCK matters here: opening a FIFO read-only otherwise
            # blocks until a writer shows up, which would hang the widget
            # poll indefinitely *before* the fstat below ever gets to
            # reject it. It has no effect on a regular file, which is all
            # we go on to accept anyway.
            fd = os.open(name,
                         os.O_RDONLY | os.O_NOFOLLOW | os.O_NONBLOCK | os.O_CLOEXEC,
                         dir_fd=self.fd)
        except FileNotFoundError:
            return None
        except OSError as e:
            if e.errno == errno.ELOOP:
                _reject(where, "is a symlink, refusing to follow it")
            if e.errno in (errno.ENXIO, errno.ENODEV):
                _reject(where, "is a special file, refusing to read it")
            raise

        try:
            st = os.fstat(fd)
            if not stat.S_ISREG(st.st_mode):
                _reject(where, "is not a regular file")
            if st.st_uid != os.geteuid():
                _reject(where, f"is owned by uid {st.st_uid}, not by you")
            if st.st_nlink > 1:
                _reject(where, "has extra hard links pointing at it")
            if st.st_size > max_bytes:
                _reject(where, f"is {st.st_size} bytes, over the "
                               f"{max_bytes} byte limit")
            if st.st_mode & 0o077:
                # Ours, just too permissive - tighten instead of failing.
                os.fchmod(fd, FILE_MODE)
            return _read_capped(fd, where, max_bytes)
        finally:
            os.close(fd)

    def read_text(self, name, max_bytes):
        data = self.read_bytes(name, max_bytes)
        if data is None:
            return None
        try:
            return data.decode("utf-8")
        except UnicodeDecodeError as e:
            _reject(os.path.join(self.path, name), f"is not valid UTF-8: {e}")

    def read_json(self, name, max_bytes):
        """Read and parse a JSON object, or return {} if absent/unparseable."""
        text = self.read_text(name, max_bytes)
        if text is None:
            return {}
        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            return {}
        return data if isinstance(data, dict) else {}

    # -- writes ------------------------------------------------------------

    def write_bytes(self, name, data):
        """
        Atomically replace ``name`` with ``data`` at mode 0600.

        The temp file is created with O_EXCL at the final mode, so the
        content is never visible to anyone else even momentarily, and the
        rename is atomic so a reader never observes a half-written file.
        """
        _check_name(name)
        tmp = f".{name}.{os.getpid()}.{secrets.token_hex(4)}.tmp"
        fd = os.open(tmp,
                     os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW | os.O_CLOEXEC,
                     FILE_MODE, dir_fd=self.fd)
        try:
            view = memoryview(data)
            while view:
                view = view[os.write(fd, view):]
            os.fsync(fd)
        except BaseException:
            os.close(fd)
            _unlink_quietly(tmp, self.fd)
            raise
        os.close(fd)

        try:
            os.replace(tmp, name, src_dir_fd=self.fd, dst_dir_fd=self.fd)
        except BaseException:
            _unlink_quietly(tmp, self.fd)
            raise
        # Durably record the rename itself, not just the bytes.
        os.fsync(self.fd)

    def write_text(self, name, text):
        self.write_bytes(name, text.encode("utf-8"))


def _check_name(name):
    """Names are fixed constants in this codebase; assert that stays true."""
    if not name or "/" in name or name in (".", ".."):
        raise SecureIOError(f"{name!r} is not a plain file name")


def _read_capped(fd, where, max_bytes):
    """
    Stream at most ``max_bytes`` from an already-verified descriptor.

    We deliberately re-check while reading rather than trusting the earlier
    ``st_size``: a file being appended to between the fstat and the read
    would otherwise sail past the limit.
    """
    chunks = []
    total = 0
    while True:
        chunk = os.read(fd, min(_READ_CHUNK, max_bytes - total + 1))
        if not chunk:
            break
        total += len(chunk)
        if total > max_bytes:
            _reject(where, f"grew past the {max_bytes} byte limit while reading")
        chunks.append(chunk)
    return b"".join(chunks)


def _unlink_quietly(name, dir_fd):
    try:
        os.unlink(name, dir_fd=dir_fd)
    except OSError:
        pass
