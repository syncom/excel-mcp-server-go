package excelops

import (
	"errors"
	"fmt"
	"io/fs"
	"syscall"
)

// PyOSError renders a Go filesystem error the way str(OSError) does in CPython:
// "[Errno 2] No such file or directory: '/path/to/file'".
//
// This is not cosmetic. openpyxl lets OSError escape load_workbook, sheet.py
// funnels it into SheetError(str(e)), and server.py interpolates that into
// "Error: {str(e)}" — so the errno text is part of the byte-exact tool result
// (SPEC 5.1). It also appears verbatim in raise-through results.
//
// The message table is CPython's (from the C library's strerror on Linux, which
// CPython uses directly). Go's syscall.Errno.Error() lower-cases the first word
// ("no such file or directory"), so it cannot be used as-is.
func PyOSError(err error) (string, bool) {
	var perr *fs.PathError
	if !errors.As(err, &perr) {
		return "", false
	}
	var errno syscall.Errno
	if !errors.As(perr.Err, &errno) {
		return "", false
	}
	msg, ok := errnoText[errno]
	if !ok {
		return "", false
	}
	return fmt.Sprintf("[Errno %d] %s: '%s'", int(errno), msg, perr.Path), true
}

// errnoText holds the strerror strings CPython reports for the errnos these
// operations can realistically produce. An errno outside the table makes
// PyOSError report failure rather than guess, so a mismatch surfaces as a
// conformance diff instead of a plausible-looking wrong string.
var errnoText = map[syscall.Errno]string{
	syscall.EPERM:        "Operation not permitted",
	syscall.ENOENT:       "No such file or directory",
	syscall.EACCES:       "Permission denied",
	syscall.EEXIST:       "File exists",
	syscall.ENOTDIR:      "Not a directory",
	syscall.EISDIR:       "Is a directory",
	syscall.EINVAL:       "Invalid argument",
	syscall.ENOSPC:       "No space left on device",
	syscall.EROFS:        "Read-only file system",
	syscall.ENAMETOOLONG: "File name too long",
	syscall.ELOOP:        "Too many levels of symbolic links",
	syscall.ENOTEMPTY:    "Directory not empty",
}
