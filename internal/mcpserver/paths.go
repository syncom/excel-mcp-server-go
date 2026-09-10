package mcpserver

import (
	"errors"
	"fmt"
	"os"
	"strings"
)

// ValueError mirrors the builtin Python ValueError that get_excel_path raises.
//
// It is deliberately NOT an excelerr kind. No tool's except clause in server.py
// names ValueError, so per SPEC 5.2 a bad path always reaches the client as a
// protocol-level tool error, never as an "Error: ..." text result.
type ValueError struct{ Msg string }

func (e *ValueError) Error() string { return e.Msg }

// ExcelPath is the port of server.get_excel_path (SPEC 5.7).
//
// The two modes differ in more than just absolute-vs-relative: stdio cleans the
// path lexically (normpath) and never touches the filesystem, while HTTP mode
// resolves symlinks on both the base and the candidate (realpath) before
// checking containment. Collapsing the two would either make stdio stat paths
// it should not, or let a symlink inside the sandbox escape it.
func (s *Server) ExcelPath(filename string) (string, error) {
	if filename == "" || strings.ContainsRune(filename, 0) {
		return "", &ValueError{fmt.Sprintf("Invalid filename: %s", filename)}
	}

	// EXCEL_FILES_PATH unset is Python's None: stdio mode.
	if s.cfg.ExcelFilesPath == "" {
		if !pyIsabs(filename) {
			return "", &ValueError{fmt.Sprintf("Invalid filename: %s, must be an absolute path when not in SSE mode", filename)}
		}
		return pyNormpath(filename), nil
	}

	if pyIsabs(filename) {
		return "", &ValueError{fmt.Sprintf("Invalid filename: %s, must be relative to EXCEL_FILES_PATH", filename)}
	}

	base := pyRealpath(s.cfg.ExcelFilesPath)
	candidate := pyRealpath(pyJoin(base, filename))

	if !resolvedPathIsWithin(base, candidate) {
		return "", &ValueError{fmt.Sprintf("Invalid filename: %s, path escapes EXCEL_FILES_PATH", filename)}
	}

	return candidate, nil
}

// resolvedPathIsWithin ports server._resolved_path_is_within. It re-resolves
// both arguments, as Python does, so it is safe to call on unresolved input.
//
// Containment is decided by commonpath, not by string prefix: "/tmp/ab" must
// not count as inside "/tmp/a".
func resolvedPathIsWithin(base, candidate string) bool {
	base = pyRealpath(base)
	candidate = pyRealpath(candidate)
	if candidate == base {
		return true
	}
	common, err := pyCommonpath(base, candidate)
	if err != nil {
		// os.path.commonpath raises ValueError on a mix of absolute and
		// relative paths; server.py turns that into False.
		return false
	}
	return common == base
}

// The helpers below reproduce CPython's posixpath. The Go stdlib equivalents
// are close but not equal in the places that matter here:
//
//   - filepath.EvalSymlinks fails on a path whose last components do not exist;
//     posixpath.realpath (strict=False) resolves what it can and keeps the rest
//     lexically. get_excel_path resolves paths for files it is about to create,
//     so the non-strict behavior is required, not incidental.
//   - filepath.Clean collapses a leading "//" to "/"; normpath preserves
//     exactly two leading slashes, which POSIX reserves as implementation
//     defined.
//
// They are written against posix semantics on every GOOS. The Python server
// would use ntpath on Windows; reproducing that is out of scope (SPEC 5.7
// specifies the posix behavior), and a fixed rule is easier to reason about
// than one that changes under the port.

// pyIsabs ports posixpath.isabs.
func pyIsabs(p string) bool { return strings.HasPrefix(p, "/") }

// pyJoin ports posixpath.join for two arguments.
func pyJoin(a, b string) string {
	switch {
	case strings.HasPrefix(b, "/"):
		return b
	case a == "" || strings.HasSuffix(a, "/"):
		return a + b
	default:
		return a + "/" + b
	}
}

// pySplit ports posixpath.split.
func pySplit(p string) (head, tail string) {
	i := strings.LastIndex(p, "/") + 1
	head, tail = p[:i], p[i:]
	if head != "" && strings.Trim(head, "/") != "" {
		head = strings.TrimRight(head, "/")
	}
	return head, tail
}

// pyNormpath ports posixpath.normpath: purely lexical, no filesystem access.
func pyNormpath(path string) string {
	if path == "" {
		return "."
	}
	initialSlashes := 0
	if strings.HasPrefix(path, "/") {
		initialSlashes = 1
		// POSIX leaves a path beginning with exactly two slashes
		// implementation defined, so normpath preserves it; three or more
		// collapse to one.
		if strings.HasPrefix(path, "//") && !strings.HasPrefix(path, "///") {
			initialSlashes = 2
		}
	}
	var out []string
	for _, comp := range strings.Split(path, "/") {
		if comp == "" || comp == "." {
			continue
		}
		if comp != ".." ||
			(initialSlashes == 0 && len(out) == 0) ||
			(len(out) > 0 && out[len(out)-1] == "..") {
			out = append(out, comp)
		} else if len(out) > 0 {
			out = out[:len(out)-1]
		}
	}
	res := strings.Repeat("/", initialSlashes) + strings.Join(out, "/")
	if res == "" {
		return "."
	}
	return res
}

// pyAbspath ports posixpath.abspath.
func pyAbspath(p string) string {
	if !pyIsabs(p) {
		if cwd, err := os.Getwd(); err == nil {
			p = pyJoin(cwd, p)
		}
	}
	return pyNormpath(p)
}

// pyRealpath ports posixpath.realpath(path, strict=False).
func pyRealpath(path string) string {
	resolved, _ := joinRealpath("", path, map[string]*string{})
	return pyAbspath(resolved)
}

// joinRealpath ports posixpath._joinrealpath.
//
// seen maps an already-visited symlink to its resolution; a nil value marks one
// currently being resolved, which is how a symlink loop is detected. On a loop
// the non-strict path gives up and returns the remainder lexically, exactly as
// CPython does, rather than erroring.
func joinRealpath(path, rest string, seen map[string]*string) (string, bool) {
	if strings.HasPrefix(rest, "/") {
		rest = rest[1:]
		path = "/"
	}

	for rest != "" {
		var name string
		if i := strings.Index(rest, "/"); i >= 0 {
			name, rest = rest[:i], rest[i+1:]
		} else {
			name, rest = rest, ""
		}

		switch name {
		case "", ".":
			continue
		case "..":
			// ".." pops the resolved-so-far path, not the input, so a ".."
			// after a symlinked directory lands in the *target's* parent.
			if path != "" {
				var last string
				path, last = pySplit(path)
				if last == ".." {
					path = pyJoin(pyJoin(path, ".."), "..")
				}
			} else {
				path = ".."
			}
			continue
		}

		newpath := pyJoin(path, name)
		fi, err := os.Lstat(newpath)
		if err != nil || fi.Mode()&os.ModeSymlink == 0 {
			// Non-strict: a component that does not exist is kept as-is.
			path = newpath
			continue
		}

		if prev, ok := seen[newpath]; ok {
			if prev != nil {
				path = *prev
				continue
			}
			return pyJoin(newpath, rest), false
		}

		target, err := os.Readlink(newpath)
		if err != nil {
			// Raced with the Lstat above; treat it as a plain component.
			path = newpath
			continue
		}

		seen[newpath] = nil
		var ok bool
		path, ok = joinRealpath(path, target, seen)
		if !ok {
			return pyJoin(path, rest), false
		}
		resolved := path
		seen[newpath] = &resolved
	}

	return path, true
}

// errCommonpath stands in for the ValueError os.path.commonpath raises.
var errCommonpath = errors.New("commonpath: paths are not comparable")

// pyCommonpath ports os.path.commonpath for two paths.
func pyCommonpath(a, b string) (string, error) {
	if a == "" || b == "" {
		return "", errCommonpath
	}
	if pyIsabs(a) != pyIsabs(b) {
		return "", errCommonpath
	}

	split := func(p string) []string {
		var out []string
		for _, c := range strings.Split(p, "/") {
			if c != "" && c != "." {
				out = append(out, c)
			}
		}
		return out
	}

	as, bs := split(a), split(b)
	n := 0
	for n < len(as) && n < len(bs) && as[n] == bs[n] {
		n++
	}

	prefix := ""
	if pyIsabs(a) {
		prefix = "/"
	}
	return prefix + strings.Join(as[:n], "/"), nil
}
