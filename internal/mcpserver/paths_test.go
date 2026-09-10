package mcpserver

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestExcelPath ports every case of excel-mcp-server/tests/test_sandbox_paths.py,
// plus the symlink-escape case Python does not have. Each case builds its own
// sandbox, so they are independent of order.
func TestExcelPath(t *testing.T) {
	// setup returns the config to test under and the filename to pass. want is
	// the expected path; if wantErrSubstr is non-empty the call must fail and
	// its message must contain that substring instead.
	cases := []struct {
		name          string
		setup         func(t *testing.T) (Config, string, string)
		wantErrSubstr string
	}{
		{
			// test_stdio_accepts_absolute_only, accept half.
			name: "stdio_accepts_absolute",
			setup: func(t *testing.T) (Config, string, string) {
				// t.TempDir() is already clean and absolute, so the expected
				// result is the input itself — computing it with pyNormpath
				// would have compared the code under test against itself.
				f := filepath.Join(t.TempDir(), "book.xlsx")
				if err := os.WriteFile(f, nil, 0o644); err != nil {
					t.Fatal(err)
				}
				return Config{}, f, f
			},
		},
		{
			// test_stdio_accepts_absolute_only, reject half.
			name: "stdio_rejects_relative",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{}, "relative_only.xlsx", ""
			},
			wantErrSubstr: "must be an absolute path when not in SSE mode",
		},
		{
			// normpath is lexical: it must not consult the filesystem, and it
			// must not resolve symlinks. Nothing here exists.
			name: "stdio_normalizes_lexically_without_touching_disk",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{}, "/nonexistent/./a/../b//c.xlsx", "/nonexistent/b/c.xlsx"
			},
		},
		{
			// test_remote_rejects_absolute
			name: "remote_rejects_absolute",
			setup: func(t *testing.T) (Config, string, string) {
				d := t.TempDir()
				return Config{ExcelFilesPath: d}, filepath.Join(d, "ok.xlsx"), ""
			},
			wantErrSubstr: "must be relative to EXCEL_FILES_PATH",
		},
		{
			// test_remote_allows_relative_inside_sandbox. Note that neither
			// subdir nor file.xlsx exists: realpath must tolerate that.
			name: "remote_allows_relative_inside_sandbox",
			setup: func(t *testing.T) (Config, string, string) {
				d := t.TempDir()
				return Config{ExcelFilesPath: d}, "subdir/file.xlsx",
					filepath.Join(evalSymlinks(t, d), "subdir", "file.xlsx")
			},
		},
		{
			// test_remote_blocks_traversal, first assertion.
			name: "remote_blocks_traversal_parent",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{ExcelFilesPath: t.TempDir()}, "../outside.xlsx", ""
			},
			wantErrSubstr: "path escapes EXCEL_FILES_PATH",
		},
		{
			// test_remote_blocks_traversal, second assertion.
			name: "remote_blocks_traversal_nested",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{ExcelFilesPath: t.TempDir()}, "a/../../outside.xlsx", ""
			},
			wantErrSubstr: "path escapes EXCEL_FILES_PATH",
		},
		{
			// test_remote_rejects_nul
			name: "remote_rejects_nul",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{ExcelFilesPath: t.TempDir()}, "a\x00b.xlsx", ""
			},
			wantErrSubstr: "Invalid filename",
		},
		{
			name: "remote_rejects_empty",
			setup: func(t *testing.T) (Config, string, string) {
				return Config{ExcelFilesPath: t.TempDir()}, "", ""
			},
			wantErrSubstr: "Invalid filename",
		},
		{
			// Not in the Python suite: a symlink inside the sandbox pointing
			// out of it. Only realpath catches this; a lexical check passes it.
			name: "remote_blocks_symlink_escape",
			setup: func(t *testing.T) (Config, string, string) {
				root := t.TempDir()
				sandbox := filepath.Join(root, "sandbox")
				outside := filepath.Join(root, "outside")
				mkdirs(t, sandbox, outside)
				symlink(t, outside, filepath.Join(sandbox, "escape"))
				return Config{ExcelFilesPath: sandbox}, "escape/secret.xlsx", ""
			},
			wantErrSubstr: "path escapes EXCEL_FILES_PATH",
		},
		{
			// The complement of the case above: a symlink that stays inside is
			// allowed, and resolves to its target.
			name: "remote_allows_symlink_inside_sandbox",
			setup: func(t *testing.T) (Config, string, string) {
				sandbox := t.TempDir()
				real := filepath.Join(sandbox, "real")
				mkdirs(t, real)
				symlink(t, real, filepath.Join(sandbox, "alias"))
				return Config{ExcelFilesPath: sandbox}, "alias/book.xlsx",
					filepath.Join(evalSymlinks(t, real), "book.xlsx")
			},
		},
		{
			// Containment is decided component-wise. "/…/ab" must not count as
			// inside "/…/a" just because the string starts with it.
			name: "remote_blocks_sibling_sharing_a_name_prefix",
			setup: func(t *testing.T) (Config, string, string) {
				root := t.TempDir()
				sandbox := filepath.Join(root, "a")
				sibling := filepath.Join(root, "ab")
				mkdirs(t, sandbox, sibling)
				symlink(t, sibling, filepath.Join(sandbox, "link"))
				return Config{ExcelFilesPath: sandbox}, "link/book.xlsx", ""
			},
			wantErrSubstr: "path escapes EXCEL_FILES_PATH",
		},
		{
			// The sandbox root itself resolves to the base, which
			// _resolved_path_is_within accepts via its equality short-circuit.
			name: "remote_allows_sandbox_root_itself",
			setup: func(t *testing.T) (Config, string, string) {
				d := t.TempDir()
				return Config{ExcelFilesPath: d}, ".", evalSymlinks(t, d)
			},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			cfg, filename, want := tc.setup(t)
			s := &Server{cfg: cfg}

			got, err := s.ExcelPath(filename)

			if tc.wantErrSubstr != "" {
				if err == nil {
					t.Fatalf("ExcelPath(%q) = %q, want error containing %q", filename, got, tc.wantErrSubstr)
				}
				var ve *ValueError
				if !asValueError(err, &ve) {
					t.Fatalf("ExcelPath(%q) error is %T, want *ValueError — SPEC 5.2 routes it as a raise-through", filename, err)
				}
				if !strings.Contains(err.Error(), tc.wantErrSubstr) {
					t.Fatalf("ExcelPath(%q) error = %q, want it to contain %q", filename, err, tc.wantErrSubstr)
				}
				return
			}

			if err != nil {
				t.Fatalf("ExcelPath(%q) failed: %v", filename, err)
			}
			if got != want {
				t.Fatalf("ExcelPath(%q) = %q, want %q", filename, got, want)
			}
		})
	}
}

func asValueError(err error, target **ValueError) bool {
	ve, ok := err.(*ValueError)
	if ok {
		*target = ve
	}
	return ok
}

// evalSymlinks resolves an existing directory with the stdlib, so the
// expectations above do not route through the pyRealpath they are testing.
func evalSymlinks(t *testing.T, dir string) string {
	t.Helper()
	resolved, err := filepath.EvalSymlinks(dir)
	if err != nil {
		t.Fatal(err)
	}
	return resolved
}

func mkdirs(t *testing.T, dirs ...string) {
	t.Helper()
	for _, d := range dirs {
		if err := os.MkdirAll(d, 0o755); err != nil {
			t.Fatal(err)
		}
	}
}

func symlink(t *testing.T, target, link string) {
	t.Helper()
	if err := os.Symlink(target, link); err != nil {
		t.Fatal(err)
	}
}

// TestExcelPathHelpers pins the posixpath helpers directly. The expectations
// are CPython's own output for these inputs, so a regression in pyNormpath
// cannot hide behind a matching regression in the expectation.
func TestExcelPathHelpers(t *testing.T) {
	t.Run("normpath", func(t *testing.T) {
		cases := []struct{ in, want string }{
			{"", "."},
			{".", "."},
			{"..", ".."},
			{"/", "/"},
			// POSIX reserves exactly two leading slashes; three collapse to one.
			{"//", "//"},
			{"///", "/"},
			{"//a/b", "//a/b"},
			{"/a/b/../c", "/a/c"},
			{"a/b/../../../c", "../c"},
			{"/./a//b/./c/", "/a/b/c"},
			{"/a/../..", "/"},
			{"../..", "../.."},
			{"foo/./bar/../baz", "foo/baz"},
			{"/nonexistent/./a/../b//c.xlsx", "/nonexistent/b/c.xlsx"},
		}
		for _, tc := range cases {
			t.Run(tc.in, func(t *testing.T) {
				if got := pyNormpath(tc.in); got != tc.want {
					t.Fatalf("pyNormpath(%q) = %q, want %q", tc.in, got, tc.want)
				}
			})
		}
	})

	t.Run("commonpath", func(t *testing.T) {
		cases := []struct {
			a, b    string
			want    string
			wantErr bool
		}{
			{a: "/tmp/a", b: "/tmp/ab", want: "/tmp"},
			{a: "/tmp/a", b: "/tmp/a/b", want: "/tmp/a"},
			{a: "/tmp/a", b: "/other", want: "/"},
			{a: "/tmp/a", b: "relative", wantErr: true},
			{a: "", b: "/tmp", wantErr: true},
		}
		for _, tc := range cases {
			t.Run(tc.a+"|"+tc.b, func(t *testing.T) {
				got, err := pyCommonpath(tc.a, tc.b)
				if tc.wantErr {
					if err == nil {
						t.Fatalf("pyCommonpath(%q, %q) = %q, want error", tc.a, tc.b, got)
					}
					return
				}
				if err != nil {
					t.Fatalf("pyCommonpath(%q, %q) failed: %v", tc.a, tc.b, err)
				}
				if got != tc.want {
					t.Fatalf("pyCommonpath(%q, %q) = %q, want %q", tc.a, tc.b, got, tc.want)
				}
			})
		}
	})

	t.Run("realpath_resolves_symlinks_and_tolerates_missing_components", func(t *testing.T) {
		root := t.TempDir()
		real := filepath.Join(root, "real")
		mkdirs(t, real)
		symlink(t, real, filepath.Join(root, "alias"))
		resolved := evalSymlinks(t, real)

		// The trailing components do not exist; realpath must keep them
		// rather than fail, which is why filepath.EvalSymlinks cannot be used
		// in ExcelPath directly.
		got := pyRealpath(filepath.Join(root, "alias", "nope", "deep.xlsx"))
		want := filepath.Join(resolved, "nope", "deep.xlsx")
		if got != want {
			t.Fatalf("pyRealpath = %q, want %q", got, want)
		}
	})

	t.Run("realpath_does_not_hang_on_a_symlink_loop", func(t *testing.T) {
		root := evalSymlinks(t, t.TempDir())
		symlink(t, "loop2", filepath.Join(root, "loop1"))
		symlink(t, "loop1", filepath.Join(root, "loop2"))

		// The previous version only checked for a non-empty result, which
		// pyRealpath cannot return — pyNormpath maps "" to ".". It would have
		// passed against a stub. Pin the real output: CPython gives up on a
		// loop and keeps the path as written.
		got := pyRealpath(filepath.Join(root, "loop1"))
		if want := filepath.Join(root, "loop1"); got != want {
			t.Fatalf("pyRealpath on a loop = %q, want %q", got, want)
		}
		if !resolvedPathIsWithin(root, got) {
			t.Fatal("a looping symlink inside the sandbox reported as outside it")
		}
	})

	t.Run("symlink_loop_pointing_outward_does_not_escape", func(t *testing.T) {
		root := t.TempDir()
		sandbox := filepath.Join(root, "sandbox")
		mkdirs(t, sandbox)
		// A loop whose components point out of the sandbox must not resolve to
		// somewhere outside it, whether it is rejected or resolved.
		symlink(t, "../escape2", filepath.Join(sandbox, "escape1"))
		symlink(t, "sandbox/escape1", filepath.Join(root, "escape2"))

		s := &Server{cfg: Config{ExcelFilesPath: sandbox}}
		got, err := s.ExcelPath("escape1/book.xlsx")
		if err == nil && !resolvedPathIsWithin(sandbox, got) {
			t.Fatalf("ExcelPath returned %q, which is outside the sandbox", got)
		}
	})
}
