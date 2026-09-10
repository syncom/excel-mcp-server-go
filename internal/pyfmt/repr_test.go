package pyfmt

import "testing"

// The expectations below are CPython's own repr() output, produced by running
// the same inputs through python3 and diffed against this implementation.
func TestReprFloat(t *testing.T) {
	cases := []struct {
		in   float64
		want string
	}{
		{0, "0.0"},
		{3, "3.0"},
		{-3, "-3.0"},
		{0.5, "0.5"},
		{1757012345.6789234, "1757012345.6789234"},
		{1e16, "1e+16"},
		{1e17, "1e+17"},
		{1234567890123456.0, "1234567890123456.0"},
		{12345678901234567.0, "1.2345678901234568e+16"},
		{0.0001, "0.0001"},
		{0.00001, "1e-05"},
		{1.5e-7, "1.5e-07"},
		{-2.5e-10, "-2.5e-10"},
		{1e100, "1e+100"},
	}
	// Go folds constant arithmetic at exact precision, so 0.1+0.2 must be
	// computed at run time to reproduce the float that CPython sees.
	a, b := 0.1, 0.2
	cases = append(cases, struct {
		in   float64
		want string
	}{a + b, "0.30000000000000004"})
	for _, tc := range cases {
		t.Run(tc.want, func(t *testing.T) {
			if got := ReprFloat(tc.in); got != tc.want {
				t.Fatalf("ReprFloat(%v) = %q, want %q", tc.in, got, tc.want)
			}
		})
	}
}

func TestReprString(t *testing.T) {
	cases := []struct{ in, want string }{
		{"abc", "'abc'"},
		{"", "''"},
		{"it's", `"it's"`},
		{`say "hi"`, `'say "hi"'`},
		{`both ' and "`, `'both \' and "'`},
		{"back\\slash", `'back\\slash'`},
		{"tab\there", `'tab\there'`},
		{"nl\n", `'nl\n'`},
		{"café", "'café'"},
		{"\x00", `'\x00'`},
		// Caught by a differential run against CPython: these look printable
		// but str.isprintable() rejects them, so repr() escapes them.
		{"a\u00a0b", `'a\xa0b'`},
		{"\u2028\u2029", `'\u2028\u2029'`},
		{"\ufeff", `'\ufeff'`},
		{"naïve → ✓", "'naïve → ✓'"},
		{"日本語", "'日本語'"},
	}
	for _, tc := range cases {
		t.Run(tc.want, func(t *testing.T) {
			if got := ReprString(tc.in); got != tc.want {
				t.Fatalf("ReprString(%q) = %q, want %q", tc.in, got, tc.want)
			}
		})
	}
}

func TestReprContainers(t *testing.T) {
	if got := Repr(NewDict()); got != "{}" {
		t.Fatalf("empty dict = %q", got)
	}
	if got := Repr([]string{}); got != "[]" {
		t.Fatalf("empty list = %q", got)
	}
	d := NewDict().Set("filename", "b.xlsx").Set("sheets", []string{"Sheet1"}).
		Set("size", 4917).Set("modified", 1757012345.6789234)
	want := "{'filename': 'b.xlsx', 'sheets': ['Sheet1'], 'size': 4917, 'modified': 1757012345.6789234}"
	if got := Repr(d); got != want {
		t.Fatalf("dict repr =\n  %s\nwant\n  %s", got, want)
	}
	nested := NewDict().Set("used_ranges", NewDict().Set("Sheet1", "A1:A1"))
	if got := Repr(nested); got != "{'used_ranges': {'Sheet1': 'A1:A1'}}" {
		t.Fatalf("nested dict = %q", got)
	}
	if got := Repr([]any{true, false, nil, 1, "x"}); got != "[True, False, None, 1, 'x']" {
		t.Fatalf("mixed list = %q", got)
	}
}
