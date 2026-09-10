package pyfmt

import "testing"

// Expectations are CPython's own f"{v:.16g}".
func TestFormatG16(t *testing.T) {
	cases := []struct {
		in   float64
		want string
	}{
		{10, "10"}, {20.5, "20.5"}, {-3, "-3"}, {2.5, "2.5"}, {3, "3"},
		{0.1, "0.1"}, {1e-5, "1e-05"}, {1e-4, "0.0001"}, {-0.0001, "-0.0001"},
		{1e15, "1000000000000000"}, {1e16, "1e+16"}, {1e17, "1e+17"}, {1e20, "1e+20"},
		{-1e20, "-1e+20"},
		{9007199254740992, "9007199254740992"},
		{9223372036854775807, "9.223372036854776e+18"},
		{1.2345678901234568e29, "1.234567890123457e+29"},
		{0, "0"},
	}
	for _, tc := range cases {
		t.Run(tc.want, func(t *testing.T) {
			if got := FormatG(tc.in, 16); got != tc.want {
				t.Fatalf("FormatG(%v, 16) = %q, want %q", tc.in, got, tc.want)
			}
		})
	}
}
