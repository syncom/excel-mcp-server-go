package excelops

import "testing"

// TestNormalizeColor covers deviation D5's full input space.
//
// Two of these cannot be conformance cases. D5's registry entry pins one Go
// expectation per case id and its `cases` list may not grow beyond what the
// entry pins, so the '#'-prefixed and lower-case inputs — where Python errors
// and Go succeeds, but with a third and fourth distinct pair — are covered
// here instead. See JOURNAL (T7).
func TestNormalizeColor(t *testing.T) {
	cases := []struct {
		in      string
		want    string
		wantErr bool
	}{
		{in: "FF0000", want: "FFFF0000"},
		{in: "00FF00", want: "FF00FF00"},
		{in: "FFC7CE", want: "FFFFC7CE"}, // 6 digits already starting with FF
		{in: "80FF0000", want: "80FF0000"},
		{in: "FFFF0000", want: "FFFF0000"},
		{in: "#FF0000", want: "FFFF0000"},
		{in: "#80FF0000", want: "80FF0000"},
		{in: "ff0000", want: "FFFF0000"},
		{in: "ZZZZZZ", wantErr: true},
		{in: "FF00", wantErr: true},
		{in: "FF0000FF00", wantErr: true},
		{in: "", wantErr: true},
		{in: "FF00 0", wantErr: true},
	}
	for _, tc := range cases {
		name := tc.in
		if name == "" {
			name = "empty"
		}
		t.Run(name, func(t *testing.T) {
			got, err := NormalizeColor(tc.in)
			if tc.wantErr {
				if err == nil {
					t.Fatalf("NormalizeColor(%q) = %q, want an error", tc.in, got)
				}
				want := "Invalid color: " + tc.in
				if err.Error() != want {
					t.Fatalf("NormalizeColor(%q) error = %q, want %q", tc.in, err, want)
				}
				return
			}
			if err != nil {
				t.Fatalf("NormalizeColor(%q) failed: %v", tc.in, err)
			}
			if got != tc.want {
				t.Fatalf("NormalizeColor(%q) = %q, want %q", tc.in, got, tc.want)
			}
			// excelize is handed the RGB triplet; it prepends its own alpha.
			if e := excelizeColor(got); len(e) != 6 {
				t.Fatalf("excelizeColor(%q) = %q, want 6 hex digits", got, e)
			}
		})
	}
}

func TestIsDateFormat(t *testing.T) {
	cases := []struct {
		in   string
		want bool
	}{
		{"yyyy-mm-dd", true},
		{"General", false},
		{"", false},
		{"0.00", false},
		{"h:mm:ss", true},
		{`"total: "0.00`, false}, // the d/m/y in a quoted literal do not count
		{`[Red]0.00;[Blue]-0.00`, false},
		{`mm/dd/yy;@`, true},
		{`#,##0`, false},
	}
	for _, tc := range cases {
		t.Run(tc.in, func(t *testing.T) {
			if got := isDateFormat(tc.in); got != tc.want {
				t.Fatalf("isDateFormat(%q) = %v, want %v", tc.in, got, tc.want)
			}
		})
	}
}
