package pyfmt

import "testing"

// Expectations are CPython's own json.dumps output for the same inputs,
// generated from the oracle rather than typed by hand.
func TestJSONString(t *testing.T) {
	cases := []struct{ in, want string }{
		{in: "abc", want: "\"abc\""},
		{in: "", want: "\"\""},
		{in: "say \"hi\"", want: "\"say \\\"hi\\\"\""},
		{in: "back\\slash", want: "\"back\\\\slash\""},
		{in: "tab\there", want: "\"tab\\there\""},
		{in: "nl\n", want: "\"nl\\n\""},
		{in: "cr\r", want: "\"cr\\r\""},
		{in: "\b\f", want: "\"\\b\\f\""},
		{in: "\u0000\u001f", want: "\"\\u0000\\u001f\""},
		{in: "<a> & </a>", want: "\"<a> & </a>\""},
		{in: "café", want: "\"caf\\u00e9\""},
		{in: "日本語", want: "\"\\u65e5\\u672c\\u8a9e\""},
		{in: "�", want: "\"\\ufffd\""},
		{in: "😀", want: "\"\\ud83d\\ude00\""},
		{in: "𝔘", want: "\"\\ud835\\udd18\""},
	}
	for _, tc := range cases {
		t.Run(tc.want, func(t *testing.T) {
			if got := JSONString(tc.in); got != tc.want {
				t.Fatalf("JSONString(%q) = %s, want %s", tc.in, got, tc.want)
			}
		})
	}
}

func TestDumps(t *testing.T) {
	t.Run("empty_containers_stay_on_one_line", func(t *testing.T) {
		if got := Dumps(NewDict()); got != "{}" {
			t.Fatalf("empty dict = %q", got)
		}
		if got := Dumps([]any{}); got != "[]" {
			t.Fatalf("empty list = %q", got)
		}
	})

	t.Run("scalars", func(t *testing.T) {
		for _, tc := range []struct {
			in   any
			want string
		}{
			{nil, "null"}, {true, "true"}, {false, "false"},
			{1, "1"}, {2.5, "2.5"}, {3.0, "3.0"},
			{RawNumber("100000000000000000000"), "100000000000000000000"},
		} {
			if got := Dumps(tc.in); got != tc.want {
				t.Fatalf("Dumps(%v) = %q, want %q", tc.in, got, tc.want)
			}
		}
	})

	t.Run("nesting_and_key_order", func(t *testing.T) {
		inner := NewDict().Set("has_validation", false)
		d := NewDict().
			Set("range", "A1:A1").
			Set("cells", []any{NewDict().Set("address", "A1").Set("validation", inner)})
		want := "{\n  \"range\": \"A1:A1\",\n  \"cells\": [\n    {\n      \"address\": \"A1\",\n      \"validation\": {\n        \"has_validation\": false\n      }\n    }\n  ]\n}"
		if got := Dumps(d); got != want {
			t.Fatalf("Dumps =\n%s\nwant\n%s", got, want)
		}
	})
}
