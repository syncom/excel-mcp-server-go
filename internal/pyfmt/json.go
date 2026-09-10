package pyfmt

import (
	"fmt"
	"strconv"
	"strings"
	"unicode/utf16"
)

// Dumps reproduces json.dumps(obj, indent=2, default=str) byte for byte
// (SPEC 5.5). Go's encoding/json disagrees with CPython's in four ways that all
// matter here:
//
//   - Go escapes <, > and & by default; Python never does.
//   - Python's ensure_ascii=True escapes every non-ASCII rune as \uXXXX, with
//     surrogate pairs for astral characters; Go emits raw UTF-8.
//   - Go's MarshalIndent sorts map keys; Python preserves insertion order.
//   - Go renders floats with its own shortest-form rules; Python's json encoder
//     uses float.__repr__, which is [ReprFloat].
//
// Supported values are the ones the tool results contain: *Dict, []any, string,
// int, float64, bool and nil. Anything else is a programming error.
func Dumps(v any) string {
	var b strings.Builder
	writeJSON(&b, v, 0)
	return b.String()
}

const jsonIndent = "  "

func writeJSON(b *strings.Builder, v any, depth int) {
	switch x := v.(type) {
	case nil:
		b.WriteString("null")
	case bool:
		if x {
			b.WriteString("true")
		} else {
			b.WriteString("false")
		}
	case string:
		b.WriteString(JSONString(x))
	case int:
		b.WriteString(strconv.Itoa(x))
	case int64:
		b.WriteString(strconv.FormatInt(x, 10))
	case float64:
		// Python's json encoder formats floats with float.__repr__.
		b.WriteString(ReprFloat(x))
	case RawNumber:
		b.WriteString(string(x))
	case *Dict:
		writeDict(b, x, depth)
	case []string:
		items := make([]any, len(x))
		for i, s := range x {
			items[i] = s
		}
		writeList(b, items, depth)
	case []any:
		writeList(b, x, depth)
	default:
		panic(fmt.Sprintf("pyfmt.Dumps: unsupported type %T", v))
	}
}

func writeDict(b *strings.Builder, d *Dict, depth int) {
	// CPython renders an empty container on one line, with no inner newline.
	if d.Len() == 0 {
		b.WriteString("{}")
		return
	}
	pad := strings.Repeat(jsonIndent, depth+1)
	b.WriteString("{\n")
	for i, k := range d.keys {
		if i > 0 {
			b.WriteString(",\n")
		}
		b.WriteString(pad)
		b.WriteString(JSONString(k))
		b.WriteString(": ")
		writeJSON(b, d.vals[i], depth+1)
	}
	b.WriteByte('\n')
	b.WriteString(strings.Repeat(jsonIndent, depth))
	b.WriteByte('}')
}

func writeList(b *strings.Builder, items []any, depth int) {
	if len(items) == 0 {
		b.WriteString("[]")
		return
	}
	pad := strings.Repeat(jsonIndent, depth+1)
	b.WriteString("[\n")
	for i, item := range items {
		if i > 0 {
			b.WriteString(",\n")
		}
		b.WriteString(pad)
		writeJSON(b, item, depth+1)
	}
	b.WriteByte('\n')
	b.WriteString(strings.Repeat(jsonIndent, depth))
	b.WriteByte(']')
}

// JSONString renders a Python json.dumps string literal with ensure_ascii=True
// and no HTML escaping.
func JSONString(s string) string {
	var b strings.Builder
	b.WriteByte('"')
	for _, r := range s {
		switch r {
		case '"':
			b.WriteString(`\"`)
		case '\\':
			b.WriteString(`\\`)
		case '\n':
			b.WriteString(`\n`)
		case '\r':
			b.WriteString(`\r`)
		case '\t':
			b.WriteString(`\t`)
		case '\b':
			b.WriteString(`\b`)
		case '\f':
			b.WriteString(`\f`)
		default:
			switch {
			case r < 0x20:
				fmt.Fprintf(&b, `\u%04x`, r)
			case r < 0x7f:
				b.WriteRune(r)
			case r < 0x10000:
				fmt.Fprintf(&b, `\u%04x`, r)
			default:
				// ensure_ascii writes astral characters as a surrogate pair.
				hi, lo := utf16.EncodeRune(r)
				fmt.Fprintf(&b, `\u%04x\u%04x`, hi, lo)
			}
		}
	}
	b.WriteByte('"')
	return b.String()
}

// RawNumber is a numeric literal carried verbatim.
//
// Python integers are arbitrary precision: openpyxl round-trips a 21-digit
// integer exactly, where float64 would round it to 1e+20. Values that do not
// fit an int64 travel as RawNumber so json.dumps parity survives them.
type RawNumber string
