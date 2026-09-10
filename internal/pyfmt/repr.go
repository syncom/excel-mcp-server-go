package pyfmt

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode"
)

// Dict is an insertion-ordered mapping.
//
// CPython's dict repr preserves insertion order, and get_workbook_metadata
// depends on it: SPEC 5.4 pins the key order as filename, sheets, size,
// modified, then used_ranges. A Go map would randomize it.
type Dict struct {
	keys []string
	vals []any
}

// NewDict returns an empty ordered dict.
func NewDict() *Dict { return &Dict{} }

// Set appends a key. A repeated key overwrites in place, keeping its original
// position, which is what CPython does.
func (d *Dict) Set(key string, val any) *Dict {
	for i, k := range d.keys {
		if k == key {
			d.vals[i] = val
			return d
		}
	}
	d.keys = append(d.keys, key)
	d.vals = append(d.vals, val)
	return d
}

// Len reports the number of keys.
func (d *Dict) Len() int { return len(d.keys) }

// Repr renders a value the way CPython's built-in repr() would (SPEC 5.4).
//
// Only the shapes the tool results actually contain are supported: dict, list,
// str, bool, None, float, int. Anything else is a programming error and says so
// loudly rather than silently emitting Go's formatting.
func Repr(v any) string {
	switch x := v.(type) {
	case nil:
		return "None"
	case bool:
		if x {
			return "True"
		}
		return "False"
	case string:
		return ReprString(x)
	case int:
		return strconv.Itoa(x)
	case int64:
		return strconv.FormatInt(x, 10)
	case float64:
		return ReprFloat(x)
	case RawNumber:
		return string(x)
	case *Dict:
		return reprDict(x)
	case []string:
		parts := make([]string, len(x))
		for i, s := range x {
			parts[i] = ReprString(s)
		}
		return "[" + strings.Join(parts, ", ") + "]"
	case []any:
		parts := make([]string, len(x))
		for i, e := range x {
			parts[i] = Repr(e)
		}
		return "[" + strings.Join(parts, ", ") + "]"
	default:
		panic(fmt.Sprintf("pyfmt.Repr: unsupported type %T", v))
	}
}

func reprDict(d *Dict) string {
	if d.Len() == 0 {
		return "{}"
	}
	var b strings.Builder
	b.WriteByte('{')
	for i, k := range d.keys {
		if i > 0 {
			b.WriteString(", ")
		}
		b.WriteString(ReprString(k))
		b.WriteString(": ")
		b.WriteString(Repr(d.vals[i]))
	}
	b.WriteByte('}')
	return b.String()
}

// ReprString renders a Python str repr.
//
// CPython prefers single quotes and switches to double quotes only when the
// value contains a single quote and no double quote.
func ReprString(s string) string {
	quote := byte('\'')
	if strings.ContainsRune(s, '\'') && !strings.ContainsRune(s, '"') {
		quote = '"'
	}

	var b strings.Builder
	b.WriteByte(quote)
	for _, r := range s {
		switch {
		case r == rune(quote) || r == '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		case r == '\n':
			b.WriteString(`\n`)
		case r == '\r':
			b.WriteString(`\r`)
		case r == '\t':
			b.WriteString(`\t`)
		case !unicode.IsPrint(r):
			// CPython escapes every character str.isprintable() rejects — not
			// just ASCII control codes. U+00A0 NO-BREAK SPACE and U+2028 LINE
			// SEPARATOR are printable-looking but escaped, and missing that was
			// the only string divergence the CPython corpus turned up.
			switch {
			case r < 0x100:
				fmt.Fprintf(&b, `\x%02x`, r)
			case r < 0x10000:
				fmt.Fprintf(&b, `\u%04x`, r)
			default:
				fmt.Fprintf(&b, `\U%08x`, r)
			}
		default:
			// Printable non-ASCII stays raw in Python 3.
			b.WriteRune(r)
		}
	}
	b.WriteByte(quote)
	return b.String()
}

// ReprFloat renders a Python float repr: the shortest string that round-trips,
// in exponential form only when CPython would use it.
//
// CPython formats with _Py_dg_dtoa mode 0 and then switches to exponential
// notation when the decimal point position is <= -4 or > 16 — not the rule Go's
// %g uses, which is why this cannot be a single FormatFloat call.
func ReprFloat(f float64) string {
	switch {
	case math.IsNaN(f):
		return "nan"
	case math.IsInf(f, 1):
		return "inf"
	case math.IsInf(f, -1):
		return "-inf"
	}

	// Shortest round-tripping digits, as mantissa and exponent.
	sci := strconv.FormatFloat(f, 'e', -1, 64)
	mantissa, expPart, _ := strings.Cut(sci, "e")
	exp, err := strconv.Atoi(expPart)
	if err != nil {
		return sci
	}

	neg := strings.HasPrefix(mantissa, "-")
	mantissa = strings.TrimPrefix(mantissa, "-")
	digits := strings.Replace(mantissa, ".", "", 1)

	// decpt is the position of the decimal point relative to the digit string.
	decpt := exp + 1

	var out string
	if decpt <= -4 || decpt > 16 {
		out = expForm(digits, decpt)
	} else {
		out = fixedForm(digits, decpt)
	}
	if neg {
		out = "-" + out
	}
	return out
}

func expForm(digits string, decpt int) string {
	// Py_DTSF_ADD_DOT_0 is not applied in exponential form; CPython writes
	// "1e+20", not "1.0e+20".
	var b strings.Builder
	b.WriteByte(digits[0])
	if len(digits) > 1 {
		b.WriteByte('.')
		b.WriteString(digits[1:])
	}
	e := decpt - 1
	sign := "+"
	if e < 0 {
		sign = "-"
		e = -e
	}
	fmt.Fprintf(&b, "e%s%02d", sign, e)
	return b.String()
}

func fixedForm(digits string, decpt int) string {
	switch {
	case decpt <= 0:
		// 0.000ddd
		return "0." + strings.Repeat("0", -decpt) + digits
	case decpt >= len(digits):
		// Integral value: CPython appends ".0" (Py_DTSF_ADD_DOT_0).
		return digits + strings.Repeat("0", decpt-len(digits)) + ".0"
	default:
		return digits[:decpt] + "." + digits[decpt:]
	}
}

// StrOf renders a value the way Python's str() would, which is how
// _extract_list_values stringifies a resolved cell and how json.dumps's
// default=str handles anything it cannot serialize.
//
// str() differs from repr() for strings — no quotes — and agrees with it for
// the numeric and boolean shapes involved here.
func StrOf(v any) string {
	if s, ok := v.(string); ok {
		return s
	}
	return Repr(v)
}

// Get returns the value stored under key.
func (d *Dict) Get(key string) (any, bool) {
	for i, k := range d.keys {
		if k == key {
			return d.vals[i], true
		}
	}
	return nil, false
}

// FormatG reproduces C's "%.*g", which is how openpyxl's safe_string writes a
// numeric cell value.
//
// It is not strconv's 'g': with an explicit precision Go keeps trailing zeros,
// while C strips them, and the two disagree about when to switch to
// exponential form. Getting this wrong is not cosmetic — the written string is
// what the reader casts back, so "10.00000000000000" would come back as a
// float where Python gets the integer 10.
func FormatG(f float64, prec int) string {
	switch {
	case math.IsNaN(f):
		return "nan"
	case math.IsInf(f, 1):
		return "inf"
	case math.IsInf(f, -1):
		return "-inf"
	}
	if prec < 1 {
		prec = 1
	}

	sci := strconv.FormatFloat(f, 'e', prec-1, 64)
	mantissa, expPart, _ := strings.Cut(sci, "e")
	exp, err := strconv.Atoi(expPart)
	if err != nil {
		return sci
	}

	// C uses exponential form when the exponent is below -4 or at least the
	// precision; fixed form otherwise, with precision-1-exp fraction digits.
	if exp < -4 || exp >= prec {
		sign := "+"
		e := exp
		if e < 0 {
			sign = "-"
			e = -e
		}
		return fmt.Sprintf("%se%s%02d", trimTrailingZeros(mantissa), sign, e)
	}
	return trimTrailingZeros(strconv.FormatFloat(f, 'f', prec-1-exp, 64))
}

func trimTrailingZeros(s string) string {
	if !strings.Contains(s, ".") {
		return s
	}
	s = strings.TrimRight(s, "0")
	return strings.TrimSuffix(s, ".")
}

// FormatGDefault is FormatG at the precision openpyxl's safe_string uses.
func FormatGDefault(f float64) string { return FormatG(f, 16) }
