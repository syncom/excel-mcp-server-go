package excelops

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// UsedRange reports the worksheet's extent with openpyxl's semantics, which is
// what every SPEC 7 comparison and every used-range message is defined against.
//
// openpyxl's max_row/max_column count any cell that exists in the sheet XML,
// including one that is empty but styled, and never fall below 1 — a brand-new
// sheet reports 1x1, not 0x0. excelize's GetRows trims trailing empty rows and
// returns nothing at all for an empty sheet, so it cannot be used directly
// (STANDARDS: excelize usage).
func UsedRange(f *excelize.File, sheet string) (maxRow, maxCol int, err error) {
	rows, err := f.GetRows(sheet, excelize.Options{RawCellValue: true})
	if err != nil {
		return 0, 0, err
	}
	for i, row := range rows {
		for j, cell := range row {
			if cell != "" {
				if i+1 > maxRow {
					maxRow = i + 1
				}
				if j+1 > maxCol {
					maxCol = j + 1
				}
			}
		}
	}

	// A styled-but-empty cell still counts for openpyxl. GetRows drops it, so
	// consult the sheet's declared dimension too and take the larger extent.
	if dim, derr := f.GetSheetDimension(sheet); derr == nil && dim != "" {
		if _, _, dr, dc, perr := parseRangeBounds(dim); perr == nil {
			if dr > maxRow {
				maxRow = dr
			}
			if dc > maxCol {
				maxCol = dc
			}
		}
	}

	// openpyxl never reports zero.
	if maxRow < 1 {
		maxRow = 1
	}
	if maxCol < 1 {
		maxCol = 1
	}
	return maxRow, maxCol, nil
}

// parseRangeBounds splits "A1:C4" (or a bare "A1") into 1-based bounds.
func parseRangeBounds(ref string) (startRow, startCol, endRow, endCol int, err error) {
	start, end := ref, ""
	for i := 0; i < len(ref); i++ {
		if ref[i] == ':' {
			start, end = ref[:i], ref[i+1:]
			break
		}
	}
	sc, sr, err := excelize.CellNameToCoordinates(start)
	if err != nil {
		return 0, 0, 0, 0, err
	}
	if end == "" {
		return sr, sc, sr, sc, nil
	}
	ec, er, err := excelize.CellNameToCoordinates(end)
	if err != nil {
		return 0, 0, 0, 0, err
	}
	return sr, sc, er, ec, nil
}

// cellRefRe mirrors cell_utils.parse_cell_range's regexp, including its use of
// re.match rather than re.fullmatch.
//
// PARITY(cell_utils.py:parse_cell_range): re.match anchors only at the start,
// so "A1:C3" parses as "A1" and the rest is silently dropped. Callers depend on
// that looseness — it is what makes read_excel_range_with_metadata accept a
// range in its start_cell argument, and it is the mechanism deviation D2
// narrows in exactly one place without changing it here (STANDARDS: parity
// discipline).
var cellRefRe = regexp.MustCompile(`^([A-Z]+)([0-9]+)`)

// ParseCellRange ports cell_utils.parse_cell_range. endRef may be empty, in
// which case endRow and endCol come back as zero, standing in for Python's None.
func ParseCellRange(cellRef, endRef string) (startRow, startCol, endRow, endCol int, err error) {
	startRow, startCol, err = parseOneCellRef(cellRef)
	if err != nil {
		return 0, 0, 0, 0, err
	}
	if endRef == "" {
		return startRow, startCol, 0, 0, nil
	}
	endRow, endCol, err = parseOneCellRef(endRef)
	if err != nil {
		return 0, 0, 0, 0, err
	}
	return startRow, startCol, endRow, endCol, nil
}

func parseOneCellRef(ref string) (row, col int, err error) {
	m := cellRefRe.FindStringSubmatch(strings.ToUpper(ref))
	if m == nil {
		return 0, 0, fmt.Errorf("Invalid cell reference: %s", ref)
	}
	row, err = strconv.Atoi(m[2])
	if err != nil {
		return 0, 0, fmt.Errorf("Invalid cell reference: %s", ref)
	}
	col, err = pyColumnIndex(m[1])
	if err != nil {
		return 0, 0, err
	}
	return row, col, nil
}

// usedRangeOrigin reports openpyxl's min_row/min_column: the first row and
// column that carry a cell. openpyxl never reports zero, and reports 1 for an
// empty sheet.
func usedRangeOrigin(f *excelize.File, sheet string) (minRow, minCol int, err error) {
	rows, err := f.GetRows(sheet, excelize.Options{RawCellValue: true})
	if err != nil {
		return 0, 0, err
	}
	minRow, minCol = 0, 0
	for i, row := range rows {
		for j, cell := range row {
			if cell == "" {
				continue
			}
			if minRow == 0 || i+1 < minRow {
				minRow = i + 1
			}
			if minCol == 0 || j+1 < minCol {
				minCol = j + 1
			}
		}
	}
	if minRow < 1 {
		minRow = 1
	}
	if minCol < 1 {
		minCol = 1
	}
	return minRow, minCol, nil
}

// CellValue reads one cell with openpyxl's typing rules.
//
// openpyxl decides int vs float with _cast_number — a numeric string containing
// '.', 'e' or 'E' becomes a float, everything else an int — and turns a
// date-formatted number into a datetime whose str() is what json.dumps(...,
// default=str) emits.
func CellValue(f *excelize.File, sheet, addr string) (any, error) {
	raw, err := f.GetCellValue(sheet, addr, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, err
	}
	if raw == "" {
		return nil, nil
	}

	typ, err := f.GetCellType(sheet, addr)
	if err != nil {
		return nil, err
	}

	switch typ {
	case excelize.CellTypeBool:
		return raw == "1" || strings.EqualFold(raw, "true"), nil
	case excelize.CellTypeDate:
		t, err := f.GetCellValue(sheet, addr)
		if err != nil {
			return nil, err
		}
		return pyDateTimeString(t), nil
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		// openpyxl turns a numeric cell into a datetime when its *number
		// format* is a date format, and json.dumps(default=str) then renders
		// str(datetime). excelize reports CellTypeDate only for t="d" cells, so
		// the format has to be inspected directly.
		if isDate, err := cellHasDateFormat(f, sheet, addr); err == nil && isDate {
			if serial, perr := strconv.ParseFloat(raw, 64); perr == nil {
				return excelSerialToPyString(serial), nil
			}
		}
		// A numeric cell carries no `t` attribute in OOXML, which excelize
		// reports as CellTypeUnset rather than CellTypeNumber. castNumber falls
		// back to the raw string if it does not parse.
		return castNumber(raw), nil
	default:
		return raw, nil
	}
}

// castNumber ports openpyxl's _cast_number.
func castNumber(raw string) any {
	if strings.ContainsAny(raw, ".eE") {
		if fl, err := strconv.ParseFloat(raw, 64); err == nil {
			return fl
		}
		return raw
	}
	if i, err := strconv.ParseInt(raw, 10, 64); err == nil {
		return int(i)
	}
	// Python ints are arbitrary precision, so openpyxl round-trips a 21-digit
	// integer exactly while float64 would round it to 1e+20. Keep the digits.
	if isDecimalInteger(raw) {
		return pyfmt.RawNumber(raw)
	}
	if fl, err := strconv.ParseFloat(raw, 64); err == nil {
		return fl
	}
	return raw
}

// isDecimalInteger reports whether raw is a plain optionally-signed run of
// digits, i.e. something Python's int() would accept exactly.
func isDecimalInteger(raw string) bool {
	body := strings.TrimPrefix(strings.TrimPrefix(raw, "-"), "+")
	if body == "" {
		return false
	}
	for i := 0; i < len(body); i++ {
		if body[i] < '0' || body[i] > '9' {
			return false
		}
	}
	return true
}

// pyDateTimeString renders a datetime the way json.dumps(default=str) does,
// i.e. as str(datetime): "2024-03-15 00:00:00", dropping a zero microsecond
// component.
func pyDateTimeString(formatted string) string {
	t, err := time.Parse("01-02-06", formatted)
	if err == nil {
		return t.Format("2006-01-02 15:04:05")
	}
	for _, layout := range []string{
		"2006-01-02 15:04:05",
		"2006-01-02T15:04:05Z07:00",
		"2006-01-02",
		"1/2/2006",
		"01/02/2006",
	} {
		if t, err := time.Parse(layout, formatted); err == nil {
			return t.Format("2006-01-02 15:04:05")
		}
	}
	return formatted
}

// splitSqref splits a space-separated sqref such as "A1:A5 C1:C5".
func splitSqref(sqref string) []string {
	return strings.Fields(strings.ReplaceAll(sqref, "$", ""))
}

// dateFormatLiteral strips the parts of a number-format string that cannot
// carry a date token: quoted literals, [bracketed] colour and condition
// sections, and backslash-escaped characters.
var dateFormatLiteral = regexp.MustCompile(`\[[^\]]*\]|"[^"]*"|\\.`)

// builtinDateFormats are the built-in numFmt ids openpyxl treats as dates.
var builtinDateFormats = map[int]bool{
	14: true, 15: true, 16: true, 17: true, 18: true, 19: true, 20: true,
	21: true, 22: true, 45: true, 46: true, 47: true,
}

// cellHasDateFormat ports openpyxl's is_date_format check for one cell.
func cellHasDateFormat(f *excelize.File, sheet, addr string) (bool, error) {
	styleID, err := f.GetCellStyle(sheet, addr)
	if err != nil {
		return false, err
	}
	if styleID == 0 {
		return false, nil
	}
	style, err := f.GetStyle(styleID)
	if err != nil || style == nil {
		return false, err
	}
	if style.CustomNumFmt != nil {
		return isDateFormat(*style.CustomNumFmt), nil
	}
	return builtinDateFormats[style.NumFmt], nil
}

// isDateFormat reports whether a number-format string denotes a date or time.
func isDateFormat(format string) bool {
	if format == "" || format == "General" {
		return false
	}
	// Only the first (positive) section decides.
	section := strings.SplitN(format, ";", 2)[0]
	section = dateFormatLiteral.ReplaceAllString(section, "")
	return strings.ContainsAny(section, "dmyhsDMYHS")
}

// excelSerialToPyString renders an Excel date serial the way
// json.dumps(..., default=str) renders what openpyxl returns: str(datetime),
// or str(time) for a serial with no date part.
func excelSerialToPyString(serial float64) string {
	t, err := excelize.ExcelDateToTime(serial, false)
	if err != nil {
		return pyfmt.FormatGDefault(serial)
	}
	if serial < 1 {
		return t.Format("15:04:05")
	}
	if t.Nanosecond() != 0 {
		return t.Format("2006-01-02 15:04:05.000000")
	}
	return t.Format("2006-01-02 15:04:05")
}

// pySplit2 reproduces Python's `a, b = s.split(sep)`.
//
// The tuple unpack requires exactly two parts, so a string with no separator or
// with more than one raises ValueError. strings.Cut never fails, which would
// silently accept input the oracle rejects.
func pySplit2(s, sep string) (string, string, error) {
	parts := strings.Split(s, sep)
	switch {
	case len(parts) == 2:
		return parts[0], parts[1], nil
	case len(parts) < 2:
		return "", "", fmt.Errorf("not enough values to unpack (expected 2, got %d)", len(parts))
	default:
		return "", "", fmt.Errorf("too many values to unpack (expected 2, got %d)", len(parts))
	}
}

// pyColumnIndex ports openpyxl's column_index_from_string, which accepts one to
// three letters (A-ZZZ) and has its own message beyond that. excelize's
// ColumnNameToNumber caps at XFD (16384) and reports a different error, so it
// cannot stand in here.
func pyColumnIndex(letters string) (int, error) {
	up := strings.ToUpper(letters)
	if len(up) < 1 || len(up) > 3 {
		return 0, fmt.Errorf("'%s' is not a valid column name. Column names are from A to ZZZ", up)
	}
	col := 0
	for i := 0; i < len(up); i++ {
		c := up[i]
		if c < 'A' || c > 'Z' {
			return 0, fmt.Errorf("'%s' is not a valid column name. Column names are from A to ZZZ", up)
		}
		col = col*26 + int(c-'A'+1)
	}
	return col, nil
}
