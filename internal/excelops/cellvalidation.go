package excelops

import (
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// DataValidationForCell ports cell_validation.get_data_validation_for_cell.
//
// It returns nil when no rule covers the cell, which the caller renders as
// {"has_validation": false}. Python swallows every failure here and returns
// None as well, so a malformed rule degrades to "no validation" rather than
// failing the read.
func DataValidationForCell(f *excelize.File, sheet, addr string) *pyfmt.Dict {
	col, row, err := excelize.CellNameToCoordinates(addr)
	if err != nil {
		// PARITY(cell_validation.py:get_data_validation_for_cell): the whole
		// body is wrapped in `except Exception: return None`.
		return nil
	}

	dvs, err := f.GetDataValidations(sheet)
	if err != nil {
		return nil
	}

	for _, dv := range dvs {
		if !cellInValidationRange(row, col, dv.Sqref) {
			continue
		}
		return extractValidationMetadata(f, sheet, dv, addr)
	}
	return nil
}

// cellInValidationRange ports _cell_in_validation_range: a cell matches if it
// falls inside any of the sqref's ranges.
func cellInValidationRange(row, col int, sqref string) bool {
	for _, ref := range splitSqref(sqref) {
		minRow, minCol, maxRow, maxCol, err := parseRangeBounds(ref)
		if err != nil {
			continue
		}
		if minRow > maxRow {
			minRow, maxRow = maxRow, minRow
		}
		if minCol > maxCol {
			minCol, maxCol = maxCol, minCol
		}
		if minRow <= row && row <= maxRow && minCol <= col && col <= maxCol {
			return true
		}
	}
	return false
}

// extractValidationMetadata ports _extract_validation_metadata. Key order is
// insertion order, because the result is rendered by json.dumps (SPEC 5.5).
func extractValidationMetadata(f *excelize.File, sheet string, dv *excelize.DataValidation, addr string) *pyfmt.Dict {
	info := pyfmt.NewDict().
		Set("cell", addr).
		Set("has_validation", true).
		Set("validation_type", nilIfEmpty(dv.Type)).
		Set("allow_blank", dv.AllowBlank)

	if dv.Operator != "" {
		info.Set("operator", dv.Operator)
	}
	// These are *string in excelize; Python tests them for truthiness, so an
	// unset pointer and an empty string are both omitted.
	if v := deref(dv.Prompt); v != "" {
		info.Set("prompt", v)
	}
	if v := deref(dv.PromptTitle); v != "" {
		info.Set("prompt_title", v)
	}
	if v := deref(dv.Error); v != "" {
		info.Set("error_message", v)
	}
	if v := deref(dv.ErrorTitle); v != "" {
		info.Set("error_title", v)
	}

	formula1 := firstFormula(dv.Formula1)
	formula2 := firstFormula(dv.Formula2)

	if dv.Type == "list" && formula1 != "" {
		info.Set("allowed_values", extractListValues(f, sheet, formula1))
	} else if formula1 != "" {
		info.Set("formula1", formula1)
		if formula2 != "" {
			info.Set("formula2", formula2)
		}
	}
	return info
}

func nilIfEmpty(s string) any {
	if s == "" {
		return nil
	}
	return s
}

// firstFormula strips excelize's leading "=" so the value matches openpyxl's
// formula1, which carries no equals sign.
func firstFormula(s string) string {
	if len(s) > 0 && s[0] == '=' {
		return s[1:]
	}
	return s
}

// extractListValues ports _extract_list_values.
//
// None of the 25 tools can create a data validation, so only a workbook
// authored elsewhere reaches the range-reference branch.
func extractListValues(f *excelize.File, sheet, formula string) []any {
	formula = strings.Trim(formula, `"`)

	if strings.Contains(formula, ",") {
		out := []any{}
		for _, part := range strings.Split(formula, ",") {
			v := strings.Trim(strings.TrimSpace(part), `"`)
			if v != "" {
				out = append(out, v)
			}
		}
		return out
	}

	if strings.Contains(formula, ":") || strings.HasPrefix(formula, "$") {
		if values, ok := resolveRangeValues(f, sheet, formula); ok {
			return values
		}
		return []any{"Range: " + formula + " (empty or unresolvable)"}
	}

	return []any{strings.Trim(formula, `"`)}
}

// resolveRangeValues walks a range reference and returns the non-empty values
// it contains, mirroring `worksheet[range_ref]` in _extract_list_values.
func resolveRangeValues(f *excelize.File, sheet, ref string) ([]any, bool) {
	ref = strings.TrimPrefix(ref, "=")
	target := sheet
	if name, rest, found := strings.Cut(ref, "!"); found {
		target = strings.Trim(name, "'")
		ref = rest
	}
	ref = strings.ReplaceAll(ref, "$", "")

	minRow, minCol, maxRow, maxCol, err := parseRangeBounds(ref)
	if err != nil {
		return nil, false
	}

	out := []any{}
	for row := minRow; row <= maxRow; row++ {
		for col := minCol; col <= maxCol; col++ {
			addr, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, false
			}
			v, err := CellValue(f, target, addr)
			if err != nil {
				return nil, false
			}
			if v == nil {
				continue
			}
			out = append(out, pyfmt.StrOf(v))
		}
	}
	if len(out) == 0 {
		return nil, false
	}
	return out, true
}

func deref(s *string) string {
	if s == nil {
		return ""
	}
	return *s
}
