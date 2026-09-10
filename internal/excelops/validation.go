package excelops

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// unsafeFuncs is validation.validate_formula's blocked set.
var unsafeFuncs = map[string]bool{
	"INDIRECT": true, "HYPERLINK": true, "WEBSERVICE": true, "DGET": true, "RTD": true,
}

// funcPattern mirrors r"([A-Z]+)\(" — uppercase only, so a lower-case
// indirect( slips through on the Python side too.
var funcPattern = regexp.MustCompile(`([A-Z]+)\(`)

// cellRefsInFormula mirrors r'[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?'.
var cellRefsInFormula = regexp.MustCompile(`[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?`)

// ValidateFormula ports validation.validate_formula.
func ValidateFormula(formula string) (bool, string) {
	if !strings.HasPrefix(formula, "=") {
		return false, "Formula must start with '='"
	}
	body := formula[1:]

	parens := 0
	for _, c := range body {
		switch c {
		case '(':
			parens++
		case ')':
			parens--
		}
		if parens < 0 {
			return false, "Unmatched closing parenthesis"
		}
	}
	if parens > 0 {
		return false, "Unclosed parenthesis"
	}

	for _, m := range funcPattern.FindAllStringSubmatch(body, -1) {
		if unsafeFuncs[m[1]] {
			return false, fmt.Sprintf("Unsafe function: %s", m[1])
		}
	}
	return true, "Formula is valid"
}

// ValidateCellReference ports cell_utils.validate_cell_reference.
//
// PARITY(cell_utils.py:validate_cell_reference): it accepts lower case and does
// not bound the column, so "a1" and "ZZZZ1" both pass. Only the letters-then-
// digits shape is enforced.
func ValidateCellReference(ref string) bool {
	if ref == "" {
		return false
	}
	var col, row string
	for _, c := range ref {
		switch {
		case c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z':
			if row != "" {
				return false // letters after numbers
			}
			col += string(c)
		case c >= '0' && c <= '9':
			row += string(c)
		default:
			return false
		}
	}
	return col != "" && row != ""
}

// ValidateRangeBounds ports validation.validate_range_bounds. endRow/endCol of
// zero stand in for Python's None.
func ValidateRangeBounds(maxRow, maxCol, startRow, startCol, endRow, endCol int) (bool, string) {
	if startRow < 1 || startRow > maxRow {
		return false, fmt.Sprintf("Start row %d out of bounds (1-%d)", startRow, maxRow)
	}
	if startCol < 1 || startCol > maxCol {
		startName, _ := excelize.ColumnNumberToName(startCol)
		maxName, _ := excelize.ColumnNumberToName(maxCol)
		return false, fmt.Sprintf("Start column %s out of bounds (A-%s)", startName, maxName)
	}

	if endRow != 0 && endCol != 0 {
		if endRow < startRow {
			return false, "End row cannot be before start row"
		}
		if endCol < startCol {
			return false, "End column cannot be before start column"
		}
		if endRow > maxRow {
			return false, fmt.Sprintf("End row %d out of bounds (1-%d)", endRow, maxRow)
		}
		if endCol > maxCol {
			endName, _ := excelize.ColumnNumberToName(endCol)
			maxName, _ := excelize.ColumnNumberToName(maxCol)
			return false, fmt.Sprintf("End column %s out of bounds (A-%s)", endName, maxName)
		}
	}
	return true, "Range is valid"
}

// ValidateFormulaInCell ports validation.validate_formula_in_cell_operation.
// It returns the message server.py hands back.
func ValidateFormulaInCell(path, sheetName, cell, formula string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}
	if !ValidateCellReference(cell) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid cell reference: %s", cell))
	}

	if ok, msg := ValidateFormula(formula); !ok {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid formula syntax: %s", msg))
	}

	for _, ref := range cellRefsInFormula.FindAllString(formula, -1) {
		if start, end, found := strings.Cut(ref, ":"); found {
			if !ValidateCellReference(start) || !ValidateCellReference(end) {
				return "", excelerr.New(excelerr.ErrValidation,
					fmt.Sprintf("Invalid cell range reference in formula: %s", ref))
			}
		} else if !ValidateCellReference(ref) {
			return "", excelerr.New(excelerr.ErrValidation,
				fmt.Sprintf("Invalid cell reference in formula: %s", ref))
		}
	}

	current, err := currentCellFormulaOrValue(f, sheetName, cell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, pyStr(err))
	}

	if strings.HasPrefix(current, "=") {
		want := formula
		if !strings.HasPrefix(formula, "=") {
			want = "=" + formula
		}
		if current != want {
			return "Formula is valid but doesn't match cell content", nil
		}
		if !strings.HasPrefix(formula, "=") {
			return "Formula is valid and matches cell content", nil
		}
		// PARITY(validation.py:validate_formula_in_cell_operation): when the
		// provided formula already starts with '=' and equals the cell's, the
		// `if current_formula != formula` branch falls through with no return,
		// so the function returns None and server.py raises TypeError on
		// result["message"].
		return "", errNoneResult
	}

	return "Formula is valid but cell contains no formula", nil
}

// errNoneResult stands for Python returning None from a branch that has no
// return statement; server.py then fails with a TypeError, which is a
// raise-through rather than an "Error: ..." result (SPEC 5.2).
var errNoneResult = &noneResultError{}

type noneResultError struct{}

func (e *noneResultError) Error() string {
	return "'NoneType' object is not subscriptable"
}

// currentCellFormulaOrValue reproduces openpyxl's cell.value for a cell that
// may hold a formula: openpyxl returns the formula text including its '='.
func currentCellFormulaOrValue(f *excelize.File, sheet, cell string) (string, error) {
	formula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		return "", err
	}
	if formula != "" {
		if !strings.HasPrefix(formula, "=") {
			formula = "=" + formula
		}
		return formula, nil
	}
	v, err := CellValue(f, sheet, cell)
	if err != nil {
		return "", err
	}
	if s, ok := v.(string); ok {
		return s, nil
	}
	return "", nil
}

// ValidateRangeInSheet ports validation.validate_range_in_sheet_operation,
// including deviation D2.
func ValidateRangeInSheet(path, sheetName, rangeStr string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	maxRow, maxCol, err := UsedRange(f, sheetName)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, pyStr(err))
	}

	// DEVIATION(D2): server.py joins start and end into "A1:Z999" and passes it
	// as the positional start_cell; parse_cell_range uses re.match, so Python
	// silently truncates to "A1" and bounds-checks only the start. We split the
	// joined range and check both ends. SPEC 5.3.2 D2; registered in
	// conformance/deviations.json.
	//
	// The fix is confined to this function: parse_cell_range keeps its re.match
	// looseness everywhere else, because K1 and read_excel_range_with_metadata
	// depend on it (STANDARDS: parity discipline).
	startPart, endPart, hasEnd := strings.Cut(rangeStr, ":")

	startRow, startCol, _, _, err := ParseCellRange(startPart, "")
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid range: %s", err.Error()))
	}
	endRow, endCol := startRow, startCol
	if hasEnd {
		endRow, endCol, _, _, err = ParseCellRange(endPart, "")
		if err != nil {
			return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid range: %s", err.Error()))
		}
	}

	if ok, msg := ValidateRangeBounds(maxRow, maxCol, startRow, startCol, endRow, endCol); !ok {
		return "", excelerr.New(excelerr.ErrValidation, msg)
	}

	maxColName, err := excelize.ColumnNumberToName(maxCol)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, pyStr(err))
	}
	return fmt.Sprintf("Range '%s' is valid. Sheet contains data in range 'A1:%s%d'",
		rangeStr, maxColName, maxRow), nil
}

// AllValidationRanges ports cell_validation.get_all_validation_ranges.
func AllValidationRanges(f *excelize.File, sheet string) ([]any, error) {
	dvs, err := f.GetDataValidations(sheet)
	if err != nil {
		// PARITY(cell_validation.py:get_all_validation_ranges): the body is
		// wrapped in `except Exception`, which logs and returns what it has.
		return []any{}, nil
	}
	out := []any{}
	for _, dv := range dvs {
		info := pyfmt.NewDict().
			Set("ranges", dv.Sqref).
			Set("validation_type", nilIfEmpty(dv.Type)).
			Set("allow_blank", dv.AllowBlank)
		if formula := firstFormula(dv.Formula1); dv.Type == "list" && formula != "" {
			info.Set("allowed_values", extractListValues(f, sheet, formula))
		}
		out = append(out, info)
	}
	return out, nil
}

// ValidationRangesForSheet opens the workbook and returns the sheet's
// validation rules. sheetExists is false when the sheet is absent, which
// server.py reports as a text result rather than raising.
func ValidationRangesForSheet(path, sheetName string) (rules []any, sheetExists bool, err error) {
	f, err := openWorkbook(path)
	if err != nil {
		return nil, false, excelerr.New(excelerr.ErrData, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return nil, false, nil
	}
	rules, err = AllValidationRanges(f, sheetName)
	return rules, true, err
}

// ErrNoneResult reports whether err is the sentinel for Python falling off the
// end of validate_formula_in_cell_operation and returning None.
func ErrNoneResult(err error) bool {
	_, ok := err.(*noneResultError)
	return ok
}
