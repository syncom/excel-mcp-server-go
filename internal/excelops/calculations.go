package excelops

import (
	"fmt"
	"strings"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// ApplyFormula ports calculations.apply_formula.
//
// PARITY(calculations.py:apply_formula): Python reaches this through
// get_or_create_workbook, but server.py always runs
// validate_formula_in_cell_operation first, and that uses load_workbook and
// raises on a missing file. The create branch is therefore unreachable, so this
// keeps a single load path — no observable difference (SPEC 5.3.1 K2).
func ApplyFormula(path, sheetName, cell, formula string) (string, error) {
	if !ValidateCellReference(cell) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid cell reference: %s", cell))
	}

	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrCalculation, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	if !strings.HasPrefix(formula, "=") {
		formula = "=" + formula
	}

	if ok, msg := ValidateFormula(formula); !ok {
		return "", excelerr.New(excelerr.ErrCalculation, fmt.Sprintf("Invalid formula syntax: %s", msg))
	}

	// excelize stores the string as given, so a leading "=" would land in the
	// XML and openpyxl would read back "==SUM(...)". openpyxl writes the
	// formula body without it.
	if err := f.SetCellFormula(sheetName, cell, strings.TrimPrefix(formula, "=")); err != nil {
		return "", excelerr.New(excelerr.ErrCalculation,
			fmt.Sprintf("Failed to apply formula to cell: %s", pyStr(err)))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrCalculation,
			fmt.Sprintf("Failed to save workbook after applying formula: %s", pyStr(err)))
	}

	return fmt.Sprintf("Applied formula '%s' to cell %s", formula, cell), nil
}
