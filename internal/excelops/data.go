package excelops

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// WriteData ports data.write_data.
func WriteData(path, sheetName string, data [][]json.RawMessage, startCell string) (string, error) {
	if len(data) == 0 {
		return "", excelerr.New(excelerr.ErrData, "No data provided to write")
	}

	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrData, pyStr(err))
	}
	defer f.Close()

	if sheetName == "" {
		// `if not sheet_name: sheet_name = wb.active.title`
		list := f.GetSheetList()
		if len(list) == 0 {
			return "", excelerr.New(excelerr.ErrData, "No active sheet found in workbook")
		}
		idx := f.GetActiveSheetIndex()
		if idx < 0 || idx >= len(list) {
			idx = 0
		}
		sheetName = list[idx]
	} else if !hasSheet(f, sheetName) {
		if _, err := f.NewSheet(sheetName); err != nil {
			return "", excelerr.New(excelerr.ErrData, pyStr(err))
		}
	}

	startRow, startCol, _, _, err := ParseCellRange(startCell, "")
	if err != nil {
		// parse_cell_range raises ValueError; write_data turns it into
		// DataError(f"Invalid start cell format: {e}").
		return "", excelerr.New(excelerr.ErrData, fmt.Sprintf("Invalid start cell format: %s", err.Error()))
	}

	for i, row := range data {
		for j, raw := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return "", excelerr.New(excelerr.ErrData, pyStr(err))
			}
			if err := setCellFromJSON(f, sheetName, cell, raw); err != nil {
				return "", excelerr.New(excelerr.ErrData, pyStr(err))
			}
		}
	}

	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrData, pyStr(err))
	}
	return fmt.Sprintf("Data written to %s", sheetName), nil
}

// setCellFromJSON writes one JSON argument value the way openpyxl's
// `cell.value = v` would.
//
// The value is kept as raw JSON up to this point on purpose: Python's json
// parser yields int for `10` and float for `10.0`, and openpyxl stores the
// distinction. Decoding into a Go `any` would collapse both to float64 and
// write "10" as 10.0.
func setCellFromJSON(f *excelize.File, sheet, cell string, raw json.RawMessage) error {
	text := strings.TrimSpace(string(raw))
	switch {
	case text == "" || text == "null" || text == `""`:
		// openpyxl maps both None and the empty string to an absent cell, so
		// writing either must leave no cell behind. Anything else shows up as
		// an extra cell in the SPEC 7 structural comparison.
		return clearCell(f, sheet, cell)
	case text == "true":
		return f.SetCellBool(sheet, cell, true)
	case text == "false":
		return f.SetCellBool(sheet, cell, false)
	case text[0] == '"':
		var s string
		if err := json.Unmarshal(raw, &s); err != nil {
			return err
		}
		return f.SetCellStr(sheet, cell, s)
	case text[0] == '[' || text[0] == '{':
		// openpyxl rejects a list or dict cell value with
		// "Cannot convert {value!r} to Excel"; reproduced in the same shape.
		return fmt.Errorf("Cannot convert %s to Excel", text)
	default:
		// MEASURED against the oracle: every JSON number reaches openpyxl as a
		// float, and openpyxl's safe_string writes it with "%.16g". That one
		// rule explains the whole observed table — 10 and 3.0 both store as
		// "10"/"3" and read back as ints, 1e16 stores as "1e+16" and reads back
		// as a float, and 2**53+1 loses its last digit exactly as Python does.
		// Writing the JSON literal instead would make Go *more* precise than
		// the oracle, which SPEC 5.1 does not permit and no registry entry
		// sanctions.
		var fl float64
		if err := json.Unmarshal(raw, &fl); err != nil {
			return err
		}
		return f.SetCellDefault(sheet, cell, pyfmt.FormatG(fl, 16))
	}
}

// clearCell leaves the cell absent, which is what openpyxl produces for a None
// or empty-string assignment. A cell that already holds something is blanked
// rather than left stale.
func clearCell(f *excelize.File, sheet, cell string) error {
	current, err := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		return err
	}
	if current == "" {
		return nil
	}
	return f.SetCellValue(sheet, cell, nil)
}

// ReadExcelRangeWithMetadata ports data.read_excel_range_with_metadata.
func ReadExcelRangeWithMetadata(path, sheetName, startCell, endCell string) (*pyfmt.Dict, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	// `if ':' in start_cell: start_cell, end_cell = start_cell.split(':')` —
	// the tuple unpack rejects a second colon, and data.py does not guard it,
	// so the ValueError reaches the caller.
	if strings.Contains(startCell, ":") {
		before, after, err := pySplit2(startCell, ":")
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, err.Error())
		}
		startCell, endCell = before, after
	}

	startRow, startCol, _, _, err := ParseCellRange(startCell+":"+startCell, "")
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Invalid start cell format: %s", err.Error()))
	}

	maxRow, maxCol, err := UsedRange(f, sheetName)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}
	minRow, minCol, err := usedRangeOrigin(f, sheetName)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}

	var endRow, endCol int
	if endCell != "" {
		endRow, endCol, _, _, err = ParseCellRange(endCell+":"+endCell, "")
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Invalid end cell format: %s", err.Error()))
		}
	} else {
		empty, err := sheetIsEmpty(f, sheetName, maxRow, maxCol)
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, pyStr(err))
		}
		if empty {
			endRow, endCol = startRow, startCol
		} else {
			endRow, endCol = maxRow, maxCol
			// PARITY(data.py:read_excel_range_with_metadata): when end_cell is
			// omitted and start_cell is the default "A1", the true start is
			// re-derived from the used range, so the response can report a
			// range that does not begin at A1. This is intended behavior, not a
			// defect — the reply names the range actually read (SPEC 5.3.1 K1).
			if startCell == "A1" {
				startRow, startCol = minRow, minCol
			}
		}
	}

	if startRow > maxRow || startCol > maxCol {
		// Python returns a range string with a trailing colon and no end.
		return pyfmt.NewDict().
			Set("range", startCell+":").
			Set("sheet_name", sheetName).
			Set("cells", []any{}), nil
	}

	// Bound the sweep before the loop below allocates a dict per cell. See
	// [checkRangeSize]; read_data_from_excel catches nothing, so this surfaces
	// as a tool error exactly as an uncaught DataError does in Python.
	if err := checkRangeSize(excelerr.ErrData, startRow, startCol, endRow, endCol); err != nil {
		return nil, err
	}

	startColName, err := excelize.ColumnNumberToName(startCol)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}
	endColName, err := excelize.ColumnNumberToName(endCol)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}

	cells := []any{}
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			addr, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrData, pyStr(err))
			}
			value, err := CellValue(f, sheetName, addr)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrData, pyStr(err))
			}
			cell := pyfmt.NewDict().
				Set("address", addr).
				Set("value", value).
				Set("row", row).
				Set("column", col)

			// server.py never passes include_validation, so it is always true.
			if info := DataValidationForCell(f, sheetName, addr); info != nil {
				cell.Set("validation", info)
			} else {
				cell.Set("validation", pyfmt.NewDict().Set("has_validation", false))
			}
			cells = append(cells, cell)
		}
	}

	return pyfmt.NewDict().
		Set("range", fmt.Sprintf("%s%d:%s%d", startColName, startRow, endColName, endRow)).
		Set("sheet_name", sheetName).
		Set("cells", cells), nil
}

// sheetIsEmpty reproduces Python's
// `ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None`.
func sheetIsEmpty(f *excelize.File, sheet string, maxRow, maxCol int) (bool, error) {
	if maxRow != 1 || maxCol != 1 {
		return false, nil
	}
	v, err := CellValue(f, sheet, "A1")
	if err != nil {
		return false, err
	}
	return v == nil, nil
}

// ReadExcelRange ports data.read_excel_range: values only, no metadata, and
// rows that are entirely empty are dropped.
func ReadExcelRange(path, sheetName, startCell, endCell string) ([][]any, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	if strings.Contains(startCell, ":") {
		before, after, err := pySplit2(startCell, ":")
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, err.Error())
		}
		startCell, endCell = before, after
	}

	startRow, startCol, _, _, err := ParseCellRange(startCell+":"+startCell, "")
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Invalid start cell format: %s", err.Error()))
	}

	maxRow, maxCol, err := UsedRange(f, sheetName)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrData, pyStr(err))
	}

	var endRow, endCol int
	if endCell != "" {
		endRow, endCol, _, _, err = ParseCellRange(endCell+":"+endCell, "")
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, fmt.Sprintf("Invalid end cell format: %s", err.Error()))
		}
	} else {
		empty, err := sheetIsEmpty(f, sheetName, maxRow, maxCol)
		if err != nil {
			return nil, excelerr.New(excelerr.ErrData, pyStr(err))
		}
		if empty {
			endRow, endCol = startRow, startCol
		} else {
			// PARITY(data.py:read_excel_range): unlike the with-metadata
			// variant, this one moves the *start* to the used range's origin
			// regardless of what start_cell was.
			minRow, minCol, err := usedRangeOrigin(f, sheetName)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrData, pyStr(err))
			}
			startRow, startCol = minRow, minCol
			endRow, endCol = maxRow, maxCol
		}
	}

	if startRow > maxRow || startCol > maxCol {
		return nil, nil
	}

	var out [][]any
	for row := startRow; row <= endRow; row++ {
		rowData := make([]any, 0, endCol-startCol+1)
		any_ := false
		for col := startCol; col <= endCol; col++ {
			addr, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrData, pyStr(err))
			}
			v, err := CellValue(f, sheetName, addr)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrData, pyStr(err))
			}
			if v != nil {
				any_ = true
			}
			rowData = append(rowData, v)
		}
		if any_ {
			out = append(out, rowData)
		}
	}
	return out, nil
}
