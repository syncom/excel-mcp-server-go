package excelops

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// CopySheet ports sheet.copy_sheet.
//
// PARITY(sheet.py:copy_sheet): rejecting an existing target is deliberate and
// protective, not a defect (SPEC 5.3.1 K3). Ported as-is.
func CopySheet(path, sourceSheet, targetSheet string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sourceSheet) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Source sheet '%s' not found", sourceSheet))
	}
	if hasSheet(f, targetSheet) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Target sheet '%s' already exists", targetSheet))
	}

	// openpyxl's copy_worksheet appends the copy at the end and then the caller
	// retitles it; excelize needs the destination to exist before copying into
	// it, which lands it in the same place.
	idx, err := f.NewSheet(targetSheet)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	srcIdx, err := f.GetSheetIndex(sourceSheet)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.CopySheet(srcIdx, idx); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Sheet '%s' copied to '%s'", sourceSheet, targetSheet), nil
}

// DeleteSheet ports sheet.delete_sheet.
//
// PARITY(sheet.py:delete_sheet): refusing to delete the only sheet is
// deliberate and protective (SPEC 5.3.1 K3). Ported as-is.
func DeleteSheet(path, sheetName string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}
	if len(f.GetSheetList()) == 1 {
		return "", excelerr.New(excelerr.ErrSheet, "Cannot delete the only sheet in workbook")
	}

	if err := f.DeleteSheet(sheetName); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Sheet '%s' deleted", sheetName), nil
}

// RenameSheet ports sheet.rename_sheet.
func RenameSheet(path, oldName, newName string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, oldName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", oldName))
	}
	if hasSheet(f, newName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' already exists", newName))
	}

	if err := f.SetSheetName(oldName, newName); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Sheet renamed from '%s' to '%s'", oldName, newName), nil
}

// FormatRangeString ports sheet.format_range_string.
func FormatRangeString(startRow, startCol, endRow, endCol int) string {
	startName, _ := excelize.ColumnNumberToName(startCol)
	endName, _ := excelize.ColumnNumberToName(endCol)
	return fmt.Sprintf("%s%d:%s%d", startName, startRow, endName, endRow)
}

// MergeRange ports sheet.merge_range.
func MergeRange(path, sheetName, startCell, endCell string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	startRow, startCol, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, err.Error())
	}
	if endRow == 0 || endCol == 0 {
		return "", excelerr.New(excelerr.ErrSheet, "Both start and end cells must be specified for merging")
	}

	rangeString := FormatRangeString(startRow, startCol, endRow, endCol)
	if err := f.MergeCell(sheetName, fmt.Sprintf("%s%d", colName(startCol), startRow),
		fmt.Sprintf("%s%d", colName(endCol), endRow)); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Range '%s' merged in sheet '%s'", rangeString, sheetName), nil
}

// UnmergeRange ports sheet.unmerge_range.
func UnmergeRange(path, sheetName, startCell, endCell string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	startRow, startCol, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, err.Error())
	}
	if endRow == 0 || endCol == 0 {
		return "", excelerr.New(excelerr.ErrSheet, "Both start and end cells must be specified for unmerging")
	}

	rangeString := FormatRangeString(startRow, startCol, endRow, endCol)

	merged, err := f.GetMergeCells(sheetName)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	found := false
	for _, m := range merged {
		if strings.EqualFold(m.GetStartAxis()+":"+m.GetEndAxis(), rangeString) {
			found = true
			break
		}
	}
	if !found {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Range '%s' is not merged", rangeString))
	}

	if err := f.UnmergeCell(sheetName, fmt.Sprintf("%s%d", colName(startCol), startRow),
		fmt.Sprintf("%s%d", colName(endCol), endRow)); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Range '%s' unmerged successfully", rangeString), nil
}

// GetMergedRanges ports sheet.get_merged_ranges.
func GetMergedRanges(path, sheetName string) ([]string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return nil, excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	merged, err := f.GetMergeCells(sheetName)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	out := make([]string, 0, len(merged))
	for _, m := range merged {
		out = append(out, m.GetStartAxis()+":"+m.GetEndAxis())
	}
	return out, nil
}

func colName(col int) string {
	name, _ := excelize.ColumnNumberToName(col)
	return name
}

// CopyRangeOperation ports sheet.copy_range_operation.
func CopyRangeOperation(path, sheetName, sourceStart, sourceEnd, targetStart, targetSheet string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Failed to copy range: %s", pyStr(err)))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}
	if targetSheet == "" {
		targetSheet = sheetName
	}
	if !hasSheet(f, targetSheet) {
		// openpyxl's wb[target_sheet] raises KeyError, which the outer handler
		// turns into SheetError("Failed to copy range: ...").
		return "", excelerr.New(excelerr.ErrSheet,
			fmt.Sprintf("Failed to copy range: 'Worksheet %s does not exist.'", targetSheet))
	}

	startRow, startCol, endRow, endCol, err := ParseCellRange(sourceStart, sourceEnd)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid source range: %s", err.Error()))
	}
	if endRow == 0 {
		// PARITY(sheet.py:copy_range_operation): there is no single-cell
		// fallback here — the helper sheet.copy_range has one, this function
		// does not. `range(start_row, None + 1)` raises a TypeError that the
		// outer handler wraps.
		return "", excelerr.New(excelerr.ErrSheet,
			"Failed to copy range: unsupported operand type(s) for +: 'NoneType' and 'int'")
	}

	targetRow, targetCol, err := parseLooseCellRef(targetStart)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid target cell: %s", err.Error()))
	}

	rowOffset := targetRow - startRow
	colOffset := targetCol - startCol

	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if err := copyCell(f, sheetName, col, row, targetSheet, col+colOffset, row+rowOffset); err != nil {
				return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Failed to copy range: %s", pyStr(err)))
			}
		}
	}

	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Failed to copy range: %s", pyStr(err)))
	}
	return "Range copied successfully", nil
}

// parseLooseCellRef ports copy_range_operation's target parsing, which pulls
// the digits and letters out separately rather than using parse_cell_range.
//
// PARITY(sheet.py:copy_range_operation): int(”.join(filter(str.isdigit, ...)))
// on a reference with no digits raises with CPython's own message, which the
// caller interpolates verbatim.
func parseLooseCellRef(ref string) (row, col int, err error) {
	var digits, letters string
	for _, c := range ref {
		switch {
		case c >= '0' && c <= '9':
			digits += string(c)
		case c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z':
			letters += string(c)
		}
	}
	if digits == "" {
		return 0, 0, fmt.Errorf("invalid literal for int() with base 10: ''")
	}
	row, err = strconv.Atoi(digits)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid literal for int() with base 10: '%s'", digits)
	}
	col, err = excelize.ColumnNameToNumber(strings.ToUpper(letters))
	if err != nil {
		return 0, 0, err
	}
	return row, col, nil
}

// copyCell moves one cell's value and style, matching openpyxl's
// `target_cell.value = source_cell.value` plus its `_style` copy.
func copyCell(f *excelize.File, srcSheet string, srcCol, srcRow int, dstSheet string, dstCol, dstRow int) error {
	if dstCol < 1 || dstRow < 1 {
		// openpyxl's Worksheet.cell rejects a non-positive coordinate rather
		// than skipping it, and sheet.py's `except Exception` turns that into
		// "Failed to copy range: ...".
		return errCellOutOfBounds
	}
	src, err := excelize.CoordinatesToCellName(srcCol, srcRow)
	if err != nil {
		return err
	}
	dst, err := excelize.CoordinatesToCellName(dstCol, dstRow)
	if err != nil {
		return err
	}
	return moveCellContents(f, srcSheet, src, dstSheet, dst, false)
}

// errCellOutOfBounds carries openpyxl's message for a non-positive coordinate.
var errCellOutOfBounds = errors.New("Row or column values must be at least 1")

// moveCellContents copies (or, when clearSource is set, moves) a cell's value,
// formula and style.
func moveCellContents(f *excelize.File, srcSheet, src, dstSheet, dst string, clearSource bool) error {
	formula, err := f.GetCellFormula(srcSheet, src)
	if err != nil {
		return err
	}
	if formula != "" {
		if err := f.SetCellFormula(dstSheet, dst, formula); err != nil {
			return err
		}
	} else {
		raw, err := f.GetCellValue(srcSheet, src, excelize.Options{RawCellValue: true})
		if err != nil {
			return err
		}
		typ, err := f.GetCellType(srcSheet, src)
		if err != nil {
			return err
		}
		if err := setRawCell(f, dstSheet, dst, raw, typ); err != nil {
			return err
		}
	}

	styleID, err := f.GetCellStyle(srcSheet, src)
	if err != nil {
		return err
	}
	if styleID != 0 {
		if err := f.SetCellStyle(dstSheet, dst, dst, styleID); err != nil {
			return err
		}
	}

	if clearSource {
		if err := f.SetCellValue(srcSheet, src, nil); err != nil {
			return err
		}
		if err := f.SetCellStyle(srcSheet, src, src, 0); err != nil {
			return err
		}
	}
	return nil
}

// setRawCell writes a raw value back with its original cell type.
func setRawCell(f *excelize.File, sheet, cell, raw string, typ excelize.CellType) error {
	if raw == "" {
		return f.SetCellValue(sheet, cell, nil)
	}
	switch typ {
	case excelize.CellTypeBool:
		return f.SetCellBool(sheet, cell, raw == "1" || strings.EqualFold(raw, "true"))
	case excelize.CellTypeSharedString, excelize.CellTypeInlineString:
		return f.SetCellStr(sheet, cell, raw)
	default:
		return f.SetCellDefault(sheet, cell, raw)
	}
}

// DeleteRangeOperation ports sheet.delete_range_operation, including D4.
func DeleteRangeOperation(path, sheetName, startCell, endCell, shiftDirection string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	maxRow, maxCol, err := UsedRange(f, sheetName)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}

	startRow, startCol, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Invalid range: %s", err.Error()))
	}
	if endRow != 0 && endRow > maxRow {
		return "", excelerr.New(excelerr.ErrSheet,
			fmt.Sprintf("End row %d out of bounds (1-%d)", endRow, maxRow))
	}
	if endCol != 0 && endCol > maxCol {
		return "", excelerr.New(excelerr.ErrSheet,
			fmt.Sprintf("End column %d out of bounds (1-%d)", endCol, maxCol))
	}

	if shiftDirection != "up" && shiftDirection != "left" {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid shift direction: %s. Must be 'up' or 'left'", shiftDirection))
	}

	effEndRow, effEndCol := endRow, endCol
	if effEndRow == 0 {
		effEndRow = startRow
	}
	if effEndCol == 0 {
		effEndCol = startCol
	}
	rangeString := FormatRangeString(startRow, startCol, effEndRow, effEndCol)

	// DEVIATION(D4): Python blanks the named range and then *additionally*
	// calls delete_rows/delete_cols for the same span, so
	// delete_range("A1:B2", shift="up") deletes entire rows 1-2 across every
	// column — silent data loss well outside what the caller named. We
	// implement Excel's actual "delete cells, shift up/left", confined to the
	// range's own column span (for "up") or row span (for "left"). Cells
	// outside that span are untouched. SPEC 5.3.2 D4; registered in
	// conformance/deviations.json.
	if err := shiftCells(f, sheetName, startRow, startCol, effEndRow, effEndCol, maxRow, maxCol, shiftDirection); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}

	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return fmt.Sprintf("Range %s deleted successfully", rangeString), nil
}

// shiftCells removes the range's cells and pulls the cells beyond it into the
// gap, staying inside the range's own span.
func shiftCells(f *excelize.File, sheet string, startRow, startCol, endRow, endCol, maxRow, maxCol int, direction string) error {
	if direction == "up" {
		span := endRow - startRow + 1
		for col := startCol; col <= endCol; col++ {
			for row := startRow; row+span <= maxRow; row++ {
				if err := moveCell(f, sheet, col, row+span, col, row); err != nil {
					return err
				}
			}
			// Clear the tail the shift vacated.
			for row := maxRow - span + 1; row <= maxRow; row++ {
				if row < startRow {
					continue
				}
				if err := clearCellAt(f, sheet, col, row); err != nil {
					return err
				}
			}
		}
		return nil
	}

	span := endCol - startCol + 1
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col+span <= maxCol; col++ {
			if err := moveCell(f, sheet, col+span, row, col, row); err != nil {
				return err
			}
		}
		for col := maxCol - span + 1; col <= maxCol; col++ {
			if col < startCol {
				continue
			}
			if err := clearCellAt(f, sheet, col, row); err != nil {
				return err
			}
		}
	}
	return nil
}

func moveCell(f *excelize.File, sheet string, srcCol, srcRow, dstCol, dstRow int) error {
	src, err := excelize.CoordinatesToCellName(srcCol, srcRow)
	if err != nil {
		return err
	}
	dst, err := excelize.CoordinatesToCellName(dstCol, dstRow)
	if err != nil {
		return err
	}
	return moveCellContents(f, sheet, src, sheet, dst, false)
}

func clearCellAt(f *excelize.File, sheet string, col, row int) error {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	if err := f.SetCellValue(sheet, cell, nil); err != nil {
		return err
	}
	return f.SetCellStyle(sheet, cell, cell, 0)
}

// InsertRow ports sheet.insert_row.
func InsertRow(path, sheetName string, startRow, count int) (string, error) {
	return rowColOp(path, sheetName, startRow, count, opInsertRow)
}

// InsertCols ports sheet.insert_cols.
func InsertCols(path, sheetName string, startCol, count int) (string, error) {
	return rowColOp(path, sheetName, startCol, count, opInsertCol)
}

// DeleteRows ports sheet.delete_rows.
func DeleteRows(path, sheetName string, startRow, count int) (string, error) {
	return rowColOp(path, sheetName, startRow, count, opDeleteRow)
}

// DeleteCols ports sheet.delete_cols.
func DeleteCols(path, sheetName string, startCol, count int) (string, error) {
	return rowColOp(path, sheetName, startCol, count, opDeleteCol)
}

type rowColKind int

const (
	opInsertRow rowColKind = iota
	opInsertCol
	opDeleteRow
	opDeleteCol
)

// rowColOp holds the shape all four functions share: the same guards in the
// same order, differing only in the axis, the bounds check the delete variants
// add, and the message.
func rowColOp(path, sheetName string, start, count int, kind rowColKind) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrSheet, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	isRow := kind == opInsertRow || kind == opDeleteRow
	if start < 1 {
		if isRow {
			return "", excelerr.New(excelerr.ErrValidation, "Start row must be 1 or greater")
		}
		return "", excelerr.New(excelerr.ErrValidation, "Start column must be 1 or greater")
	}
	if count < 1 {
		return "", excelerr.New(excelerr.ErrValidation, "Count must be 1 or greater")
	}

	maxRow, maxCol, err := UsedRange(f, sheetName)
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}

	// Only the delete variants bound-check the start.
	switch kind {
	case opDeleteRow:
		if start > maxRow {
			return "", excelerr.New(excelerr.ErrValidation,
				fmt.Sprintf("Start row %d exceeds worksheet bounds (max row: %d)", start, maxRow))
		}
	case opDeleteCol:
		if start > maxCol {
			return "", excelerr.New(excelerr.ErrValidation,
				fmt.Sprintf("Start column %d exceeds worksheet bounds (max column: %d)", start, maxCol))
		}
	}

	colName, err := excelize.ColumnNumberToName(max(start, 1))
	if err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}

	var msg string
	switch kind {
	case opInsertRow:
		if err := f.InsertRows(sheetName, start, count); err != nil {
			return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
		}
		msg = fmt.Sprintf("Inserted %d row(s) starting at row %d in sheet '%s'", count, start, sheetName)
	case opInsertCol:
		if err := f.InsertCols(sheetName, colName, count); err != nil {
			return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
		}
		msg = fmt.Sprintf("Inserted %d column(s) starting at column %d in sheet '%s'", count, start, sheetName)
	case opDeleteRow:
		// excelize removes one row per call; openpyxl takes a count.
		for i := 0; i < count; i++ {
			if err := f.RemoveRow(sheetName, start); err != nil {
				return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
			}
		}
		msg = fmt.Sprintf("Deleted %d row(s) starting at row %d in sheet '%s'", count, start, sheetName)
	case opDeleteCol:
		for i := 0; i < count; i++ {
			if err := f.RemoveCol(sheetName, colName); err != nil {
				return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
			}
		}
		msg = fmt.Sprintf("Deleted %d column(s) starting at column %d in sheet '%s'", count, start, sheetName)
	}

	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrSheet, pyStr(err))
	}
	return msg, nil
}
