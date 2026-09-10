package excelops

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// DefaultSheetName mirrors workbook.create_workbook's sheet_name default.
const DefaultSheetName = "Sheet1"

// CreateWorkbook ports workbook.create_workbook.
//
// openpyxl's Workbook() starts with one sheet called "Sheet" and renames it;
// excelize's NewFile() starts with one called "Sheet1". The Python behavior
// wins (STANDARDS), so the sheet is renamed either way and the result is a
// single sheet named sheetName.
func CreateWorkbook(path, sheetName string) error {
	if err := createWorkbook(path, sheetName); err != nil {
		return excelerr.New(excelerr.ErrWorkbook, fmt.Sprintf("Failed to create workbook: %s", pyStr(err)))
	}
	return nil
}

func createWorkbook(path, sheetName string) error {
	f := excelize.NewFile()
	defer f.Close()

	if sheetName != "Sheet1" {
		if err := f.SetSheetName("Sheet1", sheetName); err != nil {
			return err
		}
	}

	// Path(filepath).parent.mkdir(parents=True, exist_ok=True)
	if dir := filepath.Dir(path); dir != "" {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return err
		}
	}
	return f.SaveAs(path)
}

// CreateSheet ports workbook.create_sheet. Returns the success message.
func CreateSheet(path, sheetName string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrWorkbook, pyStr(err))
	}
	defer f.Close()

	if hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrWorkbook, fmt.Sprintf("Sheet %s already exists", sheetName))
	}

	if _, err := f.NewSheet(sheetName); err != nil {
		return "", excelerr.New(excelerr.ErrWorkbook, pyStr(err))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrWorkbook, pyStr(err))
	}
	return fmt.Sprintf("Sheet %s created successfully", sheetName), nil
}

// GetWorkbookInfo ports workbook.get_workbook_info.
//
// The returned Dict is in the key order SPEC 5.4 pins, because the tool result
// is its Python repr.
func GetWorkbookInfo(path string, includeRanges bool) (*pyfmt.Dict, error) {
	// Path.exists() is false only for a missing path; a directory exists, so
	// Python gets past this check and fails later in load_workbook.
	st, err := os.Stat(path)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrWorkbook, fmt.Sprintf("File not found: %s", path))
	}

	f, err := openWorkbook(path)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrWorkbook, pyStr(err))
	}
	defer f.Close()

	info := pyfmt.NewDict().
		Set("filename", filepath.Base(path)).
		Set("sheets", f.GetSheetList()).
		Set("size", int(st.Size())).
		Set("modified", float64(st.ModTime().UnixNano())/1e9)

	if includeRanges {
		ranges := pyfmt.NewDict()
		for _, name := range f.GetSheetList() {
			maxRow, maxCol, err := UsedRange(f, name)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrWorkbook, pyStr(err))
			}
			// Python guards on max_row > 0 and max_column > 0, which openpyxl
			// never violates, so the guard never excludes a sheet.
			col, err := excelize.ColumnNumberToName(maxCol)
			if err != nil {
				return nil, excelerr.New(excelerr.ErrWorkbook, pyStr(err))
			}
			ranges.Set(name, fmt.Sprintf("A1:%s%d", col, maxRow))
		}
		info.Set("used_ranges", ranges)
	}

	return info, nil
}

// OpenOrCreateWorkbook ports workbook.get_or_create_workbook.
//
// PARITY(workbook.py:get_or_create_workbook): a missing file is not an error
// here — format_range creates it. That is unlike apply_formula, where
// server.py's prior validate_formula_in_cell_operation loads the file first and
// makes this branch unreachable (SPEC 5.3.1 K2).
func OpenOrCreateWorkbook(path string) (*excelize.File, error) {
	f, err := excelize.OpenFile(path)
	if err == nil {
		return f, nil
	}
	if !os.IsNotExist(err) {
		return nil, err
	}
	if err := createWorkbook(path, DefaultSheetName); err != nil {
		return nil, err
	}
	return excelize.OpenFile(path)
}

// openWorkbook loads an existing workbook, translating a missing file into the
// OSError text openpyxl would have produced.
func openWorkbook(path string) (*excelize.File, error) {
	if _, err := os.Stat(path); err != nil {
		return nil, err
	}
	return excelize.OpenFile(path)
}

// hasSheet reports whether the workbook already has that sheet, matching
// `sheet_name in wb.sheetnames`.
func hasSheet(f *excelize.File, name string) bool {
	for _, s := range f.GetSheetList() {
		if s == name {
			return true
		}
	}
	return false
}

// pyStr renders an error the way Python's str(e) would inside the
// `raise XxxError(str(e))` handlers that wrap every unexpected failure.
func pyStr(err error) string {
	if msg, ok := PyOSError(err); ok {
		return msg
	}
	return err.Error()
}
