package excelops

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// CreateExcelTable ports tables.create_excel_table.
//
// PARITY(tables.py:create_excel_table): every failure is wrapped in DataError,
// so create_table's `except DataError` catches all of them and nothing here
// reaches the raise-through channel (SPEC 5.2).
func CreateExcelTable(path, sheetName, dataRange, tableName, tableStyle string) (string, error) {
	msg, err := createExcelTable(path, sheetName, dataRange, tableName, tableStyle)
	if err != nil {
		if e, ok := err.(*excelerr.Error); ok {
			return "", excelerr.New(excelerr.ErrData, e.Msg)
		}
		return "", excelerr.New(excelerr.ErrData, pyStr(err))
	}
	return msg, nil
}

func createExcelTable(path, sheetName, dataRange, tableName, tableStyle string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", err
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		// Note the trailing period; tables.py is the only module that adds one.
		return "", excelerr.New(excelerr.ErrData, fmt.Sprintf("Sheet '%s' not found.", sheetName))
	}

	if tableName == "" {
		tableName = "Table_" + uuidHex8()
	}

	// PARITY(tables.py:create_excel_table): the duplicate check looks in
	// wb.defined_names, where openpyxl does not register tables, so it never
	// fires. The real duplicate rejection comes from add_table below.
	existing, err := f.GetTables(sheetName)
	if err != nil {
		return "", err
	}
	for _, t := range existing {
		// openpyxl's own ValueError, reached through add_table, carries no
		// trailing period — unlike the "Sheet ... not found." above.
		if t.Name == tableName {
			return "", excelerr.New(excelerr.ErrData,
				fmt.Sprintf("Table with name %s already exists", tableName))
		}
	}

	showRowStripes := true
	no := false
	if err := f.AddTable(sheetName, &excelize.Table{
		Name:              tableName,
		Range:             dataRange,
		StyleName:         tableStyle,
		ShowFirstColumn:   false,
		ShowLastColumn:    false,
		ShowRowStripes:    &showRowStripes,
		ShowColumnStripes: no,
	}); err != nil {
		return "", err
	}
	if err := f.Save(); err != nil {
		return "", err
	}

	return fmt.Sprintf("Successfully created table '%s' in sheet '%s'.", tableName, sheetName), nil
}

// uuidHex8 reproduces uuid.uuid4().hex[:8]. The harness normalizes these to
// {TABLE_ID} / {PIVOT_ID} (SPEC 6), so only the shape matters.
func uuidHex8() string {
	var b [4]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "00000000"
	}
	return hex.EncodeToString(b[:])
}
