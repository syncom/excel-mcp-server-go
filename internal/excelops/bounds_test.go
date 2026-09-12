package excelops

import (
	"errors"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// twoCellBook is the fixture the memory findings were reproduced against: a
// real workbook whose used range is tiny, so any large range named by a caller
// is entirely outside the data.
func twoCellBook(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	f := excelize.NewFile()
	defer f.Close()
	if err := f.SetCellValue("Sheet1", "A1", "x"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "B2", "y"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

// An end_cell naming the far corner of the grid must be refused before the
// range is materialized. Unbounded this allocates until the process dies, and
// a Go OOM is fatal rather than recoverable — the bound is the only defence.
func TestReadRejectsOversizedRange(t *testing.T) {
	path := twoCellBook(t)

	_, err := ReadExcelRangeWithMetadata(path, "Sheet1", "A1", "XFD1048576")
	if err == nil {
		t.Fatal("oversized read succeeded, want an error")
	}
	if !errors.Is(err, excelerr.ErrData) {
		t.Errorf("kind = %v, want ErrData — read_data_from_excel's routing depends on it", err)
	}
	if !strings.Contains(err.Error(), "exceeds the maximum") {
		t.Errorf("message = %q, want it to name the limit", err.Error())
	}
}

// The bound must not disturb ranges a caller can actually use, including ones
// reaching past the used range — Python returns those cells with a nil value
// and the port has to agree.
func TestReadKeepsRangesWithinBound(t *testing.T) {
	path := twoCellBook(t)

	result, err := ReadExcelRangeWithMetadata(path, "Sheet1", "A1", "D10")
	if err != nil {
		t.Fatalf("in-bounds read failed: %v", err)
	}
	cells, ok := result.Get("cells")
	if !ok {
		t.Fatal("no cells key in result")
	}
	list, isSlice := cells.([]any)
	if !isSlice {
		t.Fatalf("cells is %T, want []any", cells)
	}
	// 4 columns x 10 rows, counted past the 2x2 used range exactly as
	// openpyxl's ws.cell() does. Clamping to the used range here would return
	// 4 and silently change the wire contract.
	if len(list) != 40 {
		t.Errorf("len(cells) = %d, want 40", len(list))
	}
}

// The format path has the same unbounded loop and a worse failure mode: each
// iteration mutates workbook state, so an oversized range leaves a partially
// applied format behind on its way to dying.
func TestFormatRejectsOversizedRange(t *testing.T) {
	path := twoCellBook(t)

	err := FormatRange(path, "Sheet1", "A1", "XFD1048576", FormatOptions{Bold: true})
	if err == nil {
		t.Fatal("oversized format succeeded, want an error")
	}
	if !errors.Is(err, excelerr.ErrFormatting) {
		t.Errorf("kind = %v, want ErrFormatting so format_range reports it as text", err)
	}
	if !strings.Contains(err.Error(), "exceeds the maximum") {
		t.Errorf("message = %q, want it to name the limit", err.Error())
	}
}

// Formatting past the used range is legitimate — openpyxl creates those cells
// and the format persists — so the bound must leave it working.
func TestFormatKeepsRangesWithinBound(t *testing.T) {
	path := twoCellBook(t)

	if err := FormatRange(path, "Sheet1", "A1", "D10", FormatOptions{Bold: true}); err != nil {
		t.Fatalf("in-bounds format failed: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	id, err := f.GetCellStyle("Sheet1", "D10")
	if err != nil {
		t.Fatal(err)
	}
	style, err := f.GetStyle(id)
	if err != nil {
		t.Fatal(err)
	}
	if style.Font == nil || !style.Font.Bold {
		t.Error("D10 is not bold: formatting past the used range must still apply")
	}
}

// An end row past excelize's own grid limit cannot be rendered as an A1
// address. The bound still has to report the limit that stopped the request
// rather than leaking the naming failure.
func TestReadReportsBoundForUnnameableCell(t *testing.T) {
	path := twoCellBook(t)

	_, err := ReadExcelRangeWithMetadata(path, "Sheet1", "A1", "B99999999")
	if err == nil {
		t.Fatal("oversized read succeeded, want an error")
	}
	if !strings.Contains(err.Error(), "exceeds the maximum") {
		t.Errorf("message = %q, want it to name the range limit", err.Error())
	}
}

// checkRangeSize is the shared gate; the boundary itself is worth pinning
// directly so the two callers above cannot drift from it.
func TestCheckRangeSizeBoundary(t *testing.T) {
	cases := []struct {
		name                               string
		startRow, startCol, endRow, endCol int
		wantErr                            bool
	}{
		{name: "single cell", startRow: 1, startCol: 1, endRow: 1, endCol: 1},
		{name: "exactly at the limit", startRow: 1, startCol: 1, endRow: maxRangeCells, endCol: 1},
		{name: "one past the limit", startRow: 1, startCol: 1, endRow: maxRangeCells + 1, endCol: 1, wantErr: true},
		{name: "product overflows the limit", startRow: 1, startCol: 1, endRow: 1000, endCol: 1000, wantErr: true},
		{name: "offset start stays in bounds", startRow: 1000, startCol: 1, endRow: 1000 + maxRangeCells - 1, endCol: 1},
		{name: "inverted range", startRow: 10, startCol: 10, endRow: 1, endCol: 1},
		{name: "whole grid", startRow: 1, startCol: 1, endRow: 1048576, endCol: 16384, wantErr: true},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			err := checkRangeSize(excelerr.ErrData, tc.startRow, tc.startCol, tc.endRow, tc.endCol)
			if tc.wantErr && err == nil {
				t.Error("got nil, want an error")
			}
			if !tc.wantErr && err != nil {
				t.Errorf("got %v, want nil", err)
			}
		})
	}
}
