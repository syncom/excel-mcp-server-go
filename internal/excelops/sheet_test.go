package excelops

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestDeleteRangeShiftsWithinSpan is deviation D4's proof.
//
// It cannot be a conformance case: the whole point of D4 is that the two
// workbooks differ afterwards, and the harness asserts either equality or a
// pinned response string. D4's registry entry says so, and pins the Go side
// against a golden — which is what this test is.
//
// Python's delete_range_operation blanks A1:B2 and then calls delete_rows(1, 2),
// wiping rows 1 and 2 across *every* column, so C1:D2 would be lost and C3:D4
// would move up. Go must shift only within the range's own column span.
func TestDeleteRangeShiftsWithinSpan(t *testing.T) {
	path := filepath.Join(t.TempDir(), "d4.xlsx")
	if err := CreateWorkbook(path, DefaultSheetName); err != nil {
		t.Fatal(err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	for row := 1; row <= 4; row++ {
		for col, prefix := range []string{"a", "b", "KEEP-C", "KEEP-D"} {
			cell, _ := excelize.CoordinatesToCellName(col+1, row)
			if err := f.SetCellStr("Sheet1", cell, prefix+itoa(row)); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()

	msg, err := DeleteRangeOperation(path, "Sheet1", "A1", "B2", "up")
	if err != nil {
		t.Fatalf("DeleteRangeOperation failed: %v", err)
	}
	if want := "Range A1:B2 deleted successfully"; msg != want {
		t.Fatalf("message = %q, want %q", msg, want)
	}

	got, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer got.Close()

	want := map[string]string{
		// A1:B2 removed; A3:B4 shifted up into the gap.
		"A1": "a3", "B1": "b3",
		"A2": "a4", "B2": "b4",
		// The vacated tail of the range's own columns is empty.
		"A3": "", "B3": "", "A4": "", "B4": "",
		// Everything outside the range's column span is untouched. This is the
		// data Python destroys.
		"C1": "KEEP-C1", "D1": "KEEP-D1",
		"C2": "KEEP-C2", "D2": "KEEP-D2",
		"C3": "KEEP-C3", "D3": "KEEP-D3",
		"C4": "KEEP-C4", "D4": "KEEP-D4",
	}
	for cell, expected := range want {
		v, err := got.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Fatalf("reading %s: %v", cell, err)
		}
		if v != expected {
			t.Errorf("%s = %q, want %q", cell, v, expected)
		}
	}
}

// TestDeleteRangeShiftLeftStaysInRowSpan is the "left" half of D4.
func TestDeleteRangeShiftLeftStaysInRowSpan(t *testing.T) {
	path := filepath.Join(t.TempDir(), "d4left.xlsx")
	if err := CreateWorkbook(path, DefaultSheetName); err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	for row := 1; row <= 3; row++ {
		for col := 1; col <= 4; col++ {
			cell, _ := excelize.CoordinatesToCellName(col, row)
			if err := f.SetCellStr("Sheet1", cell, cellTag(col, row)); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()

	if _, err := DeleteRangeOperation(path, "Sheet1", "A1", "B1", "left"); err != nil {
		t.Fatalf("DeleteRangeOperation failed: %v", err)
	}

	got, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer got.Close()

	want := map[string]string{
		// Row 1 shifts left by the range's width; its tail empties.
		"A1": cellTag(3, 1), "B1": cellTag(4, 1), "C1": "", "D1": "",
		// Rows outside the range's row span are untouched.
		"A2": cellTag(1, 2), "B2": cellTag(2, 2), "C2": cellTag(3, 2), "D2": cellTag(4, 2),
		"A3": cellTag(1, 3), "D3": cellTag(4, 3),
	}
	for cell, expected := range want {
		v, err := got.GetCellValue("Sheet1", cell)
		if err != nil {
			t.Fatalf("reading %s: %v", cell, err)
		}
		if v != expected {
			t.Errorf("%s = %q, want %q", cell, v, expected)
		}
	}
}

func cellTag(col, row int) string {
	name, _ := excelize.ColumnNumberToName(col)
	return name + itoa(row)
}

func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	var b []byte
	for n > 0 {
		b = append([]byte{byte('0' + n%10)}, b...)
		n /= 10
	}
	return string(b)
}
