package excelops

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// sampleSource writes the four-row source table both pivot tests use.
func sampleSource(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "pivot.xlsx")
	if err := CreateWorkbook(path, DefaultSheetName); err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	rows := [][]any{
		{"Region", "Product", "Revenue"},
		{"North", "Widget", 10},
		{"South", "Gadget", 20},
		{"North", "Gadget", 30},
		{"South", "Widget", 40},
	}
	for r, row := range rows {
		for c, v := range row {
			cell, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellValue("Sheet1", cell, v); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()
	return path
}

func assertCells(t *testing.T, path, sheet string, want map[string]string) {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	for cell, expected := range want {
		got, err := f.GetCellValue(sheet, cell)
		if err != nil {
			t.Fatalf("reading %s!%s: %v", sheet, cell, err)
		}
		if got != expected {
			t.Errorf("%s!%s = %q, want %q", sheet, cell, got, expected)
		}
	}
}

// TestPivotCrossTab is deviation D3's proof.
//
// It cannot be a conformance case: Python accepts `columns` and ignores it, so
// both servers answer "Summary table created successfully" while their _pivot
// sheets differ by design. D3's registry entry says to pin the Go side against
// a golden, which is what this is.
func TestPivotCrossTab(t *testing.T) {
	path := sampleSource(t)

	msg, err := CreatePivotTable(path, "Sheet1", "A1:C5",
		[]string{"Region"}, []string{"Revenue"}, []string{"Product"}, "sum")
	if err != nil {
		t.Fatalf("CreatePivotTable failed: %v", err)
	}
	if want := "Summary table created successfully"; msg != want {
		t.Fatalf("message = %q, want %q", msg, want)
	}

	// One column per (column combination x value field), combinations sorted.
	// Python would have produced a single "Revenue (sum)" column totalling 40
	// and 60 — a plausible-looking answer to a question the caller did not ask.
	assertCells(t, path, "Sheet1_pivot", map[string]string{
		"A1": "Region",
		"B1": "Gadget - Revenue (sum)",
		"C1": "Widget - Revenue (sum)",
		"A2": "North", "B2": "30", "C2": "10",
		"A3": "South", "B3": "20", "C3": "40",
	})
}

// TestPivotMeanAlias is deviation D1's proof: "mean" aggregates as "average",
// and the header still echoes the value the caller passed.
func TestPivotMeanAlias(t *testing.T) {
	path := sampleSource(t)

	if _, err := CreatePivotTable(path, "Sheet1", "A1:C5",
		[]string{"Region"}, []string{"Revenue"}, nil, "mean"); err != nil {
		t.Fatalf("CreatePivotTable failed: %v", err)
	}

	assertCells(t, path, "Sheet1_pivot", map[string]string{
		"A1": "Region",
		"B1": "Revenue (mean)",
		"A2": "North", "B2": "20",
		"A3": "South", "B3": "30",
	})
}

// TestPivotRejectsOtherAggFuncs keeps D1 narrow: only "mean" is added.
func TestPivotRejectsOtherAggFuncs(t *testing.T) {
	path := sampleSource(t)
	for _, bad := range []string{"median", "avg", "bogus", ""} {
		t.Run(bad, func(t *testing.T) {
			_, err := CreatePivotTable(path, "Sheet1", "A1:C5",
				[]string{"Region"}, []string{"Revenue"}, nil, bad)
			if err == nil {
				t.Fatalf("agg_func %q was accepted", bad)
			}
			want := "Invalid aggregation function. Must be one of: sum, average, count, min, max"
			if err.Error() != want {
				t.Fatalf("error = %q, want %q", err, want)
			}
		})
	}
}
