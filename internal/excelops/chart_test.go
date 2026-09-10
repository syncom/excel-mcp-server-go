package excelops

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestChartAnchorsAtTargetCell is deviation D7's proof.
//
// It cannot be a conformance case. D7 changes where *every* chart is placed, so
// SPEC 7 stops comparing chart anchors altogether — which means the anchor has
// no differential coverage and has to be pinned here instead, exactly as D7's
// registry entry prescribes.
//
// Python would place all of these at E15, stacked on top of each other.
func TestChartAnchorsAtTargetCell(t *testing.T) {
	cases := []struct {
		target string
		want   [2]int // 0-based (col, row) as openpyxl reports the anchor
	}{
		{"A1", [2]int{0, 0}},
		{"E20", [2]int{4, 19}},
		{"N20", [2]int{13, 19}},
		// The multi-letter column Python's target_cell[0] parse cannot reach.
		{"AA10", [2]int{26, 9}},
		{"ZZ100", [2]int{701, 99}},
		// Below row 1: the writer's default is kept, so the call still succeeds
		// exactly as Python's does.
		{"E0", [2]int{4, 14}},
	}

	for _, tc := range cases {
		t.Run(tc.target, func(t *testing.T) {
			path := filepath.Join(t.TempDir(), "chart.xlsx")
			if err := CreateWorkbook(path, DefaultSheetName); err != nil {
				t.Fatal(err)
			}
			f, err := excelize.OpenFile(path)
			if err != nil {
				t.Fatal(err)
			}
			for r, row := range [][]any{{"K", "V"}, {"a", 1}, {"b", 2}} {
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

			msg, err := CreateChartInSheet(path, "Sheet1", "A1:B3", "bar", tc.target, "", "", "")
			if err != nil {
				t.Fatalf("CreateChartInSheet(%q) failed: %v", tc.target, err)
			}
			if want := "Bar chart created successfully"; msg != want {
				t.Fatalf("message = %q, want %q", msg, want)
			}

			got, err := chartAnchor(tc.target)
			if err != nil {
				t.Fatalf("chartAnchor(%q) failed: %v", tc.target, err)
			}
			wantCell, _ := excelize.CoordinatesToCellName(tc.want[0]+1, tc.want[1]+1)
			if got != wantCell {
				t.Fatalf("chartAnchor(%q) = %q, want %q", tc.target, got, wantCell)
			}
		})
	}
}

// TestChartAnchorRejectionsUnchanged keeps D7 narrow: every reference Python
// rejects must still be rejected here, with the same text.
func TestChartAnchorRejectionsUnchanged(t *testing.T) {
	cases := []struct{ target, want string }{
		{"1A", "Invalid target cell: '1' is not a valid column name. Column names are from A to ZZZ"},
		{"A1B2", "Invalid target cell: invalid literal for int() with base 10: '1B2'"},
	}
	for _, tc := range cases {
		t.Run(tc.target, func(t *testing.T) {
			got, err := chartAnchor(tc.target)
			if err == nil {
				t.Fatalf("chartAnchor(%q) = %q, want an error", tc.target, got)
			}
			if err.Error() != tc.want {
				t.Fatalf("chartAnchor(%q) error = %q, want %q", tc.target, err, tc.want)
			}
		})
	}
}
