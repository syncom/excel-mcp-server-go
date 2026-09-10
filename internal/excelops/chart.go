package excelops

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// chartTypes is chart.create_chart_in_sheet's supported set. The iteration
// order of the Python dict is its literal order, which the error message
// reproduces.
var chartTypeOrder = []string{"line", "bar", "pie", "scatter", "area"}

var chartTypes = map[string]excelize.ChartType{
	"line": excelize.Line,
	// openpyxl's BarChart defaults to type="col", a vertical bar chart.
	"bar":     excelize.Col,
	"pie":     excelize.Pie,
	"scatter": excelize.Scatter,
	"area":    excelize.Area,
}

// CreateChartInSheet ports chart.create_chart_in_sheet.
func CreateChartInSheet(path, sheetName, dataRange, chartType, targetCell, title, xAxis, yAxis string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrChart,
			fmt.Sprintf("Unexpected error creating chart: %s", pyStr(err)))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	cellRange := dataRange
	if strings.Contains(dataRange, "!") {
		// chart.py:88 unpacks this split outside any try, so a second "!"
		// escapes to the outer handler rather than becoming a ValidationError.
		before, after, serr := pySplit2(dataRange, "!")
		if serr != nil {
			return "", excelerr.New(excelerr.ErrChart,
				fmt.Sprintf("Unexpected error creating chart: %s", serr.Error()))
		}
		if !hasSheet(f, before) {
			return "", excelerr.New(excelerr.ErrValidation,
				fmt.Sprintf("Sheet '%s' referenced in data range not found", before))
		}
		cellRange = after
	}

	// chart.py:97 splits inside a try that catches ValueError, so a range with
	// no colon (or two) is an "Invalid data range format" error, not a chart
	// with no series.
	startCell, endCell, serr := pySplit2(cellRange, ":")
	if serr != nil {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid data range format: %s", serr.Error()))
	}
	startRow, startCol, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid data range format: %s", err.Error()))
	}

	lower := strings.ToLower(chartType)
	kind, ok := chartTypes[lower]
	if !ok {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Unsupported chart type: %s. Supported types: %s",
				chartType, strings.Join(chartTypeOrder, ", ")))
	}

	// PARITY(chart.py:create_chart_in_sheet): the anchor uses only the FIRST
	// character of target_cell for the column and int() on everything after it
	// for the row, so a two-letter column reference like "AA10" reaches
	// int("A10") and fails with CPython's own message.
	if targetCell == "" || !strings.ContainsFunc(targetCell, isAlpha) || !strings.ContainsFunc(targetCell, isDigit) {
		// The ValidationError is raised inside the drawing try, whose
		// `except Exception` re-wraps it, so the message comes back doubled.
		return "", excelerr.New(excelerr.ErrChart,
			fmt.Sprintf("Failed to create chart drawing: Invalid target cell format: %s", targetCell))
	}
	// DEVIATION(D7): Python builds the anchor from target_cell, attaches it to a
	// SpreadsheetDrawing, and appends that to worksheet._drawings — but
	// openpyxl's writer serializes worksheet._charts and uses each chart's own
	// anchor, ChartBase's default "E15". The drawing is discarded, so every
	// chart lands at E15 whatever the caller asked for, and target_cell can
	// only reject a call, never place a chart. We anchor where the caller asked.
	// SPEC 5.3.2 D7; registered in conformance/deviations.json.
	anchor, aerr := chartAnchor(targetCell)
	if aerr != nil {
		return "", excelerr.New(excelerr.ErrValidation, aerr.Error())
	}

	series, err := chartSeries(sheetName, lower, startRow, startCol, endRow, endCol)
	if err != nil {
		return "", excelerr.New(excelerr.ErrChart,
			fmt.Sprintf("Failed to create chart data references: %s", pyStr(err)))
	}

	show := true
	chart := &excelize.Chart{
		Type:   kind,
		Series: series,
		Title:  excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: title}}},
		Legend: excelize.ChartLegend{Position: "right", ShowLegendKey: false},
		XAxis:  excelize.ChartAxis{Title: excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: xAxis}}}},
		YAxis:  excelize.ChartAxis{Title: excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: yAxis}}}},
		// chart.width = 15, chart.height = 7.5 (centimetres in openpyxl).
		Dimension: excelize.ChartDimension{
			Width:  uint(cmToPixels(15)),
			Height: uint(cmToPixels(7.5)),
		},
		PlotArea: excelize.ChartPlotArea{ShowVal: show},
	}

	if err := f.AddChart(sheetName, anchor, chart); err != nil {
		return "", excelerr.New(excelerr.ErrChart,
			fmt.Sprintf("Failed to create chart drawing: %s", pyStr(err)))
	}
	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrChart,
			fmt.Sprintf("Failed to save workbook with chart: %s", pyStr(err)))
	}

	return fmt.Sprintf("%s chart created successfully", pyCapitalize(chartType)), nil
}

// chartSeries reproduces the Reference layout of chart.py.
//
// For every type but scatter: one series per column after the first, values
// running from the row below the header to the end, the title taken from the
// header cell, and categories from the first column. Scatter is different, and
// deliberately so — see below.
func chartSeries(sheet, kind string, startRow, startCol, endRow, endCol int) ([]excelize.ChartSeries, error) {
	// openpyxl quotes the sheet name in every reference, and writes the series
	// title reference *relatively* ('Sheet1'!B1) while data and category
	// references are absolute ('Sheet1'!$B$2:$B$5).
	quoted := "'" + strings.ReplaceAll(sheet, "'", "''") + "'"
	ref := func(col, from, to int) (string, error) {
		name, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return "", err
		}
		return fmt.Sprintf("%s!$%s$%d:$%s$%d", quoted, name, from, name, to), nil
	}
	cellRef := func(col, row int) (string, error) {
		name, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return "", err
		}
		return fmt.Sprintf("%s!%s%d", quoted, name, row), nil
	}

	var out []excelize.ChartSeries
	if kind == "scatter" {
		// PARITY(chart.py:create_chart_in_sheet): the scatter branch builds
		// y_values starting at start_row+1 and then passes title_from_data=True,
		// which consumes that first cell as the series title and advances the
		// values by one more row. So a scatter series covers start_row+2..end_row,
		// one row shorter than every other chart type's.
		cats, err := ref(startCol, startRow+1, endRow)
		if err != nil {
			return nil, err
		}
		for col := startCol + 1; col <= endCol; col++ {
			name, err := cellRef(col, startRow+1)
			if err != nil {
				return nil, err
			}
			values, err := ref(col, startRow+2, endRow)
			if err != nil {
				return nil, err
			}
			out = append(out, excelize.ChartSeries{Name: name, Categories: cats, Values: values})
		}
		return out, nil
	}

	cats, err := ref(startCol, startRow+1, endRow)
	if err != nil {
		return nil, err
	}
	for col := startCol + 1; col <= endCol; col++ {
		name, err := cellRef(col, startRow)
		if err != nil {
			return nil, err
		}
		values, err := ref(col, startRow+1, endRow)
		if err != nil {
			return nil, err
		}
		out = append(out, excelize.ChartSeries{Name: name, Categories: cats, Values: values})
	}
	return out, nil
}

// cmToPixels converts openpyxl's centimetre chart size to the pixel dimension
// excelize expects, at the 96 DPI both writers assume.
func cmToPixels(cm float64) int {
	return int(cm / 2.54 * 96)
}

// pyCapitalize ports str.capitalize: first character upper, the rest lower.
func pyCapitalize(s string) string {
	if s == "" {
		return s
	}
	runes := []rune(s)
	return strings.ToUpper(string(runes[0])) + strings.ToLower(string(runes[1:]))
}

func isAlpha(r rune) bool {
	return r >= 'a' && r <= 'z' || r >= 'A' && r <= 'Z'
}

func isDigit(r rune) bool { return r >= '0' && r <= '9' }

// defaultChartAnchor is ChartBase's own default, which is where openpyxl puts
// every chart it writes.
const defaultChartAnchor = "E15"

// chartAnchor resolves target_cell to the cell the chart is anchored at.
//
// DEVIATION(D7), and bounded by its registry entry:
//   - A well-formed A1 reference is honoured in full, including the multi-letter
//     columns Python's target_cell[0] parse cannot reach.
//   - Anything else falls back to Python's own computation — first character as
//     the column, int() of the remainder as the row — so the error text and the
//     set of accepted inputs stay byte-identical to the oracle's for every input
//     Python also rejects.
//   - A reference resolving below row or column 1 keeps the writer's default, so
//     it still succeeds exactly as Python does.
func chartAnchor(targetCell string) (string, error) {
	if ValidateCellReference(targetCell) {
		if row, col, _, _, err := ParseCellRange(targetCell, ""); err == nil {
			if row >= 1 && col >= 1 {
				return excelize.CoordinatesToCellName(col, row)
			}
			return defaultChartAnchor, nil
		}
	}

	runes := []rune(targetCell)
	if _, err := pyColumnIndex(string(runes[0])); err != nil {
		return "", fmt.Errorf("Invalid target cell: %s", err.Error())
	}
	rest := string(runes[1:])
	// Python's int() tolerates surrounding whitespace and quotes the raw
	// argument when it fails.
	row, err := strconv.Atoi(strings.TrimSpace(rest))
	if err != nil {
		return "", fmt.Errorf("Invalid target cell: invalid literal for int() with base 10: '%s'", rest)
	}
	if row < 1 {
		return defaultChartAnchor, nil
	}
	col, err := pyColumnIndex(string(runes[0]))
	if err != nil {
		return "", fmt.Errorf("Invalid target cell: %s", err.Error())
	}
	return excelize.CoordinatesToCellName(col, row)
}
