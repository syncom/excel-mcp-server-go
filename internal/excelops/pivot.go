package excelops

import (
	"fmt"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

var validAggFuncs = []string{"sum", "average", "count", "min", "max"}

// aggSuffixes are the suffixes clean_field_name strips.
var aggSuffixes = []string{" (sum)", " (average)", " (count)", " (min)", " (max)"}

// cleanFieldName ports pivot.clean_field_name.
func cleanFieldName(field string) string {
	field = strings.TrimSpace(field)
	lower := strings.ToLower(field)
	for _, suffix := range aggSuffixes {
		if strings.HasSuffix(lower, suffix) {
			return field[:len(field)-len(suffix)]
		}
	}
	return field
}

// CreatePivotTable ports pivot.create_pivot_table, including D1 and D3.
func CreatePivotTable(path, sheetName, dataRange string, rows, values, columns []string, aggFunc string) (string, error) {
	f, err := openWorkbook(path)
	if err != nil {
		return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return "", excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	if !strings.Contains(dataRange, ":") {
		return "", excelerr.New(excelerr.ErrValidation, "Data range must be in format 'A1:B2'")
	}
	startCell, endCell, err := pySplit2(dataRange, ":")
	if err != nil {
		// pivot.py:51 splits inside a try that catches ValueError.
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid data range format: %s", err.Error()))
	}
	_, _, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid data range format: %s", err.Error()))
	}
	if endRow == 0 || endCol == 0 {
		return "", excelerr.New(excelerr.ErrValidation,
			"Invalid data range format: missing end coordinates")
	}

	dataAsList, err := ReadExcelRange(path, sheetName, startCell, endCell)
	if err != nil || len(dataAsList) < 2 {
		reason := "Source data must have a header row and at least one data row."
		if err != nil {
			reason = err.Error()
		}
		return "", excelerr.New(excelerr.ErrPivot,
			fmt.Sprintf("Failed to read or process source data: %s", reason))
	}

	headers := make([]string, len(dataAsList[0]))
	for i, h := range dataAsList[0] {
		headers[i] = pyfmt.StrOf(h)
	}
	records := make([]map[string]any, 0, len(dataAsList)-1)
	for _, row := range dataAsList[1:] {
		rec := map[string]any{}
		for i := 0; i < len(headers) && i < len(row); i++ {
			rec[headers[i]] = row[i]
		}
		records = append(records, rec)
	}

	// DEVIATION(D1): "mean" is accepted as an alias for "average". The schema
	// default stays "mean", so tool metadata is unchanged, and the header cell
	// still echoes the value the caller passed. Every other invalid value still
	// produces the Python error verbatim, with the unchanged five-name list.
	// SPEC 5.3.2 D1; registered in conformance/deviations.json.
	// PARITY(pivot.py:create_pivot_table): agg_func is lower-cased for the
	// validation check only — line 86 tests agg_func.lower(), but line 169
	// hands _aggregate_values the *raw* string, which compares case-sensitively
	// and falls through to `else: return sum(values)`. So "COUNT" validates and
	// then sums. Only the D1 alias is resolved here.
	if !contains(validAggFuncs, strings.ToLower(aggFunc)) && strings.ToLower(aggFunc) != "mean" {
		return "", excelerr.New(excelerr.ErrValidation,
			fmt.Sprintf("Invalid aggregation function. Must be one of: %s", strings.Join(validAggFuncs, ", ")))
	}
	effectiveAgg := aggFunc
	if strings.EqualFold(aggFunc, "mean") {
		// DEVIATION(D1): the alias resolves for aggregation as well as
		// validation, which is what makes the declared default usable.
		effectiveAgg = "average"
	}

	available := map[string]bool{}
	for _, h := range headers {
		available[strings.ToLower(cleanFieldName(h))] = true
	}
	// available_fields_raw is data[0].keys(), a dict view — duplicate header
	// names collapse to one entry.
	seenHeader := map[string]bool{}
	var sortedHeaders []string
	for _, h := range headers {
		if !seenHeader[h] {
			seenHeader[h] = true
			sortedHeaders = append(sortedHeaders, h)
		}
	}
	sort.Strings(sortedHeaders)

	for _, spec := range []struct {
		fields []string
		kind   string
	}{{rows, "row"}, {values, "value"}} {
		for _, field := range spec.fields {
			if !available[strings.ToLower(cleanFieldName(field))] {
				return "", excelerr.New(excelerr.ErrValidation,
					fmt.Sprintf("Invalid %s field '%s'. Available fields: %s",
						spec.kind, field, strings.Join(sortedHeaders, ", ")))
			}
		}
	}
	for _, field := range columns {
		if !available[strings.ToLower(cleanFieldName(field))] {
			return "", excelerr.New(excelerr.ErrValidation,
				fmt.Sprintf("Invalid column field '%s'. Available fields: %s",
					field, strings.Join(sortedHeaders, ", ")))
		}
	}

	cleanedRows := mapFields(rows)
	cleanedValues := mapFields(values)
	cleanedColumns := mapFields(columns)

	pivotSheet := sheetName + "_pivot"
	if hasSheet(f, pivotSheet) {
		if err := f.DeleteSheet(pivotSheet); err != nil {
			return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
		}
	}
	if _, err := f.NewSheet(pivotSheet); err != nil {
		return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
	}

	boldStyle, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true, ColorIndexed: -1}})
	if err != nil {
		return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
	}

	rowCombos, err := combinationsFor(records, cleanedRows)
	if err != nil {
		return "", err
	}

	// DEVIATION(D3): Python accepts `columns`, advertises it in the schema, and
	// ignores it — a caller asking to break out by region gets a row-only
	// group-by that looks like a valid answer. We build column combinations the
	// same way rows are built and emit one column per (combination x value
	// field). When `columns` is empty the output is byte-identical to Python's.
	// SPEC 5.3.2 D3; registered in conformance/deviations.json.
	colCombos, err := combinationsFor(records, cleanedColumns)
	if err != nil {
		return "", err
	}

	// Header row.
	col := 1
	for _, field := range cleanedRows {
		if err := writeHeader(f, pivotSheet, col, 1, field, boldStyle); err != nil {
			return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
		}
		col++
	}
	if len(cleanedColumns) == 0 {
		for _, field := range cleanedValues {
			if err := writeHeader(f, pivotSheet, col, 1, fmt.Sprintf("%s (%s)", field, aggFunc), boldStyle); err != nil {
				return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
			}
			col++
		}
	} else {
		for _, combo := range colCombos {
			label := comboLabel(combo, cleanedColumns)
			for _, field := range cleanedValues {
				header := fmt.Sprintf("%s - %s (%s)", label, field, aggFunc)
				if err := writeHeader(f, pivotSheet, col, 1, header, boldStyle); err != nil {
					return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
				}
				col++
			}
		}
	}
	totalCols := col - 1

	// Data rows.
	rowIdx := 2
	for _, combo := range rowCombos {
		col = 1
		for _, field := range cleanedRows {
			cell, _ := excelize.CoordinatesToCellName(col, rowIdx)
			if err := f.SetCellStr(pivotSheet, cell, combo[field]); err != nil {
				return "", excelerr.New(excelerr.ErrPivot, pyStr(err))
			}
			col++
		}

		writeAgg := func(colFilter map[string]string) error {
			filtered := filterData(records, combo, colFilter)
			for _, valueField := range cleanedValues {
				agg := aggregateValues(filtered, valueField, effectiveAgg)
				cell, _ := excelize.CoordinatesToCellName(col, rowIdx)
				if err := setAggCell(f, pivotSheet, cell, agg); err != nil {
					return err
				}
				col++
			}
			return nil
		}

		if len(cleanedColumns) == 0 {
			if err := writeAgg(nil); err != nil {
				return "", excelerr.New(excelerr.ErrPivot,
					fmt.Sprintf("Failed to aggregate values for field '%s': %s", cleanedValues[0], pyStr(err)))
			}
		} else {
			for _, colCombo := range colCombos {
				if err := writeAgg(colCombo); err != nil {
					return "", excelerr.New(excelerr.ErrPivot,
						fmt.Sprintf("Failed to aggregate values for field '%s': %s", cleanedValues[0], pyStr(err)))
				}
			}
		}
		rowIdx++
	}

	totalRows := len(rowCombos) + 1
	lastCol, err := excelize.ColumnNumberToName(totalCols)
	if err != nil {
		return "", excelerr.New(excelerr.ErrPivot,
			fmt.Sprintf("Failed to create pivot table formatting: %s", pyStr(err)))
	}
	yes := true
	if err := f.AddTable(pivotSheet, &excelize.Table{
		Name:              "PivotTable_" + uuidHex8(),
		Range:             fmt.Sprintf("A1:%s%d", lastCol, totalRows),
		StyleName:         "TableStyleMedium9",
		ShowFirstColumn:   false,
		ShowLastColumn:    false,
		ShowRowStripes:    &yes,
		ShowColumnStripes: true,
	}); err != nil {
		return "", excelerr.New(excelerr.ErrPivot,
			fmt.Sprintf("Failed to create pivot table formatting: %s", pyStr(err)))
	}

	if err := f.Save(); err != nil {
		return "", excelerr.New(excelerr.ErrPivot,
			fmt.Sprintf("Failed to save workbook: %s", pyStr(err)))
	}
	return "Summary table created successfully", nil
}

func writeHeader(f *excelize.File, sheet string, col, row int, text string, style int) error {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	if err := f.SetCellStr(sheet, cell, text); err != nil {
		return err
	}
	return f.SetCellStyle(sheet, cell, cell, style)
}

// setAggCell writes an aggregate, preserving Python's int-vs-float result type:
// sum of ints and count are ints, average is always a float.
func setAggCell(f *excelize.File, sheet, cell string, v any) error {
	switch x := v.(type) {
	case int:
		return f.SetCellInt(sheet, cell, int64(x))
	case float64:
		return f.SetCellDefault(sheet, cell, pyfmt.FormatG(x, 16))
	default:
		return f.SetCellDefault(sheet, cell, pyfmt.StrOf(v))
	}
}

// maxPivotCombinations bounds the cartesian product [combinationsFor] builds.
//
// DEVIATION(D8): the product is the count of distinct values multiplied across
// every grouped field, so three fields over high-cardinality columns is N^3 —
// and the fields that invite grouping (dates, IDs, names, emails) are exactly
// the ones whose values are all distinct. A 120-row sheet with 3 such row
// fields yields ~1.7M combinations and over 2 GB, killing the process; 700 rows
// with 2 fields "succeeds" after 21 s and 819 MB, which is its own denial of
// service. No hostile argument is needed for either — just an ordinary call on
// an ordinary spreadsheet.
//
// Python has the identical blowup and the port inherited it, so unlike the
// range bound this refusal has no counterpart in the oracle; it is a deliberate
// divergence. Documented in DEVIATIONS.md (D8); still to be registered in
// conformance/deviations.json.
//
// The limit is far past any pivot a person can read — a table with
// 50k grouped rows is not an answer to anything — and a refusal naming the
// cause is strictly better than a dead server.
const maxPivotCombinations = 50_000

// combinationsFor builds the sorted unique value sets for each field and then
// their cartesian product, matching _get_combinations.
//
// The product's size is computed from the per-field value counts *before* any
// combination is allocated, so an oversized request costs nothing beyond the
// scan it takes to count.
func combinationsFor(records []map[string]any, fields []string) ([]map[string]string, error) {
	if len(fields) == 0 {
		return nil, nil
	}

	uniqueByField := make([][]string, len(fields))
	total := int64(1)
	for i, field := range fields {
		seen := map[string]bool{}
		var uniq []string
		for _, rec := range records {
			// str(record.get(field, '')) — a missing field becomes "".
			// str(record.get(field, '')): the key exists whenever the header
			// does, so an empty cell yields str(None) == "None", not "".
			v := ""
			if raw, ok := rec[field]; ok {
				v = pyfmt.StrOf(raw)
			}
			if !seen[v] {
				seen[v] = true
				uniq = append(uniq, v)
			}
		}
		sort.Strings(uniq)
		uniqueByField[i] = uniq

		// Counted in int64 and checked every field, so the running product
		// cannot overflow before the limit catches it.
		total *= int64(len(uniq))
		if total > maxPivotCombinations {
			return nil, excelerr.New(excelerr.ErrPivot, fmt.Sprintf(
				"Grouping by %s produces more than %d combinations; "+
					"use fewer fields or fields with fewer distinct values",
				strings.Join(fields, ", "), maxPivotCombinations))
		}
	}

	result := []map[string]string{{}}
	for i, field := range fields {
		next := make([]map[string]string, 0, len(result)*len(uniqueByField[i]))
		for _, combo := range result {
			for _, v := range uniqueByField[i] {
				merged := make(map[string]string, len(combo)+1)
				for k, old := range combo {
					merged[k] = old
				}
				merged[field] = v
				next = append(next, merged)
			}
		}
		result = next
	}
	return result, nil
}

func comboLabel(combo map[string]string, fields []string) string {
	parts := make([]string, 0, len(fields))
	for _, field := range fields {
		parts = append(parts, combo[field])
	}
	return strings.Join(parts, " / ")
}

// filterData ports _filter_data.
//
// PARITY(pivot.py:_filter_data): the combination values are strings — they come
// from str(record.get(field, ”)) — while the records hold raw values, so the
// `record.get(field) != value` comparison never matches for a numeric field and
// its groups aggregate to 0. Reproduced.
func filterData(records []map[string]any, rowFilter, colFilter map[string]string) []map[string]any {
	var out []map[string]any
	for _, rec := range records {
		matches := true
		for _, filter := range []map[string]string{rowFilter, colFilter} {
			for field, want := range filter {
				raw, ok := rec[field]
				if !ok {
					matches = false
					break
				}
				s, isStr := raw.(string)
				if !isStr || s != want {
					matches = false
					break
				}
			}
			if !matches {
				break
			}
		}
		if matches {
			out = append(out, rec)
		}
	}
	return out
}

// aggregateValues ports _aggregate_values: only int and float cells count, and
// an empty selection yields the integer 0.
func aggregateValues(records []map[string]any, field, aggFunc string) any {
	var nums []any
	for _, rec := range records {
		if raw, ok := rec[field]; ok {
			switch raw.(type) {
			case int, float64:
				nums = append(nums, raw)
			}
		}
	}
	if len(nums) == 0 {
		return 0
	}

	allInt := true
	total := 0.0
	for _, n := range nums {
		if f, ok := n.(float64); ok {
			allInt = false
			total += f
		} else {
			total += float64(n.(int))
		}
	}

	// Case-sensitive, matching _aggregate_values' chain of `==` comparisons.
	// Anything that matches none of them — including "COUNT" — falls through to
	// the sum branch, which is Python's own default.
	switch aggFunc {
	case "count":
		return len(nums)
	case "average":
		// Python's `/` always produces a float.
		return total / float64(len(nums))
	case "min", "max":
		best := nums[0]
		bestF := numOf(best)
		for _, n := range nums[1:] {
			f := numOf(n)
			if (aggFunc == "min" && f < bestF) || (aggFunc == "max" && f > bestF) {
				best, bestF = n, f
			}
		}
		return best
	default: // sum, and _aggregate_values' own fallback
		if allInt {
			return int(total)
		}
		return total
	}
}

func numOf(v any) float64 {
	if f, ok := v.(float64); ok {
		return f
	}
	return float64(v.(int))
}

func mapFields(fields []string) []string {
	out := make([]string, 0, len(fields))
	for _, f := range fields {
		out = append(out, cleanFieldName(f))
	}
	return out
}

func contains(list []string, v string) bool {
	for _, item := range list {
		if item == v {
			return true
		}
	}
	return false
}
