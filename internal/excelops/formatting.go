package excelops

import (
	"errors"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
)

// NormalizeColor implements deviation D5.
//
// DEVIATION(D5): Python prefixes "FF" unless the string already starts with
// "FF". An 8-digit ARGB that does not start with FF — "80FF0000", 50% alpha —
// becomes the 10-character "FF80FF0000", which openpyxl rejects outright, and a
// 6-digit RGB that happens to start with FF is stored with a different alpha
// than every other colour. We normalize once instead: strip a leading '#',
// uppercase, require exactly 6 or 8 hex digits, and left-pad 6 with FF alpha.
// SPEC 5.3.2 D5; registered in conformance/deviations.json.
func NormalizeColor(value string) (string, error) {
	v := strings.ToUpper(strings.TrimPrefix(value, "#"))
	if len(v) != 6 && len(v) != 8 {
		return "", excelerr.New(excelerr.ErrFormatting, fmt.Sprintf("Invalid color: %s", value))
	}
	for i := 0; i < len(v); i++ {
		c := v[i]
		if !(c >= '0' && c <= '9' || c >= 'A' && c <= 'F') {
			return "", excelerr.New(excelerr.ErrFormatting, fmt.Sprintf("Invalid color: %s", value))
		}
	}
	if len(v) == 6 {
		v = "FF" + v
	}
	return v, nil
}

// excelizeColor hands excelize the 6-digit RGB.
//
// excelize prepends "FF" to whatever it is given, so passing the normalized
// 8-digit ARGB produced a 10-character rgb="FFFF0000FF" that openpyxl refuses
// to read at all. That also means excelize always stores full opacity: an
// 8-digit colour with a non-FF alpha is accepted (which is D5's point — Python
// rejects it outright) but stored opaque. Recorded in JOURNAL and DEVIATIONS.md.
func excelizeColor(normalized string) string {
	return normalized[len(normalized)-6:]
}

// borderStyles maps openpyxl's border style names to excelize's numeric codes.
var borderStyles = map[string]int{
	"none": 0, "thin": 1, "medium": 2, "dashed": 3, "dotted": 4, "thick": 5,
	"double": 6, "hair": 7, "mediumDashed": 8, "dashDot": 9, "mediumDashDot": 10,
	"dashDotDot": 11, "mediumDashDotDot": 12, "slantDashDot": 13,
}

// FormatOptions mirrors formatting.format_range's signature. A nil pointer is
// Python's None, which is what decides whether a category is replaced at all.
type FormatOptions struct {
	Bold         bool
	Italic       bool
	Underline    bool
	FontSize     *int
	FontColor    *string
	BgColor      *string
	BorderStyle  *string
	BorderColor  *string
	NumberFormat *string
	Alignment    *string
	WrapText     bool
	MergeCells   bool
	Protection   map[string]any
	Conditional  map[string]any
}

// FormatRange ports formatting.format_range.
func FormatRange(path, sheetName, startCell, endCell string, o FormatOptions) error {
	if !ValidateCellReference(startCell) {
		return excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid start cell reference: %s", startCell))
	}
	if endCell != "" && !ValidateCellReference(endCell) {
		return excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid end cell reference: %s", endCell))
	}

	f, err := OpenOrCreateWorkbook(path)
	if err != nil {
		return excelerr.New(excelerr.ErrFormatting, pyStr(err))
	}
	defer f.Close()

	if !hasSheet(f, sheetName) {
		return excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Sheet '%s' not found", sheetName))
	}

	startRow, startCol, endRow, endCol, err := ParseCellRange(startCell, endCell)
	if err != nil {
		return excelerr.New(excelerr.ErrValidation, fmt.Sprintf("Invalid cell range: %s", err.Error()))
	}
	if endRow == 0 {
		endRow = startRow
	}
	if endCol == 0 {
		endCol = startCol
	}

	// openpyxl builds Font(...) unconditionally and assigns it to every cell,
	// so the font is *replaced* even when the caller set no font options.
	// ColorIndexed must be negative: excelize treats the zero value as "indexed
	// colour 0" and emits <color indexed="0"/>, which openpyxl reads back as
	// the palette's 00000000 where Python has no colour at all.
	font := &excelize.Font{Bold: o.Bold, Italic: o.Italic, ColorIndexed: -1}
	if o.Underline {
		font.Underline = "single"
	}
	if o.FontSize != nil {
		font.Size = float64(*o.FontSize)
	}
	if o.FontColor != nil {
		c, err := NormalizeColor(*o.FontColor)
		if err != nil {
			return err
		}
		font.Color = excelizeColor(c)
	}

	var fill *excelize.Fill
	if o.BgColor != nil {
		c, err := NormalizeColor(*o.BgColor)
		if err != nil {
			return err
		}
		fill = &excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{excelizeColor(c)}}
	}

	var borders []excelize.Border
	if o.BorderStyle != nil {
		colorText := "000000"
		if o.BorderColor != nil {
			colorText = *o.BorderColor
		}
		c, err := NormalizeColor(colorText)
		if err != nil {
			return err
		}
		style, ok := borderStyles[*o.BorderStyle]
		if !ok {
			return excelerr.New(excelerr.ErrFormatting,
				fmt.Sprintf("Invalid border settings: Value must be one of {%s}", *o.BorderStyle))
		}
		for _, side := range []string{"left", "right", "top", "bottom"} {
			borders = append(borders, excelize.Border{Type: side, Color: excelizeColor(c), Style: style})
		}
	}

	var align *excelize.Alignment
	if o.Alignment != nil || o.WrapText {
		align = &excelize.Alignment{Vertical: "center", WrapText: o.WrapText}
		if o.Alignment != nil {
			align.Horizontal = *o.Alignment
		}
	}

	var protect *excelize.Protection
	if o.Protection != nil {
		// Protection(**protection) fills unsupplied keys from openpyxl's own
		// defaults, which are locked=True, hidden=False — not the Go zero
		// values. Passing only {"hidden": true} still leaves the cell locked.
		protect = &excelize.Protection{Locked: true, Hidden: false}
		if v, ok := o.Protection["locked"].(bool); ok {
			protect.Locked = v
		}
		if v, ok := o.Protection["hidden"].(bool); ok {
			protect.Hidden = v
		}
	}

	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return excelerr.New(excelerr.ErrFormatting, pyStr(err))
			}
			if err := applyCellStyle(f, sheetName, cell, font, fill, borders, align, protect, o.NumberFormat); err != nil {
				return excelerr.New(excelerr.ErrFormatting, pyStr(err))
			}
		}
	}

	if o.MergeCells && endCell != "" {
		if err := f.MergeCell(sheetName, startCell, endCell); err != nil {
			return excelerr.New(excelerr.ErrFormatting,
				fmt.Sprintf("Failed to merge cells: %s", pyStr(err)))
		}
	}

	if o.Conditional != nil {
		rangeStr := startCell
		if endCell != "" {
			rangeStr = startCell + ":" + endCell
		}
		if err := applyConditionalFormat(f, sheetName, rangeStr, o.Conditional); err != nil {
			return err
		}
	}

	if err := f.Save(); err != nil {
		return excelerr.New(excelerr.ErrFormatting, pyStr(err))
	}
	return nil
}

// applyCellStyle replaces only the categories the caller specified, leaving the
// rest of the cell's existing style alone — which is what assigning to
// cell.font / cell.fill / ... one attribute at a time does in openpyxl.
func applyCellStyle(f *excelize.File, sheet, cell string, font *excelize.Font, fill *excelize.Fill,
	borders []excelize.Border, align *excelize.Alignment, protect *excelize.Protection, numFmt *string) error {

	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return err
	}
	style := &excelize.Style{}
	if styleID != 0 {
		existing, err := f.GetStyle(styleID)
		if err == nil && existing != nil {
			style = existing
		}
	}

	style.Font = font
	if fill != nil {
		style.Fill = *fill
	}
	if borders != nil {
		style.Border = borders
	}
	if align != nil {
		style.Alignment = align
	}
	if protect != nil {
		style.Protection = protect
	}
	if numFmt != nil {
		style.CustomNumFmt = numFmt
		style.NumFmt = 0
	}

	newID, err := f.NewStyle(style)
	if err != nil {
		return err
	}
	return f.SetCellStyle(sheet, cell, cell, newID)
}

// applyConditionalFormat ports format_range's conditional-formatting block,
// mapping openpyxl's five rule classes onto excelize's option struct.
func applyConditionalFormat(f *excelize.File, sheet, rangeStr string, cf map[string]any) error {
	var err error
	ruleType, _ := cf["type"].(string)
	if ruleType == "" {
		return excelerr.New(excelerr.ErrFormatting, "Conditional format type not specified")
	}
	params, _ := cf["params"].(map[string]any)
	if params == nil {
		params = map[string]any{}
	}

	opt := excelize.ConditionalFormatOptions{}

	// wrap reproduces format_range's inner try/except, which re-raises
	// everything from the rule construction and the add() call as
	// "Failed to apply conditional formatting: {e}". The two checks *outside*
	// that try — a missing type and a bad cell_is fill colour — are not wrapped.
	wrap := func(msg string) error {
		return excelerr.New(excelerr.ErrFormatting,
			fmt.Sprintf("Failed to apply conditional formatting: %s", msg))
	}

	switch ruleType {
	case "color_scale":
		// openpyxl's ColorScaleRule takes start/mid/end triples; a mid_type
		// makes it a 3-colour scale.
		opt.Type = "2_color_scale"
		// excelize insists on a Criteria for every type outside its
		// noCriteriaTypes list, and colour scales are not on it even though the
		// draw function ignores the value. Supply a valid placeholder.
		opt.Criteria = "equal to"
		opt.MinType = strParam(params, "start_type", "min")
		opt.MinValue = strParam(params, "start_value", "")
		if opt.MinColor, err = cfColor(params, "start_color", "FFFFFFFF"); err != nil {
			return wrap(err.Error())
		}
		opt.MaxType = strParam(params, "end_type", "max")
		opt.MaxValue = strParam(params, "end_value", "")
		if opt.MaxColor, err = cfColor(params, "end_color", "FF000000"); err != nil {
			return wrap(err.Error())
		}
		if _, ok := params["mid_type"]; ok {
			opt.Type = "3_color_scale"
			opt.MidType = strParam(params, "mid_type", "percentile")
			opt.MidValue = strParam(params, "mid_value", "50")
			if opt.MidColor, err = cfColor(params, "mid_color", "FFFFFF00"); err != nil {
				return wrap(err.Error())
			}
		}

	case "data_bar":
		opt.Type = "data_bar"
		opt.Criteria = "equal to" // see the colour-scale note above
		opt.MinType = strParam(params, "start_type", "min")
		opt.MinValue = strParam(params, "start_value", "")
		opt.MaxType = strParam(params, "end_type", "max")
		opt.MaxValue = strParam(params, "end_value", "")
		if opt.BarColor, err = cfColor(params, "color", "FF638EC6"); err != nil {
			return wrap(err.Error())
		}

	case "icon_set":
		opt.Type = "icon_set"
		opt.IconStyle = strParam(params, "icon_style", "3TrafficLights")
		opt.ReverseIcons, _ = params["reverse"].(bool)
		if v, ok := params["showValue"].(bool); ok {
			opt.IconsOnly = !v
		}

	case "formula":
		opt.Type = "formula"
		opt.Criteria = firstFormulaParam(params)
		id, err := cfFormatID(f, params)
		if err != nil {
			return err
		}
		opt.Format = id

	case "cell_is":
		opt.Type = "cell"
		operator := strParam(params, "operator", "")
		criteria, ok := cfCriteria[operator]
		if !ok {
			return wrap(fmt.Sprintf("Value must be one of {%s}", operator))
		}
		opt.Criteria = criteria
		formulas := formulaList(params)
		if len(formulas) > 0 {
			opt.Value = formulas[0]
		}
		if len(formulas) > 1 {
			// excelize expresses a two-sided criterion as "a,b".
			opt.Value = formulas[0] + "," + formulas[1]
		}
		id, err := cfFormatID(f, params)
		if err != nil {
			return err
		}
		opt.Format = id

	default:
		// Python raises this inside the wrapped try, so it comes back doubled.
		return wrap(fmt.Sprintf("Invalid conditional format type: %s", ruleType))
	}

	if v, ok := params["stopIfTrue"].(bool); ok {
		opt.StopIfTrue = v
	}

	if err := f.SetConditionalFormat(sheet, rangeStr, []excelize.ConditionalFormatOptions{opt}); err != nil {
		return wrap(pyStr(err))
	}
	return nil
}

// cfCriteria maps openpyxl's CellIsRule operator names to the criteria strings
// excelize accepts. It mirrors excelize's own (unexported) operatorType table;
// importing it is not an option, and a wrong key is rejected as
// "parameter is invalid" with no indication of which field was at fault.
var cfCriteria = map[string]string{
	"beginsWith":         "begins with",
	"between":            "between",
	"containsText":       "containing",
	"endsWith":           "ends with",
	"equal":              "equal to",
	"greaterThan":        "greater than",
	"greaterThanOrEqual": "greater than or equal to",
	"lessThan":           "less than",
	"lessThanOrEqual":    "less than or equal to",
	"notBetween":         "not between",
	"notContains":        "not containing",
	"notEqual":           "not equal to",
}

// cfFormatID builds the differential style a formula/cell_is rule applies.
//
// DEVIATION(D5): the fill colour goes through the same normalization as every
// other colour. Python special-cases it with the same FF-prefix rule it uses
// elsewhere. SPEC 5.3.2 D5.
func cfFormatID(f *excelize.File, params map[string]any) (*int, error) {
	fillParams, ok := params["fill"].(map[string]any)
	if !ok {
		return nil, nil
	}
	raw := "FFC7CE" // openpyxl's default light red
	if v, ok := fillParams["fgColor"].(string); ok && v != "" {
		raw = v
	}
	color, err := NormalizeColor(raw)
	if err != nil {
		return nil, excelerr.New(excelerr.ErrFormatting,
			fmt.Sprintf("Invalid conditional format fill color: %s", strings.TrimPrefix(err.Error(), "")))
	}
	// A conditional format's differential style lives in dxfs, not cellXfs.
	// NewStyle appends to the wrong list, leaving dxfs count="0" with a
	// dangling dxfId that openpyxl cannot resolve at all.
	id, nerr := f.NewConditionalStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{excelizeColor(color)}},
	})
	if nerr != nil {
		return nil, excelerr.New(excelerr.ErrFormatting, pyStr(nerr))
	}
	return &id, nil
}

func strParam(params map[string]any, key, fallback string) string {
	switch v := params[key].(type) {
	case string:
		return v
	case float64:
		return trimTrailingZeroFloat(v)
	case nil:
		return fallback
	default:
		return fmt.Sprint(v)
	}
}

func trimTrailingZeroFloat(v float64) string {
	s := fmt.Sprintf("%g", v)
	return s
}

// cfColor reads a colour parameter of a conditional-format rule. Python passes
// these straight to openpyxl without its FF-prefix step, so they are not part
// of D5 and are only uppercased for storage.
func cfColor(params map[string]any, key, fallback string) (string, error) {
	v, ok := params[key].(string)
	if !ok || v == "" {
		return fallback, nil
	}
	// openpyxl still requires a valid aRGB value here — it just does not apply
	// the FF-prefix step, so "#FFFFFF" and "GGGGGG" are both rejected.
	if !isARGBHex(v) {
		return "", errNotARGB
	}
	return strings.ToUpper(v), nil
}

// errNotARGB is openpyxl's Color validation message.
var errNotARGB = errors.New("Colors must be aRGB hex values")

func isARGBHex(v string) bool {
	if len(v) != 6 && len(v) != 8 {
		return false
	}
	for i := 0; i < len(v); i++ {
		c := v[i]
		if !(c >= '0' && c <= '9' || c >= 'A' && c <= 'F' || c >= 'a' && c <= 'f') {
			return false
		}
	}
	return true
}

func formulaList(params map[string]any) []string {
	raw, ok := params["formula"]
	if !ok {
		return nil
	}
	switch v := raw.(type) {
	case string:
		return []string{v}
	case []any:
		out := make([]string, 0, len(v))
		for _, item := range v {
			if s, ok := item.(string); ok {
				out = append(out, s)
			} else {
				out = append(out, fmt.Sprint(item))
			}
		}
		return out
	}
	return nil
}

func firstFormulaParam(params map[string]any) string {
	if list := formulaList(params); len(list) > 0 {
		return list[0]
	}
	return ""
}
