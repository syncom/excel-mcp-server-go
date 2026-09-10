package mcpserver

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// Tier D — formatting and merges (SPEC 4).

type formatRangeIn struct {
	Filepath          string         `json:"filepath"`
	SheetName         string         `json:"sheet_name"`
	StartCell         string         `json:"start_cell"`
	EndCell           *string        `json:"end_cell,omitempty"`
	Bold              bool           `json:"bold,omitempty"`
	Italic            bool           `json:"italic,omitempty"`
	Underline         bool           `json:"underline,omitempty"`
	FontSize          *int           `json:"font_size,omitempty"`
	FontColor         *string        `json:"font_color,omitempty"`
	BgColor           *string        `json:"bg_color,omitempty"`
	BorderStyle       *string        `json:"border_style,omitempty"`
	BorderColor       *string        `json:"border_color,omitempty"`
	NumberFormat      *string        `json:"number_format,omitempty"`
	Alignment         *string        `json:"alignment,omitempty"`
	WrapText          bool           `json:"wrap_text,omitempty"`
	MergeCells        bool           `json:"merge_cells,omitempty"`
	Protection        map[string]any `json:"protection,omitempty"`
	ConditionalFormat map[string]any `json:"conditional_format,omitempty"`
}

type mergeIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	StartCell string `json:"start_cell"`
	EndCell   string `json:"end_cell"`
}

type mergedCellsIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
}

func (s *Server) registerTierD(srv *mcp.Server) {
	yes := true

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "format_range",
		Description: descFormatRange,
		Annotations: &mcp.ToolAnnotations{Title: "Format Range", DestructiveHint: &yes},
		InputSchema: object("format_rangeArguments",
			[]string{"filepath", "sheet_name", "start_cell"}, map[string]*prop{
				"filepath":           str("Filepath"),
				"sheet_name":         str("Sheet Name"),
				"start_cell":         str("Start Cell"),
				"end_cell":           nullableStr("End Cell"),
				"bold":               withDefault(boolean("Bold"), false),
				"italic":             withDefault(boolean("Italic"), false),
				"underline":          withDefault(boolean("Underline"), false),
				"font_size":          nullableOf(&prop{Type: "integer"}, "Font Size"),
				"font_color":         nullableStr("Font Color"),
				"bg_color":           nullableStr("Bg Color"),
				"border_style":       nullableStr("Border Style"),
				"border_color":       nullableStr("Border Color"),
				"number_format":      nullableStr("Number Format"),
				"alignment":          nullableStr("Alignment"),
				"wrap_text":          withDefault(boolean("Wrap Text"), false),
				"merge_cells":        withDefault(boolean("Merge Cells"), false),
				"protection":         nullableOf(objectProp(), "Protection"),
				"conditional_format": nullableOf(objectProp(), "Conditional Format"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in formatRangeIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrFormatting}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("format_range", caught, "", err), nil, nil
		}
		endCell := ""
		if in.EndCell != nil {
			endCell = *in.EndCell
		}
		err = excelops.FormatRange(full, in.SheetName, in.StartCell, endCell, excelops.FormatOptions{
			Bold: in.Bold, Italic: in.Italic, Underline: in.Underline,
			FontSize: in.FontSize, FontColor: in.FontColor, BgColor: in.BgColor,
			BorderStyle: in.BorderStyle, BorderColor: in.BorderColor,
			NumberFormat: in.NumberFormat, Alignment: in.Alignment,
			WrapText: in.WrapText, MergeCells: in.MergeCells,
			Protection: in.Protection, Conditional: in.ConditionalFormat,
		})
		// server.py discards the impl's message and returns a constant.
		return resultFor("format_range", caught, "Range formatted successfully", err), nil, nil
	})

	mergeSchema := func(title string) *schema {
		return object(title, []string{"filepath", "sheet_name", "start_cell", "end_cell"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
			"start_cell": str("Start Cell"),
			"end_cell":   str("End Cell"),
		})
	}

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "merge_cells",
		Description: descMergeCells,
		Annotations: &mcp.ToolAnnotations{Title: "Merge Cells", DestructiveHint: &yes},
		InputSchema: mergeSchema("merge_cellsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in mergeIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("merge_cells", caught, "", err), nil, nil
		}
		msg, err := excelops.MergeRange(full, in.SheetName, in.StartCell, in.EndCell)
		return resultFor("merge_cells", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "unmerge_cells",
		Description: descUnmergeCells,
		Annotations: &mcp.ToolAnnotations{Title: "Unmerge Cells", DestructiveHint: &yes},
		InputSchema: mergeSchema("unmerge_cellsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in mergeIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("unmerge_cells", caught, "", err), nil, nil
		}
		msg, err := excelops.UnmergeRange(full, in.SheetName, in.StartCell, in.EndCell)
		return resultFor("unmerge_cells", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "get_merged_cells",
		Description: descGetMergedCells,
		Annotations: &mcp.ToolAnnotations{Title: "Get Merged Cells", ReadOnlyHint: true},
		InputSchema: object("get_merged_cellsArguments",
			[]string{"filepath", "sheet_name"}, map[string]*prop{
				"filepath":   str("Filepath"),
				"sheet_name": str("Sheet Name"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in mergedCellsIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("get_merged_cells", caught, "", err), nil, nil
		}
		ranges, err := excelops.GetMergedRanges(full, in.SheetName)
		if err != nil {
			return resultFor("get_merged_cells", caught, "", err), nil, nil
		}
		// server.py returns str(list): the Python repr (SPEC 5.4).
		return textResult(pyfmt.Repr(ranges)), nil, nil
	})
}
