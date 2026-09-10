package mcpserver

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
)

// Tier E — range and structure operations (SPEC 4).

type copyRangeIn struct {
	Filepath    string  `json:"filepath"`
	SheetName   string  `json:"sheet_name"`
	SourceStart string  `json:"source_start"`
	SourceEnd   string  `json:"source_end"`
	TargetStart string  `json:"target_start"`
	TargetSheet *string `json:"target_sheet,omitempty"`
}

type deleteRangeIn struct {
	Filepath       string `json:"filepath"`
	SheetName      string `json:"sheet_name"`
	StartCell      string `json:"start_cell"`
	EndCell        string `json:"end_cell"`
	ShiftDirection string `json:"shift_direction,omitempty"`
}

type rowsIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	StartRow  int    `json:"start_row"`
	Count     int    `json:"count,omitempty"`
}

type colsIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	StartCol  int    `json:"start_col"`
	Count     int    `json:"count,omitempty"`
}

func (s *Server) registerTierE(srv *mcp.Server) {
	yes := true
	// Every tool in this tier catches the same pair.
	caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "copy_range",
		Description: descCopyRange,
		Annotations: &mcp.ToolAnnotations{Title: "Copy Range", DestructiveHint: &yes},
		InputSchema: object("copy_rangeArguments",
			[]string{"filepath", "sheet_name", "source_start", "source_end", "target_start"},
			map[string]*prop{
				"filepath":     str("Filepath"),
				"sheet_name":   str("Sheet Name"),
				"source_start": str("Source Start"),
				"source_end":   str("Source End"),
				"target_start": str("Target Start"),
				"target_sheet": nullableStr("Target Sheet"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in copyRangeIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("copy_range", caught, "", err), nil, nil
		}
		// server.py passes `target_sheet or sheet_name`, so the impl never sees
		// a null target sheet.
		target := in.SheetName
		if in.TargetSheet != nil && *in.TargetSheet != "" {
			target = *in.TargetSheet
		}
		msg, err := excelops.CopyRangeOperation(full, in.SheetName, in.SourceStart, in.SourceEnd, in.TargetStart, target)
		return resultFor("copy_range", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "delete_range",
		Description: descDeleteRange,
		Annotations: &mcp.ToolAnnotations{Title: "Delete Range", DestructiveHint: &yes},
		InputSchema: object("delete_rangeArguments",
			[]string{"filepath", "sheet_name", "start_cell", "end_cell"}, map[string]*prop{
				"filepath":        str("Filepath"),
				"sheet_name":      str("Sheet Name"),
				"start_cell":      str("Start Cell"),
				"end_cell":        str("End Cell"),
				"shift_direction": withDefault(str("Shift Direction"), "up"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in deleteRangeIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("delete_range", caught, "", err), nil, nil
		}
		msg, err := excelops.DeleteRangeOperation(full, in.SheetName, in.StartCell, in.EndCell, in.ShiftDirection)
		return resultFor("delete_range", caught, msg, err), nil, nil
	})

	rowSchema := func(title string) *schema {
		return object(title, []string{"filepath", "sheet_name", "start_row"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
			"start_row":  &prop{Type: "integer", Title: "Start Row"},
			"count":      withDefault(&prop{Type: "integer", Title: "Count"}, 1),
		})
	}
	colSchema := func(title string) *schema {
		return object(title, []string{"filepath", "sheet_name", "start_col"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
			"start_col":  &prop{Type: "integer", Title: "Start Col"},
			"count":      withDefault(&prop{Type: "integer", Title: "Count"}, 1),
		})
	}

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "insert_rows",
		Description: descInsertRows,
		Annotations: &mcp.ToolAnnotations{Title: "Insert Rows", DestructiveHint: &yes},
		InputSchema: rowSchema("insert_rowsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in rowsIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("insert_rows", caught, "", err), nil, nil
		}
		msg, err := excelops.InsertRow(full, in.SheetName, in.StartRow, in.Count)
		return resultFor("insert_rows", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "insert_columns",
		Description: descInsertColumns,
		Annotations: &mcp.ToolAnnotations{Title: "Insert Columns", DestructiveHint: &yes},
		InputSchema: colSchema("insert_columnsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in colsIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("insert_columns", caught, "", err), nil, nil
		}
		msg, err := excelops.InsertCols(full, in.SheetName, in.StartCol, in.Count)
		return resultFor("insert_columns", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "delete_sheet_rows",
		Description: descDeleteSheetRows,
		Annotations: &mcp.ToolAnnotations{Title: "Delete Rows", DestructiveHint: &yes},
		InputSchema: rowSchema("delete_sheet_rowsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in rowsIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("delete_sheet_rows", caught, "", err), nil, nil
		}
		msg, err := excelops.DeleteRows(full, in.SheetName, in.StartRow, in.Count)
		return resultFor("delete_sheet_rows", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "delete_sheet_columns",
		Description: descDeleteSheetColumns,
		Annotations: &mcp.ToolAnnotations{Title: "Delete Columns", DestructiveHint: &yes},
		InputSchema: colSchema("delete_sheet_columnsArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in colsIn) (*mcp.CallToolResult, any, error) {
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("delete_sheet_columns", caught, "", err), nil, nil
		}
		msg, err := excelops.DeleteCols(full, in.SheetName, in.StartCol, in.Count)
		return resultFor("delete_sheet_columns", caught, msg, err), nil, nil
	})
}
