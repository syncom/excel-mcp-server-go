package mcpserver

import (
	"context"
	"encoding/json"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// Tier B — data I/O (SPEC 4).

type writeDataIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	// Cell values stay as raw JSON so an integer literal survives as an
	// integer: Python's json parser gives int for 10 and float for 10.0, and
	// openpyxl stores the difference (SPEC 5.5).
	Data      [][]json.RawMessage `json:"data"`
	StartCell string              `json:"start_cell,omitempty"`
}

type readDataIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	StartCell string `json:"start_cell,omitempty"`
	EndCell   string `json:"end_cell,omitempty"`
	// DEVIATION(D6): accepted and ignored, exactly as Python does. It stays in
	// the schema so existing callers keep validating; only its description
	// changes. SPEC 5.3.2 D6; registered in conformance/deviations.json.
	PreviewOnly bool `json:"preview_only,omitempty"`
}

func (s *Server) registerTierB(srv *mcp.Server) {
	yes := true

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "write_data_to_excel",
		Description: descWriteDataToExcel,
		Annotations: &mcp.ToolAnnotations{Title: "Write Data to Excel", DestructiveHint: &yes},
		InputSchema: object("write_data_to_excelArguments",
			[]string{"filepath", "sheet_name", "data"}, map[string]*prop{
				"filepath":   str("Filepath"),
				"sheet_name": str("Sheet Name"),
				"data": {
					Type:  "array",
					Title: "Data",
					// pydantic renders List[List] with an untyped inner items.
					Items: &prop{Type: "array", Items: &prop{}},
				},
				"start_cell": withDefault(str("Start Cell"), "A1"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in writeDataIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrData}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("write_data_to_excel", caught, "", err), nil, nil
		}
		msg, err := excelops.WriteData(full, in.SheetName, in.Data, in.StartCell)
		return resultFor("write_data_to_excel", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "read_data_from_excel",
		Description: descReadDataFromExcel,
		Annotations: &mcp.ToolAnnotations{Title: "Read Data from Excel", ReadOnlyHint: true},
		InputSchema: object("read_data_from_excelArguments",
			[]string{"filepath", "sheet_name"}, map[string]*prop{
				"filepath":     str("Filepath"),
				"sheet_name":   str("Sheet Name"),
				"start_cell":   withDefault(str("Start Cell"), "A1"),
				"end_cell":     nullableStr("End Cell"),
				"preview_only": withDefault(boolean("Preview Only"), false),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in readDataIn) (*mcp.CallToolResult, any, error) {
		// SPEC 5.2: this tool's only except clause re-raises, so every failure
		// is a protocol-level tool error and nothing becomes "Error: ..." text.
		var caught []error
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("read_data_from_excel", caught, "", err), nil, nil
		}
		result, err := excelops.ReadExcelRangeWithMetadata(full, in.SheetName, in.StartCell, in.EndCell)
		if err != nil {
			return resultFor("read_data_from_excel", caught, "", err), nil, nil
		}
		if cells, ok := result.Get("cells"); !ok || len(cells.([]any)) == 0 {
			return textResult("No data found in specified range"), nil, nil
		}
		return textResult(pyfmt.Dumps(result)), nil, nil
	})
}
