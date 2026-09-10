package mcpserver

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
)

// Tier F — tables, charts and pivots (SPEC 4).

type createTableIn struct {
	Filepath   string  `json:"filepath"`
	SheetName  string  `json:"sheet_name"`
	DataRange  string  `json:"data_range"`
	TableName  *string `json:"table_name,omitempty"`
	TableStyle string  `json:"table_style,omitempty"`
}

type createChartIn struct {
	Filepath   string `json:"filepath"`
	SheetName  string `json:"sheet_name"`
	DataRange  string `json:"data_range"`
	ChartType  string `json:"chart_type"`
	TargetCell string `json:"target_cell"`
	Title      string `json:"title,omitempty"`
	XAxis      string `json:"x_axis,omitempty"`
	YAxis      string `json:"y_axis,omitempty"`
}

type createPivotIn struct {
	Filepath  string   `json:"filepath"`
	SheetName string   `json:"sheet_name"`
	DataRange string   `json:"data_range"`
	Rows      []string `json:"rows"`
	Values    []string `json:"values"`
	Columns   []string `json:"columns,omitempty"`
	AggFunc   string   `json:"agg_func,omitempty"`
}

func (s *Server) registerTierF(srv *mcp.Server) {
	yes := true
	strList := func(title string) *prop {
		return &prop{Type: "array", Items: &prop{Type: "string"}, Title: title}
	}

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "create_table",
		Description: descCreateTable,
		Annotations: &mcp.ToolAnnotations{Title: "Create Table", DestructiveHint: &yes},
		InputSchema: object("create_tableArguments",
			[]string{"filepath", "sheet_name", "data_range"}, map[string]*prop{
				"filepath":    str("Filepath"),
				"sheet_name":  str("Sheet Name"),
				"data_range":  str("Data Range"),
				"table_name":  nullableStr("Table Name"),
				"table_style": withDefault(str("Table Style"), "TableStyleMedium9"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in createTableIn) (*mcp.CallToolResult, any, error) {
		// create_table catches only DataError — but create_excel_table wraps
		// every exception in one, so in practice everything becomes text.
		caught := []error{excelerr.ErrData}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("create_table", caught, "", err), nil, nil
		}
		name := ""
		if in.TableName != nil {
			name = *in.TableName
		}
		// No coercion: the schema default supplies "TableStyleMedium9" when the
		// caller omits the field, and an explicit "" must reach the impl the
		// way it does in Python.
		msg, err := excelops.CreateExcelTable(full, in.SheetName, in.DataRange, name, in.TableStyle)
		return resultFor("create_table", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "create_chart",
		Description: descCreateChart,
		Annotations: &mcp.ToolAnnotations{Title: "Create Chart", DestructiveHint: &yes},
		InputSchema: object("create_chartArguments",
			[]string{"filepath", "sheet_name", "data_range", "chart_type", "target_cell"},
			map[string]*prop{
				"filepath":    str("Filepath"),
				"sheet_name":  str("Sheet Name"),
				"data_range":  str("Data Range"),
				"chart_type":  str("Chart Type"),
				"target_cell": str("Target Cell"),
				"title":       withDefault(str("Title"), ""),
				"x_axis":      withDefault(str("X Axis"), ""),
				"y_axis":      withDefault(str("Y Axis"), ""),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in createChartIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrChart}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("create_chart", caught, "", err), nil, nil
		}
		msg, err := excelops.CreateChartInSheet(full, in.SheetName, in.DataRange, in.ChartType,
			in.TargetCell, in.Title, in.XAxis, in.YAxis)
		return resultFor("create_chart", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "create_pivot_table",
		Description: descCreatePivotTable,
		Annotations: &mcp.ToolAnnotations{Title: "Create Pivot Table", DestructiveHint: &yes},
		InputSchema: object("create_pivot_tableArguments",
			[]string{"filepath", "sheet_name", "data_range", "rows", "values"},
			map[string]*prop{
				"filepath":   str("Filepath"),
				"sheet_name": str("Sheet Name"),
				"data_range": str("Data Range"),
				"rows":       strList("Rows"),
				"values":     strList("Values"),
				"columns":    nullableOf(&prop{Type: "array", Items: &prop{Type: "string"}}, "Columns"),
				// DEVIATION(D1): the schema default stays "mean"; D1 makes it
				// meaningful rather than removing it. SPEC 5.3.2 D1.
				"agg_func": withDefault(str("Agg Func"), "mean"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in createPivotIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrPivot}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("create_pivot_table", caught, "", err), nil, nil
		}
		// No coercion, for the same reason as create_table above: an explicit
		// agg_func of "" is invalid in Python and must stay invalid here.
		msg, err := excelops.CreatePivotTable(full, in.SheetName, in.DataRange,
			in.Rows, in.Values, in.Columns, in.AggFunc)
		return resultFor("create_pivot_table", caught, msg, err), nil, nil
	})
}
