package mcpserver

import (
	"context"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// Tier C — formulas and validation (SPEC 4).

type formulaIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	Cell      string `json:"cell"`
	Formula   string `json:"formula"`
}

type validateRangeIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
	StartCell string `json:"start_cell"`
	EndCell   string `json:"end_cell,omitempty"`
}

type validationInfoIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
}

func (s *Server) registerTierC(srv *mcp.Server) {
	yes := true

	formulaSchema := func(title string) *schema {
		return object(title, []string{"filepath", "sheet_name", "cell", "formula"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
			"cell":       str("Cell"),
			"formula":    str("Formula"),
		})
	}

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "apply_formula",
		Description: descApplyFormula,
		Annotations: &mcp.ToolAnnotations{Title: "Apply Formula", DestructiveHint: &yes},
		InputSchema: formulaSchema("apply_formulaArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in formulaIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrCalculation}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("apply_formula", caught, "", err), nil, nil
		}
		// server.py validates first, with the formula exactly as given — so a
		// formula missing its '=' is rejected here even though the apply step
		// would have added one.
		if _, err := excelops.ValidateFormulaInCell(full, in.SheetName, in.Cell, in.Formula); err != nil &&
			!excelops.ErrNoneResult(err) {
			return resultFor("apply_formula", caught, "", err), nil, nil
		}
		msg, err := excelops.ApplyFormula(full, in.SheetName, in.Cell, in.Formula)
		return resultFor("apply_formula", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "validate_formula_syntax",
		Description: descValidateFormulaSyntax,
		Annotations: &mcp.ToolAnnotations{Title: "Validate Formula Syntax", ReadOnlyHint: true},
		InputSchema: formulaSchema("validate_formula_syntaxArguments"),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in formulaIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrCalculation}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("validate_formula_syntax", caught, "", err), nil, nil
		}
		msg, err := excelops.ValidateFormulaInCell(full, in.SheetName, in.Cell, in.Formula)
		return resultFor("validate_formula_syntax", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "validate_excel_range",
		Description: descValidateExcelRange,
		Annotations: &mcp.ToolAnnotations{Title: "Validate Excel Range", ReadOnlyHint: true},
		InputSchema: object("validate_excel_rangeArguments",
			[]string{"filepath", "sheet_name", "start_cell"}, map[string]*prop{
				"filepath":   str("Filepath"),
				"sheet_name": str("Sheet Name"),
				"start_cell": str("Start Cell"),
				"end_cell":   nullableStr("End Cell"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in validateRangeIn) (*mcp.CallToolResult, any, error) {
		// Only ValidationError is caught (SPEC 5.2).
		caught := []error{excelerr.ErrValidation}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("validate_excel_range", caught, "", err), nil, nil
		}
		rangeStr := in.StartCell
		if in.EndCell != "" {
			rangeStr = in.StartCell + ":" + in.EndCell
		}
		msg, err := excelops.ValidateRangeInSheet(full, in.SheetName, rangeStr)
		return resultFor("validate_excel_range", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "get_data_validation_info",
		Description: descGetDataValidationInfo,
		Annotations: &mcp.ToolAnnotations{Title: "Get Data Validation Info", ReadOnlyHint: true},
		InputSchema: object("get_data_validation_infoArguments",
			[]string{"filepath", "sheet_name"}, map[string]*prop{
				"filepath":   str("Filepath"),
				"sheet_name": str("Sheet Name"),
			}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in validationInfoIn) (*mcp.CallToolResult, any, error) {
		// SPEC 5.2: nothing is caught, so every failure raises through. The
		// missing-sheet case is not an exception in Python — it is a plain
		// return of an "Error: ..." string.
		var caught []error
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("get_data_validation_info", caught, "", err), nil, nil
		}
		rules, sheetExists, err := excelops.ValidationRangesForSheet(full, in.SheetName)
		if err != nil {
			return resultFor("get_data_validation_info", caught, "", err), nil, nil
		}
		if !sheetExists {
			return textResult(fmt.Sprintf("Error: Sheet '%s' not found", in.SheetName)), nil, nil
		}
		if len(rules) == 0 {
			return textResult("No data validation rules found in this worksheet"), nil, nil
		}
		out := pyfmt.NewDict().
			Set("sheet_name", in.SheetName).
			Set("validation_rules", rules)
		return textResult(pyfmt.Dumps(out)), nil, nil
	})
}
