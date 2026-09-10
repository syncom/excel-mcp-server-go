package mcpserver

import (
	"context"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/excelerr"
	"github.com/syncom/excel-mcp-server-go/internal/excelops"
	"github.com/syncom/excel-mcp-server-go/internal/pyfmt"
)

// Tier A — workbook and sheet lifecycle (SPEC 4).

type createWorkbookIn struct {
	Filepath string `json:"filepath"`
}

type createWorksheetIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
}

type getWorkbookMetadataIn struct {
	Filepath      string `json:"filepath"`
	IncludeRanges bool   `json:"include_ranges,omitempty"`
}

type copyWorksheetIn struct {
	Filepath    string `json:"filepath"`
	SourceSheet string `json:"source_sheet"`
	TargetSheet string `json:"target_sheet"`
}

type deleteWorksheetIn struct {
	Filepath  string `json:"filepath"`
	SheetName string `json:"sheet_name"`
}

type renameWorksheetIn struct {
	Filepath string `json:"filepath"`
	OldName  string `json:"old_name"`
	NewName  string `json:"new_name"`
}

func (s *Server) registerTierA(srv *mcp.Server) {
	yes := true

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "create_workbook",
		Description: descCreateWorkbook,
		Annotations: &mcp.ToolAnnotations{Title: "Create Workbook", DestructiveHint: &yes},
		InputSchema: object("create_workbookArguments", []string{"filepath"}, map[string]*prop{
			"filepath": str("Filepath"),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in createWorkbookIn) (*mcp.CallToolResult, any, error) {
		// except WorkbookError -> "Error: ..."; everything else re-raises.
		caught := []error{excelerr.ErrWorkbook}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("create_workbook", caught, "", err), nil, nil
		}
		if err := excelops.CreateWorkbook(full, excelops.DefaultSheetName); err != nil {
			return resultFor("create_workbook", caught, "", err), nil, nil
		}
		return textResult("Created workbook at " + full), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "create_worksheet",
		Description: descCreateWorksheet,
		Annotations: &mcp.ToolAnnotations{Title: "Create Worksheet", DestructiveHint: &yes},
		InputSchema: object("create_worksheetArguments", []string{"filepath", "sheet_name"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in createWorksheetIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrWorkbook}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("create_worksheet", caught, "", err), nil, nil
		}
		msg, err := excelops.CreateSheet(full, in.SheetName)
		return resultFor("create_worksheet", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "get_workbook_metadata",
		Description: descGetWorkbookMetadata,
		Annotations: &mcp.ToolAnnotations{Title: "Get Workbook Metadata", ReadOnlyHint: true},
		InputSchema: object("get_workbook_metadataArguments", []string{"filepath"}, map[string]*prop{
			"filepath":       str("Filepath"),
			"include_ranges": withDefault(boolean("Include Ranges"), false),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in getWorkbookMetadataIn) (*mcp.CallToolResult, any, error) {
		// Only WorkbookError is caught here (SPEC 5.2).
		caught := []error{excelerr.ErrWorkbook}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("get_workbook_metadata", caught, "", err), nil, nil
		}
		info, err := excelops.GetWorkbookInfo(full, in.IncludeRanges)
		if err != nil {
			return resultFor("get_workbook_metadata", caught, "", err), nil, nil
		}
		// server.py returns str(result): the dict's Python repr (SPEC 5.4).
		return textResult(pyfmt.Repr(info)), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "copy_worksheet",
		Description: descCopyWorksheet,
		Annotations: &mcp.ToolAnnotations{Title: "Copy Worksheet", DestructiveHint: &yes},
		InputSchema: object("copy_worksheetArguments", []string{"filepath", "source_sheet", "target_sheet"}, map[string]*prop{
			"filepath":     str("Filepath"),
			"source_sheet": str("Source Sheet"),
			"target_sheet": str("Target Sheet"),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in copyWorksheetIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("copy_worksheet", caught, "", err), nil, nil
		}
		msg, err := excelops.CopySheet(full, in.SourceSheet, in.TargetSheet)
		return resultFor("copy_worksheet", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "delete_worksheet",
		Description: descDeleteWorksheet,
		Annotations: &mcp.ToolAnnotations{Title: "Delete Worksheet", DestructiveHint: &yes},
		InputSchema: object("delete_worksheetArguments", []string{"filepath", "sheet_name"}, map[string]*prop{
			"filepath":   str("Filepath"),
			"sheet_name": str("Sheet Name"),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in deleteWorksheetIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("delete_worksheet", caught, "", err), nil, nil
		}
		msg, err := excelops.DeleteSheet(full, in.SheetName)
		return resultFor("delete_worksheet", caught, msg, err), nil, nil
	})

	mcp.AddTool(srv, &mcp.Tool{
		Name:        "rename_worksheet",
		Description: descRenameWorksheet,
		Annotations: &mcp.ToolAnnotations{Title: "Rename Worksheet", DestructiveHint: &yes},
		InputSchema: object("rename_worksheetArguments", []string{"filepath", "old_name", "new_name"}, map[string]*prop{
			"filepath": str("Filepath"),
			"old_name": str("Old Name"),
			"new_name": str("New Name"),
		}),
	}, func(ctx context.Context, _ *mcp.CallToolRequest, in renameWorksheetIn) (*mcp.CallToolResult, any, error) {
		caught := []error{excelerr.ErrValidation, excelerr.ErrSheet}
		full, err := s.ExcelPath(in.Filepath)
		if err != nil {
			return resultFor("rename_worksheet", caught, "", err), nil, nil
		}
		msg, err := excelops.RenameSheet(full, in.OldName, in.NewName)
		return resultFor("rename_worksheet", caught, msg, err), nil, nil
	})
}
