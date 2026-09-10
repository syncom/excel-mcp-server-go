package mcpserver

// Tool descriptions, copied verbatim from the Python docstrings as FastMCP
// cleans them (SPEC 5.6). They are transcribed from
// conformance/golden/tools_list_python.json, which is the frozen reference, and
// live in one block so a reviewer can diff them against it in one place
// (STANDARDS: MCP SDK usage).
//
// The only sanctioned difference is read_data_from_excel, per deviation D6.
const (
	descCreateWorkbook = "Create new Excel workbook."

	descCreateWorksheet = "Create new worksheet in workbook."

	descGetWorkbookMetadata = "Get metadata about workbook including sheets, ranges, etc."

	descCopyWorksheet = "Copy worksheet within workbook."

	descDeleteWorksheet = "Delete worksheet from workbook."

	descRenameWorksheet  = "Rename worksheet in workbook."
	descWriteDataToExcel = "\nWrite data to Excel worksheet.\nExcel formula will write to cell without any verification.\n\nPARAMETERS:  \nfilepath: Path to Excel file\nsheet_name: Name of worksheet to write to\ndata: List of lists containing data to write to the worksheet, sublists are assumed to be rows\nstart_cell: Cell to start writing to, default is \"A1\"\n\n"
	// DEVIATION(D6): preview_only stays in the schema with its false default so
	// existing client calls keep validating; only the description changes, to say
	// the parameter has no effect. This is the one sanctioned departure from
	// SPEC 5.6's description-equality rule. SPEC 5.3.2 D6; registered in
	// conformance/deviations.json.
	descReadDataFromExcel     = "\nRead data from Excel worksheet with cell metadata including validation rules.\n\nArgs:\n    filepath: Path to Excel file\n    sheet_name: Name of worksheet\n    start_cell: Starting cell (default A1)\n    end_cell: Ending cell (optional, auto-expands if not provided)\n    preview_only: Accepted for compatibility and has no effect; the full requested range is always returned\n\nReturns:  \nJSON string containing structured cell data with validation metadata.\nEach cell includes: address, value, row, column, and validation info (if any).\n"
	descApplyFormula          = "\nApply Excel formula to cell.\nExcel formula will write to cell with verification.\n"
	descValidateFormulaSyntax = "Validate Excel formula syntax without applying it."
	descValidateExcelRange    = "Validate if a range exists and is properly formatted."
	descGetDataValidationInfo = "\nGet all data validation rules in a worksheet.\n\nThis tool helps identify which cell ranges have validation rules\nand what types of validation are applied.\n\nArgs:\n    filepath: Path to Excel file\n    sheet_name: Name of worksheet\n    \nReturns:\n    JSON string containing all validation rules in the worksheet\n"
	descFormatRange           = "Apply formatting to a range of cells."
	descMergeCells            = "Merge a range of cells."
	descUnmergeCells          = "Unmerge a range of cells."
	descGetMergedCells        = "Get merged cells in a worksheet."
	descCopyRange             = "Copy a range of cells to another location."
	descDeleteRange           = "Delete a range of cells and shift remaining cells."
	descInsertRows            = "Insert one or more rows starting at the specified row."
	descInsertColumns         = "Insert one or more columns starting at the specified column."
	descDeleteSheetRows       = "Delete one or more rows starting at the specified row."
	descDeleteSheetColumns    = "Delete one or more columns starting at the specified column."
	descCreateTable           = "Creates a native Excel table from a specified range of data."
	descCreateChart           = "Create chart in worksheet."
	descCreatePivotTable      = "Create pivot table in worksheet."
)
