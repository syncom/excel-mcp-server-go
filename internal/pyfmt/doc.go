// Package pyfmt reproduces the two CPython serializations the tool results are
// specified in terms of: the built-in repr() used by get_workbook_metadata and
// get_merged_cells (SPEC 5.4), and json.dumps(..., indent=2, default=str) used
// by read_data_from_excel and get_data_validation_info (SPEC 5.5).
//
// Both are filled in by the tiers that first need them (T4 and T5).
package pyfmt
