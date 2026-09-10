// Package excelerr mirrors excel_mcp/exceptions.py.
//
// server.py routes a failure to one of two different MCP outcomes based on the
// exception *class*, not its message (SPEC 5.2), so the port needs the same
// taxonomy as a thing it can branch on. Each Python exception class becomes a
// sentinel error; a concrete [Error] carries the byte-exact message and unwraps
// to its sentinel, so handlers use errors.Is on the kind and the wire sees only
// the message.
package excelerr

import "errors"

// One sentinel per class in exceptions.py. ExcelMCPError is the base class
// there and the other eight derive from it, but no tool's except clause names
// the base, so the sentinels are flat: routing compares against the concrete
// kind exactly as server.py's except clauses do.
var (
	ErrExcelMCP    = errors.New("excel mcp error")
	ErrWorkbook    = errors.New("workbook error")
	ErrSheet       = errors.New("sheet error")
	ErrData        = errors.New("data error")
	ErrValidation  = errors.New("validation error")
	ErrFormatting  = errors.New("formatting error")
	ErrCalculation = errors.New("calculation error")
	ErrPivot       = errors.New("pivot error")
	ErrChart       = errors.New("chart error")
)

// Error is a Python exception instance: a kind plus the exact str(e) that
// server.py interpolates into "Error: {str(e)}".
type Error struct {
	Kind error
	Msg  string
}

// Error returns the message byte-for-byte as Python's str(e) would.
func (e *Error) Error() string { return e.Msg }

// Unwrap exposes the kind so callers can use errors.Is(err, excelerr.ErrSheet).
func (e *Error) Unwrap() error { return e.Kind }

// New builds an [Error] of the given kind.
func New(kind error, msg string) *Error { return &Error{Kind: kind, Msg: msg} }
