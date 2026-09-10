// Package mcpserver is the port of excel_mcp/server.py: tool registration,
// annotations, schemas, dispatch, and the path sandbox.
package mcpserver

import (
	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// Identity of the server as declared in server.py's FastMCP(...) call
// (SPEC 5.8). These strings travel in the initialize response, so the
// conformance harness diffs them.
const (
	ServerName = "excel-mcp"

	// ServerVersion mirrors the version FastMCP reports for this package.
	// pyproject.toml declares 0.1.8.
	ServerVersion = "0.1.8"

	Instructions = "Excel MCP Server for manipulating Excel files"
)

// Config is the ported form of server.py's module-level EXCEL_FILES_PATH
// global. STANDARDS forbids package-level mutable state, so the sandbox root
// is carried in a value handed to the handlers instead.
type Config struct {
	// ExcelFilesPath is the sandbox root. The empty string is Python's None:
	// stdio mode, where filenames must already be absolute (SPEC 5.7).
	ExcelFilesPath string
}

// Server holds the configuration the tool handlers close over.
type Server struct {
	cfg Config
}

// New builds the MCP server. Tools are registered by [Server.register], which
// is filled in tier by tier.
func New(cfg Config) *mcp.Server {
	s := &Server{cfg: cfg}
	srv := mcp.NewServer(
		&mcp.Implementation{Name: ServerName, Version: ServerVersion},
		&mcp.ServerOptions{Instructions: Instructions},
	)
	s.register(srv)
	return srv
}

// register attaches every tool to srv, one call per gate tier of SPEC 4.
func (s *Server) register(srv *mcp.Server) {
	s.registerTierA(srv)
	s.registerTierB(srv)
	s.registerTierC(srv)
	s.registerTierD(srv)
	s.registerTierE(srv)
	s.registerTierF(srv)
}
