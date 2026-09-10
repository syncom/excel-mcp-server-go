// Command excel-mcp-server is the Go port of the excel-mcp-server Python CLI.
//
// It accepts the two transports the Python entry point exposes that are still
// supported: stdio and streamable-http. SSE is deprecated upstream and out of
// scope (SPEC 4), so `sse` is rejected like any other unknown subcommand.
package main

import (
	"context"
	"flag"
	"fmt"
	"log"
	"net"
	"net/http"
	"os"
	"strconv"

	"github.com/modelcontextprotocol/go-sdk/mcp"

	"github.com/syncom/excel-mcp-server-go/internal/mcpserver"
)

const usage = `Excel MCP Server

Usage:
  excel-mcp-server stdio
  excel-mcp-server streamable-http
`

func main() {
	// SPEC 5.8: in stdio mode nothing but JSON-RPC frames may reach stdout.
	// Python keeps stdout clean by logging to a file; stderr is equally safe
	// and is the Go convention.
	log.SetOutput(os.Stderr)
	log.SetFlags(log.LstdFlags)

	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

func run(args []string) error {
	if len(args) == 0 {
		fmt.Fprint(os.Stderr, usage)
		return fmt.Errorf("no subcommand given")
	}

	switch args[0] {
	case "stdio":
		fs := flag.NewFlagSet("stdio", flag.ContinueOnError)
		fs.SetOutput(os.Stderr)
		if err := fs.Parse(args[1:]); err != nil {
			return err
		}
		return runStdio(context.Background())
	case "streamable-http":
		fs := flag.NewFlagSet("streamable-http", flag.ContinueOnError)
		fs.SetOutput(os.Stderr)
		if err := fs.Parse(args[1:]); err != nil {
			return err
		}
		return runStreamableHTTP(context.Background())
	default:
		fmt.Fprint(os.Stderr, usage)
		return fmt.Errorf("unknown subcommand: %s", args[0])
	}
}

// runStdio mirrors server.run_stdio: no EXCEL_FILES_PATH, so get_excel_path
// requires absolute filenames.
func runStdio(ctx context.Context) error {
	log.Print("Starting Excel MCP server with stdio transport")
	srv := mcpserver.New(mcpserver.Config{})
	return srv.Run(ctx, &mcp.StdioTransport{})
}

// runStreamableHTTP mirrors server.run_streamable_http: EXCEL_FILES_PATH is
// read from the environment (default ./excel_files) and created if absent.
func runStreamableHTTP(ctx context.Context) error {
	filesPath := os.Getenv("EXCEL_FILES_PATH")
	if filesPath == "" {
		filesPath = "./excel_files"
	}
	if err := os.MkdirAll(filesPath, 0o755); err != nil {
		return err
	}

	host := os.Getenv("FASTMCP_HOST")
	if host == "" {
		host = "0.0.0.0"
	}
	port := os.Getenv("FASTMCP_PORT")
	if port == "" {
		port = "8017"
	}
	if _, err := strconv.Atoi(port); err != nil {
		// Python does int(os.environ.get("FASTMCP_PORT", "8017")) at import
		// time, so a non-numeric value is fatal there too.
		return fmt.Errorf("invalid FASTMCP_PORT: %s", port)
	}

	srv := mcpserver.New(mcpserver.Config{ExcelFilesPath: filesPath})
	handler := mcp.NewStreamableHTTPHandler(
		func(*http.Request) *mcp.Server { return srv },
		nil,
	)

	mux := http.NewServeMux()
	mux.Handle("/mcp", handler)

	addr := net.JoinHostPort(host, port)
	log.Printf("Starting Excel MCP server with streamable HTTP transport (files directory: %s) on %s/mcp", filesPath, addr)

	httpSrv := &http.Server{Addr: addr, Handler: mux}
	return httpSrv.ListenAndServe()
}
