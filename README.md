# excel-mcp-server-go

A Go port of [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server), built to be
a drop-in replacement for the Python server over MCP stdio and streamable HTTP.

It exposes the same 25 tools, with responses that are byte-for-byte identical to the
Python server's — apart from seven deliberate, enumerated differences. Those are documented
in [DEVIATIONS.md](DEVIATIONS.md), and you should read that file before swapping the
servers, because two of them change behaviour callers may be relying on.

## Requirements

Go 1.25 or later. Building pulls two direct module dependencies
([excelize](https://github.com/xuri/excelize) and the
[MCP Go SDK](https://github.com/modelcontextprotocol/go-sdk)); the result is a single
self-contained binary with no runtime dependencies. Python is not needed to build or run
the server.

## Build

```bash
git clone https://github.com/syncom/excel-mcp-server-go.git
cd excel-mcp-server-go
go build -o bin/excel-mcp-server ./cmd/excel-mcp-server
```

Cross-compiling works the usual way:

```bash
GOOS=darwin  GOARCH=arm64 go build -o bin/excel-mcp-server-darwin-arm64  ./cmd/excel-mcp-server
GOOS=windows GOARCH=amd64 go build -o bin/excel-mcp-server-windows-amd64.exe ./cmd/excel-mcp-server
GOOS=linux   GOARCH=amd64 go build -o bin/excel-mcp-server-linux-amd64   ./cmd/excel-mcp-server
```

## Install

From a clone:

```bash
cd excel-mcp-server-go
go install ./cmd/excel-mcp-server
```

This puts `excel-mcp-server` in `$(go env GOPATH)/bin`.

## Running

The server has two subcommands. Anything else exits non-zero.

```bash
excel-mcp-server stdio            # for a local MCP client
excel-mcp-server streamable-http  # for a remote client, on /mcp
```

SSE is not supported. It is deprecated upstream, and there is deliberately no `sse`
subcommand — invoking one is an error rather than a silent fallback.

### Environment

| Variable | Default | Applies to | Meaning |
|---|---|---|---|
| `EXCEL_FILES_PATH` | `./excel_files` | streamable-http | Sandbox root. Created if absent. |
| `FASTMCP_HOST` | `0.0.0.0` | streamable-http | Listen address. |
| `FASTMCP_PORT` | `8017` | streamable-http | Listen port. |

The variable names are the Python server's, kept so an existing deployment's configuration
keeps working unchanged.

### How file paths are resolved

This differs by transport, and the difference is a security boundary rather than a
convenience:

- **stdio** — `EXCEL_FILES_PATH` is not consulted. Every `filepath` argument must be
  **absolute**; it is cleaned lexically and used as given. The client is trusted, because
  it already runs the server as a subprocess with the same privileges.
- **streamable-http** — every `filepath` must be **relative**. It is resolved against
  `EXCEL_FILES_PATH` with symlinks followed on both sides, and rejected if the result
  lands outside the sandbox. A symlink inside the sandbox that points out of it is
  rejected, which a purely lexical check would miss.

## Logging

Logs go to **stderr**, never stdout. In stdio mode a single stray byte on stdout corrupts
the JSON-RPC framing for the whole session, so stdout carries nothing but JSON-RPC frames.

The Python server writes to `excel-mcp.log` instead; if you were tailing that file, tail
stderr now.

## MCP client configuration

### Claude Desktop / any stdio client

```json
{
  "mcpServers": {
    "excel": {
      "command": "/absolute/path/to/excel-mcp-server",
      "args": ["stdio"]
    }
  }
}
```

Remember that in stdio mode tools take absolute paths, so ask for
`/Users/you/Documents/book.xlsx`, not `book.xlsx`.

### Claude CLI

```bash
claude mcp add excel -- /absolute/path/to/excel-mcp-server stdio
```

### Streamable HTTP

```bash
EXCEL_FILES_PATH=/srv/excel FASTMCP_PORT=8017 excel-mcp-server streamable-http
```

```json
{
  "mcpServers": {
    "excel": {
      "type": "http",
      "url": "http://localhost:8017/mcp"
    }
  }
}
```

Here tools take paths relative to `EXCEL_FILES_PATH` — `reports/q3.xlsx`.

## Tools

The same 25 as the Python server, in six groups:

| Group | Tools |
|---|---|
| Workbook & sheet lifecycle | `create_workbook`, `create_worksheet`, `get_workbook_metadata`, `copy_worksheet`, `delete_worksheet`, `rename_worksheet` |
| Data I/O | `write_data_to_excel`, `read_data_from_excel` |
| Formulas & validation | `apply_formula`, `validate_formula_syntax`, `validate_excel_range`, `get_data_validation_info` |
| Formatting & merges | `format_range`, `merge_cells`, `unmerge_cells`, `get_merged_cells` |
| Range & structure | `copy_range`, `delete_range`, `insert_rows`, `insert_columns`, `delete_sheet_rows`, `delete_sheet_columns` |
| Tables, charts, pivots | `create_table`, `create_chart`, `create_pivot_table` |

Tool descriptions, parameter names, types, defaults and annotations are copied from the
Python server, so a client that works against one works against the other without changes.

## Layout

```
cmd/excel-mcp-server/   CLI entry point: subcommands, transports, environment
internal/mcpserver/     Tool schemas, descriptions, dispatch, path resolution
internal/excelops/      The Excel operations themselves
internal/pyfmt/         Python-compatible number, repr and JSON formatting
internal/excelerr/      Error kinds, mapped to the Python server's channels
```

`internal/pyfmt` exists because parity is byte-level: Python's `repr`, `%g` and
`json.dumps` spacing all differ from Go's defaults, and the responses have to match.

## Development

```bash
go build ./...          # compile
go test ./... -race     # unit tests
go vet ./...            # vet
gofmt -l .              # list files gofmt would change
```

Parity with the Python server was established with a differential harness that runs both
servers against a shared transcript and compares responses and resulting workbooks
structurally. That harness needs a Python checkout of the upstream server and is not part
of this repository; `DEVIATIONS.md` records everything it found that differs.
