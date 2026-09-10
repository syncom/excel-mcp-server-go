package mcpserver

import (
	"errors"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// SPEC 5.2 splits failures into two different MCP outcomes, and which one a
// tool uses depends on the exception *class* its except clause names — not on
// the message, and not uniformly across tools. The kind sets below are that
// clause, transcribed as data so the routing is auditable against server.py in
// one place rather than reconstructed from if-chains at each call site
// (STANDARDS: Errors).

// textResult builds the successful result whose text server.py returns.
func textResult(text string) *mcp.CallToolResult {
	return &mcp.CallToolResult{
		Content: []mcp.Content{&mcp.TextContent{Text: text}},
	}
}

// raiseThrough builds the tool-error result an unhandled Python exception
// produces. FastMCP wraps the message, and that wrapper is part of the
// byte-exact contract, so it is reproduced here rather than left to the SDK.
func raiseThrough(toolName, msg string) *mcp.CallToolResult {
	return &mcp.CallToolResult{
		IsError: true,
		Content: []mcp.Content{
			&mcp.TextContent{Text: fmt.Sprintf("Error executing tool %s: %s", toolName, msg)},
		},
	}
}

// resultFor routes one tool outcome per SPEC 5.2.
//
// caught is the tool's except clause. An error whose kind appears there becomes
// a normal result reading "Error: {str(e)}"; anything else — including the
// ValueError from the path sandbox, which no tool catches — is re-raised and
// surfaces as a tool error.
func resultFor(toolName string, caught []error, text string, err error) *mcp.CallToolResult {
	if err == nil {
		return textResult(text)
	}
	for _, kind := range caught {
		if errors.Is(err, kind) {
			return textResult("Error: " + err.Error())
		}
	}
	return raiseThrough(toolName, err.Error())
}
