package mcpserver

import (
	"context"
	"errors"
	"fmt"
	"log"
	"runtime/debug"

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

// panicBarrier converts a panic inside any handler into a failed request.
//
// The MCP SDK has no recover() on the receiving path, so without this an
// unhandled panic in one tool call takes down the whole process — every other
// session and every in-flight request with it. Installed once as receiving
// middleware in [New] rather than wrapped around each of the 25 handlers.
//
// A panic is the Go analogue of the unhandled Python exception SPEC 5.2 routes
// through FastMCP's wrapper, so for tools/call the result is the same
// [raiseThrough] shape that an uncaught error produces. The panic value and its
// stack go to stderr, where they stay visible to operators without entering the
// protocol stream.
//
// This is a barrier for *panics* only. A Go out-of-memory is a fatal error, not
// a panic, and cannot be recovered — the bounds in excelops are what keep those
// from happening.
func panicBarrier(next mcp.MethodHandler) mcp.MethodHandler {
	return func(ctx context.Context, method string, req mcp.Request) (result mcp.Result, err error) {
		defer func() {
			r := recover()
			if r == nil {
				return
			}
			name := toolNameOf(req)
			log.Printf("panic in %s %s: %v\n%s", method, name, r, debug.Stack())
			if method == "tools/call" {
				result, err = raiseThrough(name, fmt.Sprintf("%v", r)), nil
				return
			}
			result, err = nil, fmt.Errorf("internal error in %s: %v", method, r)
		}()
		return next(ctx, method, req)
	}
}

// toolNameOf recovers the tool name for a tools/call request, for the error
// message and the log line. The SDK delivers tool arguments unparsed, so the
// name is read from the raw params.
func toolNameOf(req mcp.Request) string {
	if p, ok := req.GetParams().(*mcp.CallToolParamsRaw); ok && p != nil {
		return p.Name
	}
	return "unknown"
}
