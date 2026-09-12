package mcpserver

import (
	"context"
	"io"
	"log"
	"strings"
	"testing"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// connectPanicTool stands up a server carrying the panic barrier plus one tool
// that panics, and returns a connected client session.
func connectPanicTool(t *testing.T, panicWith any) *mcp.ClientSession {
	t.Helper()

	srv := mcp.NewServer(&mcp.Implementation{Name: ServerName, Version: ServerVersion}, nil)
	srv.AddReceivingMiddleware(panicBarrier)
	mcp.AddTool(srv, &mcp.Tool{
		Name:        "boom",
		InputSchema: object("boomArguments", nil, map[string]*prop{}),
	}, func(context.Context, *mcp.CallToolRequest, struct{}) (*mcp.CallToolResult, any, error) {
		panic(panicWith)
	})

	ct, st := mcp.NewInMemoryTransports()
	ss, err := srv.Connect(context.Background(), st, nil)
	if err != nil {
		t.Fatalf("server connect: %v", err)
	}
	t.Cleanup(func() { ss.Close() })

	cs, err := mcp.NewClient(&mcp.Implementation{Name: "test", Version: "0"}, nil).
		Connect(context.Background(), ct, nil)
	if err != nil {
		t.Fatalf("client connect: %v", err)
	}
	t.Cleanup(func() { cs.Close() })
	return cs
}

// A panicking handler must return a tool error rather than killing the process.
// Without the barrier this test crashes the whole test binary.
func TestPanicBarrierReturnsToolError(t *testing.T) {
	log.SetOutput(io.Discard)
	t.Cleanup(func() { log.SetOutput(nil) })

	cs := connectPanicTool(t, "kaboom")
	res, err := cs.CallTool(context.Background(), &mcp.CallToolParams{Name: "boom"})
	if err != nil {
		t.Fatalf("CallTool returned a transport error, want a result: %v", err)
	}
	if !res.IsError {
		t.Fatalf("IsError = false, want true")
	}
	text := res.Content[0].(*mcp.TextContent).Text
	// Same shape raiseThrough produces for an uncaught error (SPEC 5.2).
	if want := "Error executing tool boom: kaboom"; text != want {
		t.Errorf("text = %q, want %q", text, want)
	}
}

// The session must stay usable after a panic — the barrier contains the failure
// to the one call.
func TestPanicBarrierKeepsSessionAlive(t *testing.T) {
	log.SetOutput(io.Discard)
	t.Cleanup(func() { log.SetOutput(nil) })

	cs := connectPanicTool(t, "kaboom")
	for i := 0; i < 3; i++ {
		if _, err := cs.CallTool(context.Background(), &mcp.CallToolParams{Name: "boom"}); err != nil {
			t.Fatalf("call %d: session died: %v", i, err)
		}
	}
	tools, err := cs.ListTools(context.Background(), nil)
	if err != nil {
		t.Fatalf("ListTools after panics: %v", err)
	}
	if len(tools.Tools) != 1 {
		t.Errorf("len(tools) = %d, want 1", len(tools.Tools))
	}
}

// A nil-map or nil-pointer dereference is the realistic case, and its panic
// value is an error rather than a string.
func TestPanicBarrierHandlesRuntimeError(t *testing.T) {
	log.SetOutput(io.Discard)
	t.Cleanup(func() { log.SetOutput(nil) })

	cs := connectPanicTool(t, nil) // panic(nil) becomes a runtime.PanicNilError
	res, err := cs.CallTool(context.Background(), &mcp.CallToolParams{Name: "boom"})
	if err != nil {
		t.Fatalf("CallTool returned a transport error, want a result: %v", err)
	}
	if !res.IsError {
		t.Fatalf("IsError = false, want true")
	}
	if text := res.Content[0].(*mcp.TextContent).Text; !strings.HasPrefix(text, "Error executing tool boom: ") {
		t.Errorf("text = %q, want the raiseThrough prefix", text)
	}
}
