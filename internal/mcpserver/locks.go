package mcpserver

import (
	"context"
	"encoding/json"
	"sync"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// Every excelops operation is open -> modify -> Save, with no coordination
// between calls. Two requests naming the same workbook interleave their saves:
// concurrent callers see "zip: not a valid zip file" and EOF reading a file
// that was just written, and a torn write can outlive the storm. This matters
// most in streamable-http mode, where one server serves concurrent clients.
//
// The lock is per resolved path rather than global so that work on unrelated
// workbooks still runs in parallel — the contention that needs serializing is
// between calls naming the same file, and nothing else.

// pathLocks hands out one mutex per workbook path.
//
// Entries are reference-counted and dropped when the last holder releases, so
// a long-lived server that touches many paths does not accumulate a mutex per
// path it has ever seen.
type pathLocks struct {
	mu sync.Mutex
	m  map[string]*pathLock
}

type pathLock struct {
	mu   sync.Mutex
	refs int
}

func newPathLocks() *pathLocks {
	return &pathLocks{m: make(map[string]*pathLock)}
}

// acquire blocks until key is held and returns the release function.
func (p *pathLocks) acquire(key string) func() {
	p.mu.Lock()
	l, ok := p.m[key]
	if !ok {
		l = &pathLock{}
		p.m[key] = l
	}
	// Counted while the registry lock is held, so the entry cannot be dropped
	// between here and the Lock below.
	l.refs++
	p.mu.Unlock()

	l.mu.Lock()

	return func() {
		l.mu.Unlock()

		p.mu.Lock()
		l.refs--
		if l.refs == 0 {
			delete(p.m, key)
		}
		p.mu.Unlock()
	}
}

// serializeByPath is receiving middleware that holds the workbook's lock for
// the duration of a tools/call.
//
// The path is taken from the raw arguments and resolved through the same
// [Server.ExcelPath] the handler uses, so the key is the file the handler will
// actually touch — two spellings of one path share a lock. Every tool takes at
// most one filepath, so a request never holds more than one of these and the
// locks cannot deadlock against each other.
//
// Anything without a usable filepath — a different method, a malformed
// argument object, a path the sandbox rejects — passes straight through. Those
// requests either touch no workbook or fail in the handler, which reports the
// error with the message SPEC 5.2 requires; swallowing it here would change
// what the caller sees.
func (s *Server) serializeByPath(next mcp.MethodHandler) mcp.MethodHandler {
	return func(ctx context.Context, method string, req mcp.Request) (mcp.Result, error) {
		if method != "tools/call" {
			return next(ctx, method, req)
		}
		params, ok := req.GetParams().(*mcp.CallToolParamsRaw)
		if !ok || params == nil {
			return next(ctx, method, req)
		}

		var args struct {
			Filepath string `json:"filepath"`
		}
		if err := json.Unmarshal(params.Arguments, &args); err != nil || args.Filepath == "" {
			return next(ctx, method, req)
		}
		full, err := s.ExcelPath(args.Filepath)
		if err != nil {
			return next(ctx, method, req)
		}

		release := s.locks.acquire(full)
		defer release()
		return next(ctx, method, req)
	}
}
