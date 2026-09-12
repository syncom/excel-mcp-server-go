package mcpserver

import (
	"context"
	"fmt"
	"io"
	"log"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

// connect stands up the real server over the sandbox root dir and returns a
// connected client session.
func connect(t *testing.T, dir string) *mcp.ClientSession {
	t.Helper()

	srv := New(Config{ExcelFilesPath: dir})
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

func callText(t *testing.T, res *mcp.CallToolResult) string {
	t.Helper()
	if len(res.Content) == 0 {
		return ""
	}
	tc, ok := res.Content[0].(*mcp.TextContent)
	if !ok {
		return ""
	}
	return tc.Text
}

// Concurrent calls naming one workbook interleave their open-modify-save
// windows. excelize's SaveAs opens the target with O_TRUNC and streams the zip
// in place, so a save is not atomic: for most of its duration the file on disk
// is truncated or partial, and a reader that opens it then fails with "zip: not
// a valid zip file" or "unexpected EOF".
//
// The reproduction is a loop rather than a fixed batch of calls. A fixed batch
// mostly misses — the ops have to land inside each other's save window, which a
// handful of one-shot calls does only by luck — while a reader looping against
// a continuous writer hits it almost every time.
//
// Run with -race: the lock is also what keeps excelize's per-file state from
// being touched by two goroutines at once.
func TestConcurrentAccessOneWorkbook(t *testing.T) {
	log.SetOutput(io.Discard)
	t.Cleanup(func() { log.SetOutput(nil) })

	dir := t.TempDir()
	cs := connect(t, dir)
	ctx := context.Background()

	if _, err := cs.CallTool(ctx, &mcp.CallToolParams{
		Name:      "create_workbook",
		Arguments: map[string]any{"filepath": "book.xlsx"},
	}); err != nil {
		t.Fatalf("create_workbook: %v", err)
	}

	// Padded so each save is slow enough to have a window worth racing.
	pad := make([][]any, 0, 8000)
	for i := 0; i < 8000; i++ {
		pad = append(pad, []any{fmt.Sprintf("row-%d-padding-value", i), i})
	}
	if _, err := cs.CallTool(ctx, &mcp.CallToolParams{
		Name: "write_data_to_excel",
		Arguments: map[string]any{
			"filepath": "book.xlsx", "sheet_name": "Sheet1",
			"data": pad, "start_cell": "A1",
		},
	}); err != nil {
		t.Fatalf("padding write: %v", err)
	}

	deadline := time.Now().Add(3 * time.Second)
	var wg sync.WaitGroup
	var readFails, reads, writes int64
	var firstErr atomic.Value

	wg.Add(1)
	go func() { // writer: save the workbook continuously
		defer wg.Done()
		for i := 0; time.Now().Before(deadline); i++ {
			res, err := cs.CallTool(ctx, &mcp.CallToolParams{
				Name: "write_data_to_excel",
				Arguments: map[string]any{
					"filepath": "book.xlsx", "sheet_name": "Sheet1",
					"data":       [][]any{{fmt.Sprintf("v%d", i)}},
					"start_cell": "D1",
				},
			})
			if err == nil && !res.IsError {
				atomic.AddInt64(&writes, 1)
			}
		}
	}()

	wg.Add(1)
	go func() { // reader: open the same workbook continuously
		defer wg.Done()
		for time.Now().Before(deadline) {
			res, err := cs.CallTool(ctx, &mcp.CallToolParams{
				Name:      "get_workbook_metadata",
				Arguments: map[string]any{"filepath": "book.xlsx"},
			})
			atomic.AddInt64(&reads, 1)
			// get_workbook_metadata catches WorkbookError, so a corrupt-file
			// failure arrives as an ordinary text result reading "Error: ..."
			// with IsError unset (SPEC 5.2). Checking IsError alone would count
			// every torn read as a success.
			text := callText(t, res)
			switch {
			case err != nil:
				atomic.AddInt64(&readFails, 1)
				firstErr.CompareAndSwap(nil, err.Error())
			case res.IsError || strings.HasPrefix(text, "Error"):
				atomic.AddInt64(&readFails, 1)
				firstErr.CompareAndSwap(nil, text)
			}
		}
	}()
	wg.Wait()

	t.Logf("reads=%d writes=%d readFails=%d", reads, writes, atomic.LoadInt64(&readFails))
	if reads == 0 || writes == 0 {
		t.Fatalf("nothing raced: reads=%d writes=%d", reads, writes)
	}
	if n := atomic.LoadInt64(&readFails); n != 0 {
		t.Errorf("%d of %d reads saw a corrupt workbook (%d writes); first: %v",
			n, reads, writes, firstErr.Load())
	}
}

// Calls naming different workbooks must not serialize against each other: the
// lock is per path precisely so unrelated work still runs in parallel.
func TestDifferentWorkbooksDoNotShareALock(t *testing.T) {
	locks := newPathLocks()

	releaseA := locks.acquire("/tmp/a.xlsx")
	done := make(chan struct{})
	go func() {
		releaseB := locks.acquire("/tmp/b.xlsx")
		releaseB()
		close(done)
	}()
	<-done // would block forever if b waited on a
	releaseA()
}

// The registry must not accumulate one mutex per path a long-lived server has
// ever touched.
func TestPathLocksReleaseDropsEntry(t *testing.T) {
	locks := newPathLocks()

	for i := 0; i < 100; i++ {
		release := locks.acquire(fmt.Sprintf("/tmp/%d.xlsx", i))
		release()
	}

	locks.mu.Lock()
	defer locks.mu.Unlock()
	if len(locks.m) != 0 {
		t.Errorf("registry holds %d entries after every release, want 0", len(locks.m))
	}
}

// Two holders of one path must not overlap. The counter is deliberately
// unsynchronized: under -race an overlap is a reported data race, which is the
// assertion.
func TestPathLocksSerializeSamePath(t *testing.T) {
	locks := newPathLocks()
	key := filepath.Join(t.TempDir(), "book.xlsx")

	shared := 0
	var wg sync.WaitGroup
	for i := 0; i < 32; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			release := locks.acquire(key)
			defer release()
			shared++
		}()
	}
	wg.Wait()

	if shared != 32 {
		t.Errorf("counter = %d, want 32 — increments were lost to an overlap", shared)
	}
}
