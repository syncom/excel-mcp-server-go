# Deviations from the Python server

This Go port reproduces `excel-mcp-server` byte for byte, including several of its
oddities, with **eight** deliberate exceptions. This document is for someone deciding
whether to swap the servers: what the Python server does, what this one does instead, why,
and who would notice.

The set is closed: these eight are the only intended differences. Each was found by
differential testing against the Python server, and each is pinned to an expectation on
both sides, so a sanctioned difference cannot mask an unrelated regression.

**Two of these change results you may be relying on** — D2 and D4. The rest either fix
something that never worked (D1, D3, D5, D7), refuse a request that would otherwise kill
the server (D8), or change only documentation (D6).

---

## D1 — `create_pivot_table` accepts `agg_func: "mean"`

**Python:** the parameter's declared default is `"mean"`, but the implementation only
accepts `sum`, `average`, `count`, `min`, `max`. So *every* call that omits `agg_func`
fails with:

```
Error: Invalid aggregation function. Must be one of: sum, average, count, min, max
```

The tool is unusable at its own default.

**Go:** `mean` is accepted as a case-insensitive alias for `average`. The schema default
stays `"mean"`, so tool metadata is unchanged, and the pivot's header cell still reads
`Revenue (mean)` — the caller asked for "mean", so that is what it is labelled. Every other
invalid value still produces the error above, with the same five-name list.

**Who notices:** anyone omitting `agg_func`, or passing `mean`, now gets a pivot table
instead of an error. If you built error handling around that failure, it stops firing.

---

## D2 — `validate_excel_range` actually validates the end of the range

**Python:** `server.py` joins the start and end cells into `"A1:Z999"` and passes the
result as the *start cell*. The parser uses `re.match`, which stops at the first match, so
the string is silently truncated to `A1` and only the start cell is bounds-checked. On a
sheet with four rows of data:

```
validate_excel_range(start_cell="A1", end_cell="Z999")
-> Range 'A1:Z999' is valid. Sheet contains data in range 'A1:C4'
```

A validator that calls `A1:Z999` valid on a four-row sheet is worse than no validator.

**Go:** both ends are parsed and the whole range is bounds-checked. The success message is
unchanged. An out-of-range end now produces the error strings that already existed, unused,
in `validation.py`:

```
Error: End row 999 out of bounds (1-4)
Error: End column AA out of bounds (A-C)
Error: End row cannot be before start row
Error: End column cannot be before start column
```

**Who notices:** callers who passed deliberately oversized ranges — "validate A1:Z999, I'll
write into whatever fits" — now get an error where they used to get approval. This is the
deviation most likely to surface in existing automation.

---

## D3 — `create_pivot_table` honours `columns`

**Python:** `columns` is accepted, advertised in the tool schema, echoed back in the
details payload, and **ignored**. The internal filter that would use it is always called
with an empty argument. A caller asking to break revenue out by product gets a row-only
group-by that looks like a perfectly good answer.

**Go:** `columns` produces a real cross-tab. Column combinations are built exactly the way
row combinations are — unique values per field, sorted, cartesian product in field order —
and one column is emitted per (column combination x value field), with headers like:

```
Gadget - Revenue (sum)   Widget - Revenue (sum)
```

When `columns` is empty the output is byte-identical to Python's.

**Who notices:** anyone who passed `columns` and built on the result. Their `_pivot` sheet
changes shape — but the previous shape was answering a different question than the one
they asked.

---

## D4 — `delete_range` no longer destroys data outside the range

**This is the most important one.**

**Python:** `delete_range_operation` blanks the cells you named and *then also* deletes the
entire rows (or columns) spanning them. `delete_range("A1:B2", shift_direction="up")` blanks
A1:B2 and then deletes rows 1 and 2 across **every column in the sheet**. Data in C1:Z2 —
which the caller never mentioned — is gone, and everything below moves up.

**Go:** Excel's actual "delete cells, shift up/left". The range's cells are removed and the
cells below (for `up`) or to the right (for `left`) shift into the gap, **confined to the
range's own column span (`up`) or row span (`left`)**. Cells outside that span are not
touched. Bounds and shift-direction validation are unchanged, including
`Invalid shift direction: {d}. Must be 'up' or 'left'`, and so is the success message:

```
Range A1:B2 deleted successfully
```

**Who notices:** anyone who was relying — deliberately or, far more likely, unknowingly —
on whole rows disappearing. They keep data they used to lose. If you actually want to
delete whole rows, `delete_sheet_rows` does that and always did.

---

## D5 — colour handling

**Python:** prefixes `FF` to any colour string that does not already start with `FF`. Two
things fall out of that:

- An 8-digit ARGB that does not begin with `FF` — `80FF0000`, a 50%-opacity red — becomes
  the 10-character `FF80FF0000`, which openpyxl rejects. A legitimate colour is unusable.
- A 6-digit colour that happens to begin with `FF` (`FFC7CE`, a common light red) skips the
  prefix and is stored with a different alpha than every other colour.

Both a malformed colour and a valid 8-digit one produce the same opaque message:

```
Error: Invalid font color: Colors must be aRGB hex values
```

**Go:** colours are normalized once, the same way for `font_color`, `bg_color`,
`border_color` and conditional-format fill colours: strip a leading `#`, uppercase, require
exactly 6 or 8 hex digits, left-pad 6 digits with `FF`. Anything else is rejected with a
message that names the value:

```
Error: Invalid color: ZZZZZZ
```

**Who notices:** `80FF0000` and `#FF0000` now work where they used to error. Stored colour
strings differ in their alpha byte for some inputs; the rendered colour is the same.

**One limitation:** the underlying Go library forces full opacity on every colour it
writes, so an 8-digit colour with a non-`FF` alpha is *accepted* — which is the point, since
Python rejects it outright — but stored opaque. If you need partial transparency, neither
server delivers it today.

---

## D7 — `create_chart` puts the chart where you asked

**Python:** `chart.py` builds an anchor from `target_cell`, attaches it to a drawing, and
appends that drawing to the worksheet. But openpyxl's writer serializes the worksheet's
*charts*, not its drawings, and uses each chart's own anchor — which is the library
default, `E15`. The drawing, and with it your anchor, is thrown away. So **every chart the
Python server writes lands at E15**, stacked on top of whatever was written before it.
`target_cell` is parsed and validated, so it can reject a call, but it can never place a
chart.

The validation is broken in its own right: it reads only the *first character* of
`target_cell` as the column and `int()` of everything after it as the row. So any
multi-letter column — `AA10`, and every column past `Z` — is rejected outright:

```
create_chart(target_cell="AA10")
-> Error: Invalid target cell: invalid literal for int() with base 10: 'A10'
```

**Go:** `target_cell` is parsed as a whole A1 reference and the chart is anchored there.
Multi-letter columns work. The fix is bounded so that nothing Python rejects starts
succeeding by accident:

- A reference with no letter or no digit is rejected exactly as before, with the same
  `Failed to create chart drawing: Invalid target cell format: ...` text.
- A reference that is not a well-formed A1 name falls back to Python's own first-character
  computation, so the error text is byte-identical for every input Python also rejects.
- A reference resolving below row or column 1 keeps the library default, so the call still
  succeeds exactly as Python's does.
- The success message is unchanged.

**Who notices:** charts stop piling up at E15 and appear where they were requested, and
`target_cell` values with two-letter columns now succeed instead of erroring. If you had
worked around the old behaviour — hard-coding E15 as the known landing spot, or moving
charts afterwards — that workaround is now wrong.

---

## D6 — `preview_only` is documented rather than removed

**Python:** `read_data_from_excel` accepts a `preview_only` parameter and ignores it
completely.

**Go:** the parameter stays in the schema, with its type and its `false` default, so
existing client calls keep validating. Only its **description** changes, to say that it is
accepted for compatibility and has no effect. This is the one place where a tool
description differs from the Python server's.

**Who notices:** nobody at call time — the behaviour is identical. A person reading the tool
list sees accurate prose instead of a parameter that promises something it never did.

---

## D8 — oversized requests are refused instead of killing the server

**This one is about availability, not results.** It is the only deviation that makes a
request fail that Python would attempt.

**Python:** three code paths are unbounded, and all three are reachable from a single
ordinary tool call.

- `read_data_from_excel` iterates `range(start_row, end_row + 1)` over an `end_cell` taken
  straight from the caller, with no check against the sheet's extent. `openpyxl`'s
  `ws.cell()` creates each cell it is asked for, so the loop happily walks the empty part
  of the grid.
- `format_range` has the same unclamped loop, and each iteration mutates workbook state.
- `create_pivot_table` builds the full cartesian product of the distinct values of every
  grouped field. Three fields over high-cardinality columns is N³ — and dates, IDs, names
  and emails are exactly the columns people group by.

CPython survives the first two: a `MemoryError` is an ordinary exception, so
`except Exception` turns it into a `DataError` and the caller gets an error result back.

**Go:** the same three paths are bounded, and exceeding a bound raises the calling tool's
own error kind — so the refusal routes to the same MCP outcome as any other failure of
that tool, and says what the limit was.

| Path | Bound | Raised as |
|---|---|---|
| `read_data_from_excel` | 200,000 cells per range | `DataError` |
| `format_range` | 200,000 cells per range | `FormattingError` |
| `create_pivot_table` | 50,000 grouped combinations | `PivotError` |

**Why this cannot be left to match Python.** A Go out-of-memory is `fatal error: out of
memory` — not a panic, and not recoverable. There is no `recover()` that catches it. It
takes down the whole process, every session on it and every in-flight request, so the
failure is not confined to the caller who caused it. Reproduced against the built binary:
`end_cell: "XFD1048576"` on a workbook with two data rows kills the server, and a pivot
over three all-distinct columns of a 120-row sheet allocates 2 GB doing it. Neither needs
a malformed argument or a crafted file.

So for reads and formats the bound *restores* the observable Python outcome — an error
result instead of a dead server. For pivots it is a genuine new refusal: Python has the
same blowup and the port inherited it, and neither server produces anything useful at that
size.

**Where the numbers come from.** Measured cost per visited cell is ~800 B on the read path
and ~2 KB on the format path, so 200,000 cells is a transient of roughly 150 MB and 360 MB.
Both limits sit far above anything a caller can consume — 200,000 cells is already tens of
megabytes of response text, and a pivot table with 50,000 grouped rows is not an answer to
any question — and far below the point where the process is at risk.

**Who notices:** nobody making a request whose result they intended to read. A caller who
passed a deliberately oversized range — "read A1:XFD1048576, I'll take whatever is there" —
now gets an error naming the limit instead of a dropped connection.

---

## Deliberately *not* changed

Four Python behaviours look like bugs and were left alone, because reproducing them is
correct or because fixing them is not this port's call:

- **Auto-detected read ranges.** `read_data_from_excel` with no `end_cell` and a
  `start_cell` of `A1` re-derives the start from the sheet's used range, so it may return a
  range that does not begin at A1. The response names the range it actually read, so the
  behaviour is self-describing, and auto-detecting is the useful reading of "no end cell
  given".
- **Protective guards.** `copy_worksheet` refuses an existing target and `delete_worksheet`
  refuses the last remaining sheet. Both are correct.
- **Numeric pivot fields aggregate to zero.** Group-by values are stringified before
  filtering but compared against raw cell values, so grouping by a numeric column yields
  zeros. Matched, not fixed.
- **`agg_func` is case-sensitive for aggregation.** `create_pivot_table` validates
  `agg_func` case-insensitively but aggregates case-sensitively, so `"COUNT"` passes
  validation and then silently *sums* instead of counting. Matched, not fixed. Use the
  lower-case spellings.

## Known gaps this port cannot close

Four differences remain that are not deliberate design choices — they are places the Go
libraries cannot express what openpyxl does, or where the Python message is not
reproducible at all. They are listed here so nobody has to rediscover them.

| Area | Difference |
|---|---|
| `format_range` conditional `icon_set` | The `type` and `values` parameters are accepted and ignored; thresholds are always the 3-band preset (percent 0/33/67). The Go library exposes no way to set custom icon thresholds. Python honours them. |
| `format_range` invalid `border_style`, `alignment`, or `cell_is` `operator` | Python's message lists the whole set of legal values, and Python renders a `set`, whose order varies between processes. That string cannot be matched byte-for-byte, so this port reports the offending value instead. |
| `format_range` unknown `conditional_format.params` keys | Python splats `params` into an openpyxl rule constructor, so an unrecognised key raises `TypeError: ... unexpected keyword argument`. This port ignores unknown keys. |
| Column references beyond `XFD` (16384) | openpyxl accepts `A`–`ZZZ` (up to 18278); the Go library caps at `XFD`. References in `XFE`–`ZZZ` work in Python and fail here. Four-letter and longer references produce the same error on both. |
| `create_table` with `table_style: ""` | Python stores an empty style name; the Go library omits the style element entirely. The tool response is identical. |
