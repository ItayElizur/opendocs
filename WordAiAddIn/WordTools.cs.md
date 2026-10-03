# WordTools.cs

## `Execute` - undo-record grouping (one tool call = one undo step)

Word's native undo stack registers one entry PER COM write,
not per tool call - confirmed live, 2026-10-02: add_table's
cell-by-cell `Range.Text =` loop (WordTools.Tables.cs) left
undo_last_action peeling back one cell at a time instead of
reverting the whole table in one step (the mirror image of
PowerPoint's issue, which coalesces too much instead of too
little - see PowerPointTools.cs's Execute() for that fix).
UndoRecord.StartCustomRecord/EndCustomRecord (confirmed via
reflection against the referenced PIA) groups every COM
write between the two calls into one user-visible undo
entry, so one tool call always maps to exactly one undo
step. Not used for always-allowed (read-only) tools or
undo/redo themselves - wrapping Document.Undo()/Redo() in a
custom record would be meaningless. Must run in try/finally:
a tool throwing mid-mutation without EndCustomRecord would
leave Word recording forever, silently absorbing every
later edit (including the user's own) into one entry.

## `Execute` - finally does not let EndCustomRecord mask success

Caught separately (PR review, 2026-10-02): a successful
ToolResult already computed by the switch above is
still in flight when a finally block runs - if
EndCustomRecord() itself threw uncaught here, C#'s
finally-after-return semantics would discard that
already-successful result and propagate this exception
to the outer catch instead, reporting a real mutation
as a generic failure. Logged, not rethrown, so a
cosmetic undo-grouping failure can never mask a
mutation that actually succeeded.

## `ActiveDoc`

Known limitation (PP-1 Task 5 Step 5): resolves whichever document is
ACTIVE right now, not necessarily the one whose pane initiated this
tool call. A tool call is always initiated by a user in the
currently-focused window, so this is normally correct - but a
long-running run whose user switches windows mid-run would write
into the newly-active document instead of the one the run started
against. Fixing this needs per-document COM target resolution
across every executor method - out of scope here; left as a known
issue for a follow-up item.

## Trailing note (end of file, chart workbook retry rationale)

PP-9: ported from PowerPointTools.AddChartPpt's data-writing block -
the embedded chart workbook MUST be closed and released in a
finally, or a leaked hidden Excel process stays alive for the rest
of the Word session. seriesArray items are {name?, values}.
Post-hoc fix (2026-08-24, user-reported): the embedded workbook's
OLE server occasionally still throws "The remote procedure call
failed" (HRESULT 0x800706BE) even after the Clear()+batched-write
fix above - a known, documented transient failure mode for rapid
COM calls against Office's embedded chart-data Excel object, not
something a single call can eliminate. Retrying after a short
delay is the standard mitigation; only the specific known
transient RPC HRESULTs are retried, so a genuine logic error
(bad range, etc.) still fails immediately rather than being
masked for 3 attempts.
