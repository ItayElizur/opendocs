# WordTools.Commands.cs

## `ApplyCommands`

PP-12 Task 3 (the half PP-5 Task 4 Step 1 did not cover): each
result line is prefixed with the command's 0-based position in the
batch, and a summary header states how many succeeded/failed - with
partial batches now the norm (no rollback - Word COM offers no
batch transaction, and a hand-rolled undo would be less reliable
than this honest report; the user retains Word's own Ctrl+Z), the
model needs to know WHICH command in a batch of several identical
kinds failed, not just that "one of them" did.

## `ApplyCommands` - `set_bullet` case

Post-hoc addition (2026-08-27, user-reported): a model
sent kind:"set_bullet" and got a dead-end "unknown
command kind". It is an entirely reasonable guess -
the neighbouring commands are set_bold/set_italic/
set_heading, so a snake_case set_X for bullets reads
as the obvious name, while the real ones are
camelCase createParagraphBullets/deleteParagraphBullets.
Rather than expect the model to memorise an
inconsistency, accept the guess: set_bullet takes the
same target as the two it delegates to, plus a
value:true|false picking which.

## `ResolveTargetParagraphs` - forward enumeration, not positional indexing

Walks forward via the collection's own enumerator instead of
positional paragraphs[i + 1] indexing - Paragraphs is not a
real array in Word's COM object model, so indexing it by
position re-walks the document from the start on EVERY single
access, turning a full scan into roughly O(n^2) internally
(confirmed root cause of a real reported freeze in find_text/
get_headings, fixed there the same way). Every command that
funnels through this one function - updateTextStyle,
updateParagraphStyle, deleteBlocks, createParagraphBullets,
deleteParagraphBullets - inherits the fix.
