# Agent guidance for officeoffice

This file is for an AI coding agent working in this repo. It only covers things that
are specific to *this* codebase's conventions and pitfalls — not general engineering
advice. Read `README.md` first for what the repo is and how to build/test it, and
`docs/ai-tool-surface.md` for the full current tool catalog.

## The `.csproj` files do not glob — a new `.cs` file is silently dropped otherwise

`WordAiAddIn.csproj`, `ExcelAiAddIn.csproj`, `PowerPointAiAddIn.csproj`,
`OutlookAiAddIn.csproj`, and `OfficeAi.Shared.csproj` are all **classic MSBuild format
with explicit `<Compile Include="...">` items** — none of them glob `*.cs`. If you add
a new `.cs` file to one of these projects and don't add a matching `<Compile Include>`
line to its `.csproj`, the file **silently does not compile**. There is no error at the
new file itself — MSBuild just never sees it, and the first symptom is usually a
compile error at some unrelated call site that expected the new type/method to exist.
This bit the project once already: see `docs/ai-tool-surface-changelog.md`'s
"Phases 1+3" entry, which split the three `*Tools.cs` files into ~10 partial-class
files each and had to add every new file to its `.csproj` by hand.

`tools/split-partial.py` exists for exactly this — it did that original split and can
do further ones; it refuses to run unless every member is assigned to exactly one
destination file. If you're adding a new `.cs` file yourself (not via that script),
add the `<Compile Include>` line in the same commit, and verify with an actual MSBuild
run, not just "the IDE didn't complain."

## The four-tier `EditingMode` gate has two independent copies that must agree

Every add-in enforces `OfficeAi.Shared.EditingMode` (`ReadOnly < CommentOnly <
TrackChanges < FullAutonomy`, ordinal comparison) **twice**: once client-side (which
tools `entry.ts` advertises to the model for the current mode) and once server-side (an
independent check inside the C# tool executor, so a model can never reach a mutating
handler just because the client forgot to hide the tool). These two lists are
maintained by hand in two different languages and **will silently drift** if you add a
tool to one side and forget the other — the failure mode is either "the model sees a
tool it can't actually call" or "a mutating tool is reachable in a mode meant to block
it," and neither throws a compile error.

The clearest real example of the pattern (and the one to copy) is Outlook's tier 2
("Draft only"): `OutlookAiAddIn/OutlookTools.cs`'s `DraftTierTools` hash set and
`OutlookAiAddIn/web-src/entry.ts`'s `commentOnlyExtraTools` array must list the exact
same tool names. Both sides already carry an explicit comment saying so
(`OutlookTools.cs`: "Must stay in sync with entry.ts's commentOnlyExtraTools"; `entry.ts`:
"Must stay in sync with OutlookTools.cs's DraftTierTools") — keep that comment pattern
if you touch either list, and update both in the same commit. The same applies to
`SendTierTools`/`ApprovalTierTools` (C#) against `trackChangesExtraTools` (TS), and to
each of Word/Excel/PowerPoint's own `readOnlyTools`/`commentOnlyExtraTools` arrays
against their C# `Tools.Execute()` gating.

## There is no live Office to test against in most automated contexts

Almost none of the COM-calling code (`*Tools*.cs`) has automated test coverage, and
most sessions working on this repo have **no interactive Word/Excel/PowerPoint/Outlook
to run against**. `dotnet test OfficeAi.Shared.Tests` only covers pure logic
(`TextUtil`, `ColorUtil`, `ChartTypes`, `GeometryUtil`, `OutlookDasl`, `MeetingSlots`,
etc.) that has been deliberately extracted to be COM-free and unit-testable — new pure
logic belongs there by default; new logic that has to touch a COM type doesn't have
that option here.

This repo's convention is to **say so explicitly** rather than claim untested COM code
works, and you should follow it: `docs/ai-tool-surface.md` and
`docs/superpowers/verification/*.md` are full of precedent for the expected tone —
e.g. "**NOT VERIFIED AGAINST LIVE POWERPOINT** — compiles clean in Debug and Release,
but the SmartArt COM paths … were never exercised against a running instance" and
"Manual verification matrix (none of this has been run — no interactive Word session
reachable from this environment)". When you implement or change a COM-calling tool
without being able to run it live, say what you verified (build clean, reflection
against the referenced PIA to confirm a method signature/enum value exists, a
cross-check against an equivalent working call elsewhere in the repo) and what you
didn't (the actual rendered/mutated result in a real Office document) — don't imply
more confidence than a build-clean compile actually gives you.

## Where to update docs after a change

`docs/ai-tool-surface.md` is the current-state reference for every tool the AI can
call — its per-app tables are meant to reflect the tool surface as it is *right now*,
not as a growing sequence of dated patches layered on top of stale tables (that used to
be this file's structure; it was reorganized specifically to stop that). If you add,
remove, or change the behavior of a tool:

1. Update the relevant per-app table in `docs/ai-tool-surface.md` directly, in place —
   don't just append a new dated note at the top or bottom and leave the table stale.
2. If the change is worth a historical record (a bug fix, a phased rollout, a decision
   that was made a specific way for a specific reason), add a dated entry to
   `docs/ai-tool-surface-changelog.md` instead of to the main doc.
3. If the change touches a tool tier list, update both the C# and `entry.ts` sides (see
   above) and double check the doc's tier description still matches.

Do not resurrect `docs/tool-surface-todo.md` as a place to track gaps — it's retired
(see the changelog) specifically because a checklist like that drifted badly out of
sync with actual implementation state in the past.

## Long rationale comments live in a companion `<File>.md`, not inline

Source comments in this repo come in two flavors. A short "what this does/why" note
stays inline, right where it applies. Longer historical/rationale material — confirmed
repro steps, "we tried X, it didn't work, here's why", design decisions with a real
trade-off — moves into a companion file sitting next to the source file, named
`<SourceFile>.md` (e.g. `TaskPaneHost.cs` → `TaskPaneHost.cs.md`), with a short inline
comment left behind pointing at it (`// See TaskPaneHost.cs.md.`). This keeps the
source readable without losing the reasoning.

**Hard rule: a file's companion `.md` must be self-contained — never point at a
different file's `.md`.** Even when two files share genuinely identical rationale (e.g.
Word's and Excel's/PowerPoint's `TaskPaneHost.cs`, which all share the same COM-timing
construction quirk), each gets its own local `.md` with the full text, not a pointer
into another app's companion file. A one-line "kept in sync with
`OtherApp/OtherFile.cs.md`'s identical note" cross-reference is fine as a bonus, but
never as the only record — a cross-file-only pointer breaks the moment either file is
split, renamed, or the comment is trimmed further on one side but not the other (this
exact bug was caught and fixed once already for `PowerPointAiAddIn/TaskPaneHost.cs`,
commit `6947058`). When you touch a file that has a `.md` companion and the underlying
rationale changes, update both files in the same commit — don't let the inline pointer
outlive the thing it used to say.

## Outlook specifically: native query APIs are mandatory, not a nice-to-have

Outlook's read tools (`list_emails`, `search_emails`, `list_events`, etc.) must use
Outlook's native query surfaces — `Folder.GetTable` (an in-memory rowset, no per-item
COM object per row) or `Items.Restrict("@SQL=" + DASL)` — rather than iterating
`Items` and inspecting each item in a loop. This isn't a style preference: a real
performance incident (documented in `docs/ai-tool-surface-changelog.md`'s 2026-08-27
entries) hit the same underlying problem in Word first — `find_text`/`get_headings`'s
first cut used positional `Paragraphs[i]` indexing, which isn't a real array access in
Word's COM object model, so each indexed read re-walked the document from the start,
producing an effectively O(n²) scan. On a large document this froze Word visibly,
because the automation call runs synchronously on Word's own UI thread and can't pump
its message loop until the call returns. Outlook's mailbox/calendar collections have
the same trap (a linear per-item COM scan over thousands of mail items is exactly this
mistake, just in a different app), which is why the native-query rule is written down
here as a hard requirement rather than left to be rediscovered: a capped linear scan is
allowed only as an explicit fallback when `Restrict` itself rejects a filter, never as
the default path.
