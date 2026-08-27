# P13 · Task 4 — Reconcile and record

**Part of:** `2026-08-27-phase1-3-file-split.md` (Phases 1+3).
**Prerequisites:** Tasks 1, 2 and 3 all done and committed.

**What this task does:** proves the three splits together did not lose anything, then writes down what changed so the next person is not surprised. **No source code changes** beyond documentation.

**Time:** ~20 minutes.

---

## Step 1 — Full rebuild, Debug *and* Release

Release matters: `deploy/package.ps1` builds Release, and **only Release signs the manifests**. A Release-only break would not show up until someone tries to package.

```bash
cd /c/dev/officeoffice
MSBUILD="/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe"
for CFG in Debug Release; do
  for APP in Word Excel PowerPoint; do
    echo "=== $APP $CFG ==="
    "$MSBUILD" ${APP}AiAddIn/${APP}AiAddIn.csproj -t:Build -p:Configuration=$CFG -nologo -v:minimal 2>&1 | tail -4
  done
done
```

**Expected:** six `-> ...dll` lines, no `error`.

`MSB3061 ... Access to the path ... is denied ... locked by: Microsoft Word/Excel/PowerPoint` is a **warning**, not a failure — that Office app is open. The build succeeded if the `-> ...dll` line appears.

---

## Step 2 — Tests

```bash
cd /c/dev/officeoffice
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo -v q 2>&1 | tail -3
```

**Expected:** `Total: 114` passed. This phase adds no testable logic, so the count must be **unchanged** — a *higher* number means someone added tests that were not part of this phase, and a lower one means something was lost.

---

## Step 3 — Line accounting

```bash
cd /c/dev/officeoffice
for A in "WordAiAddIn/WordTools" "ExcelAiAddIn/ExcelTools" "PowerPointAiAddIn/PowerPointTools"; do
  echo "$(basename $A): $(cat ${A}*.cs | wc -l) lines across $(ls ${A}*.cs | wc -l) files"
done
```

**Expected, against the Task 0 baselines (2531 / 2183 / 1820):**

| Class | Before | After (approx) | Growth |
|---|---|---|---|
| `WordTools` | 2,543 | ~2,670 | +5% |
| `ExcelTools` | 2,157 | ~2,285 | +6% |
| `PowerPointTools` | 1,683 | 1,819 | +8% |

**Growth of 5–10% is correct and expected** — each new file repeats the `using` block, `namespace {`, `class {` and the two closing braces (~14 lines × 9 extra files). PowerPoint's 1,819 is the measured figure from the validated run.

**If growth exceeds ~12%**, something is duplicated rather than moved. Check for a member appearing in two files:

```bash
cd /c/dev/officeoffice
inv() { grep -hnE '^        (private|public|internal)' $1 | sed -E 's/^[0-9]+:[[:space:]]*//' | sort; }
inv 'WordAiAddIn/WordTools*.cs' | uniq -d
inv 'ExcelAiAddIn/ExcelTools*.cs' | uniq -d
inv 'PowerPointAiAddIn/PowerPointTools*.cs' | uniq -d
```

Any output = a duplicated member. Nothing = clean.

---

## Step 4 — Size ceiling

```bash
cd /c/dev/officeoffice
wc -l WordAiAddIn/WordTools*.cs ExcelAiAddIn/ExcelTools*.cs PowerPointAiAddIn/PowerPointTools*.cs \
  | sort -rn | head -8
```

**Expected:** the largest file (excluding the `total` line) is under ~450 lines.

**If one is over**, split it now — do not leave the phase half-done. Each split task's Step 7 names the specific fallback for its biggest file.

---

## Step 5 — csproj entry counts

```bash
cd /c/dev/officeoffice
grep -c "<Compile" WordAiAddIn/WordAiAddIn.csproj ExcelAiAddIn/ExcelAiAddIn.csproj PowerPointAiAddIn/PowerPointAiAddIn.csproj
```

**Expected:** `15` for each (6 original + 9 new).

Cross-check that every file on disk is actually listed — this catches the silent "exists but not compiled" case:

```bash
cd /c/dev/officeoffice
for A in Word Excel PowerPoint; do
  D=$(ls ${A}AiAddIn/${A}Tools*.cs | wc -l)
  C=$(grep -c "Include=\"${A}Tools" ${A}AiAddIn/${A}AiAddIn.csproj)
  echo "$A: $D files on disk, $C listed in csproj $([ "$D" = "$C" ] && echo OK || echo MISMATCH)"
done
```

**Expected:** three `OK` lines, 10 and 10 each.

---

## Step 6 — Commit history check

```bash
cd /c/dev/officeoffice
git log --oneline -4
git show --stat HEAD~2 | tail -14
```

**Expected:** three commits, one per app, each touching only that app's files plus its csproj.

**If a commit swept in unrelated files** (this repo has had a long-standing backlog of uncommitted work), do not rewrite history to fix it. Note it in the doc update below instead — a truthful record beats a tidy one.

---

## Step 7 — Write the dated note in `docs/ai-tool-surface.md`

Insert immediately **above** the `## Architecture` heading, matching the existing `> **Update YYYY-MM-DD (...)**` convention:

```markdown
> **Update 2026-08-27 (Phases 1+3 - tool files split into partial classes):**
> the three `*Tools.cs` files were split by tool area into `partial class`
> file sets - Word 2,543 lines -> 10 files, Excel 2,157 -> 10, PowerPoint
> 1,683 -> 10, largest file now under 450 lines. **Structure only: no method
> body, tool schema, `entry.ts`, or system prompt changed**, and `dotnet test`
> stayed at 114. Each split was verified by an order-independent member-set
> diff (sorted declaration list identical before and after), which is what
> makes an otherwise unreviewable move-diff trustworthy.
>
> Two things worth knowing before adding code to these projects:
> - The `.csproj` files are **classic format with explicit `<Compile Include>`
>   items** - they do not glob `*.cs`. A new file must be added by hand, or it
>   is silently not compiled and the error points at a call site elsewhere.
> - `RetryTransientCom`/`TransientComHResults` (Word) stayed in the core file
>   rather than moving to `.Charts.cs` with their only current callers - they
>   are general COM-retry infrastructure, not chart code.
>
> Still deferred to Phase 2: the Word/PowerPoint duplication of
> `RetryTransientCom` and `SmartArtLayoutNames`.
```

---

## Step 8 — Update `docs/superpowers/plans/STATUS.md`

Add to the build-commands block:

```markdown
- Each app's tool code is a `partial class` split across ~10 files
  (`WordTools.cs`, `WordTools.Charts.cs`, …). **New files must be added to the
  classic `.csproj` by hand** - it lists `<Compile Include>` explicitly and does
  not glob. `tools/split-partial.py` did the original split and can do further
  ones; it refuses to run unless every member is assigned exactly once.
```

---

## Step 9 — Mark the phases done

In `docs/superpowers/plans/2026-08-27-refactor-proposal.md`, update the Phase 1 and Phase 3 headings to `— **DONE (2026-08-27)**` and add one line under each noting they were merged into a single pass, with a pointer to `2026-08-27-phase1-3-file-split.md`.

In `2026-08-27-phase1-3-file-split.md` itself, add a `> **DONE**` banner at the top, matching the one on the Phase 0 plan.

---

## Step 10 — Commit

```bash
cd /c/dev/officeoffice
git add docs/
git commit -m "docs: record the Phase 1+3 file split and mark both phases done

Notes the classic-csproj gotcha and the RetryTransientCom placement decision
in ai-tool-surface.md and STATUS.md.

P13 Task 4 of docs/superpowers/plans/2026-08-27-phase1-3-file-split.md."
```

---

## Definition of done

- [ ] All three apps build clean in **Debug and Release**.
- [ ] `dotnet test` = 114, unchanged.
- [ ] Line growth 5–10%; `uniq -d` finds no duplicated member.
- [ ] No file over ~450 lines.
- [ ] Three `OK` lines from the disk-vs-csproj cross-check.
- [ ] `ai-tool-surface.md`, `STATUS.md`, and both plan docs updated and committed.

> **A smoke pass is not required for this phase.** Unlike Phase 0, nothing here changes a method body or a COM call site, so a clean compile plus an identical member set genuinely covers it. Worth doing opportunistically if an Office app is already open, but it is not the gate.

**Next:** `2026-08-27-p13-task5-archive.md` (independent — can be done any time)
