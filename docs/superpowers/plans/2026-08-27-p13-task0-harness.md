# P13 · Task 0 — Build the verification harness

**Part of:** `2026-08-27-phase1-3-file-split.md` (Phases 1+3). Do this task **first**; Tasks 1–3 depend on it.

**What this task produces:** three baseline files on disk that later tasks diff against. It changes **no source code**.

**Time:** ~5 minutes. **Risk:** none — read-only.

---

## Why this exists

Tasks 1–3 move ~200 members between files. Git renders that as hundreds of deletions plus hundreds of additions, and you cannot eyeball it for "did a method get dropped?".

The check: list every member declaration, strip line numbers, sort. That makes the list **order-independent** — which is exactly what you need when reordering is the whole point. If the sorted list is identical before and after, nothing was lost, duplicated, or altered.

---

## Step 1 — Confirm your starting state is clean

```bash
cd /c/dev/officeoffice
git status --short | grep -E "Tools\.cs|\.csproj" || echo "no tool/csproj changes pending"
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo -v q 2>&1 | tail -3
```

**Expected:** tests report `Passed! ... Total: 114`.

- **If the test count is not 114**, stop. Something landed since this plan was written. Re-read `docs/superpowers/plans/STATUS.md` and reconcile before continuing.
- **If `*Tools.cs` files show as modified**, commit or stash them first. Starting from a dirty tool file makes every later diff untrustworthy.

---

## Step 2 — Capture the baselines

> **Use `.split-work/`, not `/tmp`.** Under Git Bash on Windows, MSYS rewrites `/tmp` when it appears as a *command-line argument* to a native program, but **not** inside a Python string literal — where it resolves to a nonexistent `C:\tmp` and fails. A repo-relative directory behaves identically from both. `.split-work/` is already in `.gitignore`.

```bash
cd /c/dev/officeoffice
mkdir -p .split-work
inv() { grep -hnE '^        (private|public|internal)' "$@" | sed -E 's/^[0-9]+:[[:space:]]*//' | sort; }

inv WordAiAddIn/WordTools.cs             > .split-work/word.before.txt
inv ExcelAiAddIn/ExcelTools.cs           > .split-work/excel.before.txt
inv PowerPointAiAddIn/PowerPointTools.cs > .split-work/ppt.before.txt

wc -l .split-work/word.before.txt .split-work/excel.before.txt .split-work/ppt.before.txt
```

**Expected output** (member counts):

```
   63 .split-work/word.before.txt
   83 .split-work/excel.before.txt
   65 .split-work/ppt.before.txt
```

- **If a count differs from the above**, the code has changed since this plan was written. That is fine — **your freshly captured numbers are the correct baseline**. Write down what you actually got and use that. Do not try to force the numbers to match this document.
- Note the `-h` in `grep -hnE`: it suppresses filename prefixes, so the same function works on a glob of several files later. Do not drop it.

---

## Step 3 — Record the line and csproj baselines

```bash
cd /c/dev/officeoffice
wc -l WordAiAddIn/WordTools.cs ExcelAiAddIn/ExcelTools.cs PowerPointAiAddIn/PowerPointTools.cs
grep -c "<Compile" WordAiAddIn/WordAiAddIn.csproj ExcelAiAddIn/ExcelAiAddIn.csproj PowerPointAiAddIn/PowerPointAiAddIn.csproj
```

**Expected:** 2531 / 2183 / 1820 lines; `6` Compile entries in each csproj.

Write these down — Task 4 reconciles against them.

---

## Step 4 — Prove the check actually catches a dropped member

Do not skip this. A verification step you have not tested is not a verification step.

```bash
cd /c/dev/officeoffice
# Simulate the "after" state, then delete one member from it.
cp .split-work/ppt.before.txt .split-work/ppt.broken.txt
grep -v "ModeLabel" .split-work/ppt.broken.txt > .split-work/ppt.broken2.txt

diff .split-work/ppt.before.txt .split-work/ppt.broken2.txt > /dev/null \
  && echo "FAIL: harness did NOT catch a dropped member" \
  || echo "PASS: harness catches a dropped member"
```

**Expected:** `PASS: harness catches a dropped member`.

If you get `FAIL`, the `inv` function is wrong — most likely a copy-paste error in the `grep`/`sed`. Re-copy Step 2 verbatim.

---

## Definition of done

- [ ] `.split-work/word.before.txt`, `.split-work/excel.before.txt`, `.split-work/ppt.before.txt` all exist and are non-empty.
- [ ] You have recorded the three member counts, the three line counts, and the three csproj Compile counts.
- [ ] Step 4 printed `PASS`.
- [ ] No source file was modified (`git status --short` shows nothing new).

**Nothing to commit in this task.** The baseline files live in `.split-work/` and are throwaway.

> **If your shell session ends before Task 3 finishes**, `.split-work/` may be stale. Just re-run Step 2 against the *current* state of any file you have not split yet — the baseline only needs to be captured before *that particular file* is touched.

**Next:** `2026-08-27-p13-task1-powerpoint.md`
