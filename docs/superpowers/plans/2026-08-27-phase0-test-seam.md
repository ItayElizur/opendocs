# Phase 0 — Thin Test Seam Implementation Plan

> **DONE (2026-08-27).** All 6 tasks implemented, each its own commit (7 total - Task 2 and Task 3 each split their pure move from their behavior fix, per plan). `dotnet test` went from 23 to 90 passed. All three add-ins verified building clean (Debug) after every task. See `docs/ai-tool-surface.md`'s 2026-08-27 "Phase 0 complete" note for the tool-facing summary, and `git log --oneline` for the commit sequence (`refactor(shared): extract pure text helpers to TextUtil` through this task's doc commit).

> **For agentic workers:** Steps use checkbox (`- [x]`) syntax for tracking. Each Task ends with its own build + test + commit, and is independently revertable.

**Goal:** Get testable logic out of the three untested `*Tools.cs` files and under `dotnet test` coverage **before** Phase 1–4 start moving code around. Phase 0 is finished when every extraction below is covered by xunit tests in `OfficeAi.Shared.Tests` and all three add-ins still build clean.

**Parent plan:** `docs/superpowers/plans/2026-08-27-refactor-proposal.md` (Phase 0).

---

## The constraint that shapes this whole plan

The original Phase 0 sketch said "write unit tests for the pure logic already embedded in these files." **That is not directly possible**, for two verified reasons:

1. **Every target method is `private static`.** Testing them in place would mean `InternalsVisibleTo` plus flipping dozens of methods from `private` to `internal` — a larger and more invasive diff than moving the handful of genuinely pure ones out.
2. **The three add-in projects are VSTO projects, not plain class libraries.** Each carries `<ProjectTypeGuids>{BAA0C2D2-...};{FAE04EC0-...}</ProjectTypeGuids>`, classic (non-SDK) csproj format, signed manifests, and Office Interop references. `OfficeAi.Shared.csproj` already carries a long comment documenting that `dotnet build`/`dotnet test` and full VS `MSBuild.exe` resolve Office PIA references *differently* in this repo — direct evidence that pulling VSTO projects into the `dotnet test` graph is friction, not a formality.

**Therefore Phase 0 is an extraction, not just a test-writing pass:** move the pure logic into `OfficeAi.Shared` (an SDK-style class library that all three add-ins already reference, and that `OfficeAi.Shared.Tests` already references), then test it there.

This means **Phase 0 deliberately overlaps Phase 2's de-duplication goal.** That is intentional and good: the same move that makes a helper testable also removes its duplicate copies. Phase 2 is left with the harder, non-identical maps (chart types, alignment maps) that need per-app adapters.

### Correction to the parent plan

The refactor proposal listed *"paragraph-index resolution logic"* as a Phase 0 candidate. **That was wrong.** `ParagraphIndexResolver` (`WordTools.cs`) holds `Word.Paragraph` state and walks `doc.Paragraphs.First` / `.Next()` — it is COM-bound in both its fields and its constructor, and cannot be unit-tested without the Phase 5 interop seam. It is explicitly **out of scope** here. This plan supersedes that claim.

---

## Baseline (verified 2026-08-27, before any change)

- `dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo` → **23 passed, 0 failed, ~3s**.
- Test stack: xunit 2.9.2, `Microsoft.NET.Test.Sdk` 17.11.1, `net48`.
- `OfficeAi.Shared` is SDK-style `net48`, `LangVersion 7.3`, already references `System.Drawing`, `System.Text.Json`, and the `Office` PIA (`EmbedInteropTypes=true`).
- All three add-ins already `ProjectReference` `OfficeAi.Shared`.

**Tech stack:** C# 7.3 (`LangVersion` is pinned to 7.3 in both `OfficeAi.Shared` and the test project, and the classic app csproj default to it too — **no switch expressions, no `record`, no target-typed `new`, no nullable reference types**). xunit `[Fact]`/`[Theory]`.

**Test conventions** (from the existing `DocSettingsStoreTests.cs` / `ChatStoreTests.cs`): no namespace on test classes, `MethodName_ScenarioBeingTested` naming, `[Fact]` for single cases, comments that explain *why* a non-obvious case exists.

---

## Global Constraints

- **Behavior-preserving only.** Every extraction in this plan is a move, not a rewrite. Tests must pin down *current* behavior — including current bad behavior (see the validation note in Task 2). Any actual behavior change is a separate, later decision, not part of Phase 0.
- **`LangVersion 7.3`** — do not introduce newer syntax; it will compile in neither the shared library nor the app projects.
- After each Task: build **all three** add-ins (they are separate csproj — there is no `.sln`) and run `dotnet test`. A green shared-library test run does **not** prove the add-ins still compile against the moved member.
- Do **not** change any tool's JSON schema, tool description, or system prompt in this phase. Phase 0 is invisible to the model.
- Keep the existing `// why` comments attached to the code as it moves. They are the repo's most valuable asset and are easy to drop during a copy-paste move.

**Build commands** (from `docs/superpowers/plans/STATUS.md`):
```bash
# C# app build (repeat for Excel and PowerPoint)
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" \
  WordAiAddIn/WordAiAddIn.csproj -t:Build -p:Configuration=Debug -nologo -v:minimal

# Shared library tests
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo
```

---

## Extraction inventory (verified by diff, 2026-08-27)

Line numbers are **as of today** and will shift as Tasks land — locate by method name, not by line.

| Helper | Locations | Verified status |
|---|---|---|
| `ColumnLetter(int)` | `WordTools.cs:254`, `ExcelTools.cs:1366`, **and `PowerPointTools.cs:1337` (missed in the original audit - found during implementation, same body, same "not shared directly" comment)** | **Byte-for-byte identical** across all three |
| `HexToWdColor` / `HexToOleColor` / `HexToOle` | `WordTools.cs:2182`, `ExcelTools.cs:1639`, `PowerPointTools.cs:1098` | **Identical bodies**; differ only in return cast (`Word.WdColor` vs `int`) |
| `ReplaceAllOccurrences` / `...Ppt` | `ExcelTools.cs:517`, `PowerPointTools.cs:313` | Identical except `System.Text.StringBuilder` vs `StringBuilder` (using-directive only) |
| `ShapeTypeMap` | `ExcelTools.cs:43`, `PowerPointTools.cs:665` | Both `Dictionary<string, Microsoft.Office.Core.MsoAutoShapeType>` — the PIA `OfficeAi.Shared` **already references**. Gated on the Task 1 spike. |
| `ValidateRequired` | `WordTools.cs:1842`, `ExcelTools.cs:982` | Identical logic; differ in the `RequiredFields` dictionary used and the error noun ("Command" vs "Operation") |
| `CountOccurrencesPpt` | `PowerPointTools.cs:299` | Single copy, fully pure |
| `IsRtlMajority` | `PowerPointTools.cs:431` | Single copy, fully pure (char arithmetic only) |
| `HtmlEscape` | `WordTools.cs:1617` | Single copy, pure |
| `ResolveImageSize` | `WordTools.cs:2560` | Single copy, pure float math |
| `ValidateKnownFields` | `WordTools.cs:2121` | Single copy, pure |
| `JsonValueToObject` | `ExcelTools.cs:1124` | Single copy, pure (`JsonElement` → `object`) |

**Explicitly out of scope for Phase 0** (COM-bound, needs the Phase 5 seam): `ParagraphIndexResolver`, `NativeFindInSheet`, `NativeFindReplaceInSheet`, `ResolveTargetParagraphs`, `ReadBlockAsHtml`, `FindDataTables`' COM half, and every `ToolResult`-returning tool method.

---

### Task 1: Create the shared home for pure text helpers

**Files:** Create `OfficeAi.Shared/TextUtil.cs`; test `OfficeAi.Shared.Tests/TextUtilTests.cs`.

> **The embedded-interop spike has already been run (2026-08-27) — see "Spike result" below.** Task 4 no longer needs gating; its design has been corrected to match what the spike proved. Step 1 of this Task is therefore already resolved and is recorded rather than repeated.

#### Spike result (run and reverted, 2026-08-27)

`OfficeAi.Shared` references the `Office` PIA with `EmbedInteropTypes=true`, and so do the app projects. Three probes were compiled against the real toolchain (MSBuild for the VSTO project, `dotnet test` for the test project):

| Probe | Shape | Result |
|---|---|---|
| A | `Dictionary<string, MsoAutoShapeType>` exposed from `OfficeAi.Shared`, consumed by `ExcelTools.cs` | ❌ **`CS1769`** — *"cannot be used across assembly boundaries because it has a generic type argument that is an embedded interop type"* |
| B | Bare `MsoAutoShapeType` return value across the boundary | ✅ compiles |
| C | `Dictionary<string, int>` + app-side cast to `MsoAutoShapeType` | ✅ compiles; ✅ also consumable from `OfficeAi.Shared.Tests` **without that project referencing the Office PIA at all** |

**Conclusion:** the originally-planned shape for Task 4 is impossible, but the goal is still achievable. An embedded interop type cannot appear as a *generic type argument* on a public member crossing the boundary — but it can appear bare, and it can be carried as its underlying `int`. Task 4 below is rewritten around probe C.

Note this lands on exactly the same pattern Task 2 already chose independently for color (`int` in the shared library, app-side cast to the app's own enum). Two unrelated extractions converging on the same shape is a good sign it is the right seam: **`OfficeAi.Shared` never exposes an Office interop type in a generic; the app casts at the call site.** Make that the written rule in Task 6.

- [x] **Step 1: Spike embedded-interop type equivalence** — done, results above. Spike code was fully reverted (`dotnet test` back to the 23-test baseline, Excel rebuilds clean).

- [x] **Step 2: Create `TextUtil` with the two verified-identical text helpers**

Create `OfficeAi.Shared/TextUtil.cs`. Move `ColumnLetter` from Word/Excel and `ReplaceAllOccurrences` from Excel/PowerPoint verbatim — bodies unchanged, only `private static` → `public static` and the `StringBuilder` reference fully qualified:

```csharp
using System;
using System.Text;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Pure string helpers shared by the Word/Excel/PowerPoint tool layers.
    /// Extracted here (Phase 0) so they are unit-testable at all - the
    /// *Tools.cs files live in VSTO projects whose private members no test
    /// project can reach. Every method here must stay free of COM types.
    /// </summary>
    public static class TextUtil
    {
        public static string ColumnLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                result = (char)('A' + rem) + result;
                col = (col - 1) / 26;
            }
            return result;
        }

        // .NET Framework 4.8 has no String.Replace(string, string, StringComparison)
        // overload - hence the hand-rolled scan.
        public static string ReplaceAllOccurrences(string input, string find, string replace, StringComparison comparison)
        {
            if (comparison == StringComparison.Ordinal) return input.Replace(find, replace);
            var sb = new StringBuilder();
            int pos = 0;
            while (true)
            {
                int idx = input.IndexOf(find, pos, comparison);
                if (idx < 0) { sb.Append(input, pos, input.Length - pos); break; }
                sb.Append(input, pos, idx - pos);
                sb.Append(replace);
                pos = idx + find.Length;
            }
            return sb.ToString();
        }

        public static int CountOccurrences(string haystack, string needle, StringComparison comparison)
        {
            if (string.IsNullOrEmpty(needle)) return 0;
            int count = 0, pos = 0;
            while (true)
            {
                int idx = haystack.IndexOf(needle, pos, comparison);
                if (idx < 0) break;
                count++;
                pos = idx + needle.Length;
            }
            return count;
        }

        // MOVE VERBATIM from PowerPointTools.cs IsRtlMajority - body omitted
        // here deliberately. It contains \uXXXX char-range escapes for the
        // Hebrew/Arabic blocks; copy them from the source file exactly rather
        // than retyping, and keep the per-range trailing comments.
        public static bool IsRtlMajority(string text) { /* ...unchanged... */ }
    }
}
```

- [x] **Step 3: Delete the originals and repoint every call site**

- `WordTools.cs`: delete `ColumnLetter`; call `TextUtil.ColumnLetter`.
- `ExcelTools.cs`: delete `ColumnLetter` and `ReplaceAllOccurrences`; call `TextUtil.*`.
- `PowerPointTools.cs`: delete `ReplaceAllOccurrencesPpt`, `CountOccurrencesPpt`, `IsRtlMajority`; call `TextUtil.ReplaceAllOccurrences` / `CountOccurrences` / `IsRtlMajority`.

All three files already have `using OfficeAi.Shared;`. Verify no stale references remain:
```bash
grep -n "ColumnLetter\|ReplaceAllOccurrences\|CountOccurrencesPpt\|IsRtlMajority" \
  WordAiAddIn/WordTools.cs ExcelAiAddIn/ExcelTools.cs PowerPointAiAddIn/PowerPointTools.cs
```
Every hit should now be a `TextUtil.`-prefixed call.

- [x] **Step 4: Write `TextUtilTests.cs`**

Cover the behavior that actually carries risk, not just happy paths:

- `ColumnLetter`: `1 → "A"`, `26 → "Z"`, `27 → "AA"`, `52 → "AZ"`, `53 → "BA"`, `702 → "ZZ"`, `703 → "AAA"`. Also `0 → ""` (current behavior for a non-positive input — **pin it, don't fix it**).
- `ReplaceAllOccurrences`: ordinal vs. `OrdinalIgnoreCase`; a replacement string that *contains* the search term (`"a" → "aa"` must not loop forever — this is the single most important case, since `pos` advances by `find.Length`, not by the replacement's); no-match returns input unchanged; empty input.
- `CountOccurrences`: overlapping candidates (`"aaa"`, find `"aa"` → **1**, not 2, because `pos` skips past the match — pin the current non-overlapping semantics); empty needle → `0`; case-insensitive counting.
- `IsRtlMajority`: pure-Hebrew → true; pure-English → false; empty/null → false; **digits and punctuation only → false** (no letters at all); a 50/50 mix → true (the rule is `rtl >= ltr`, so ties go RTL — pin it deliberately); Hebrew with Latin punctuation.

- [x] **Step 5: Build all three add-ins + run tests, then commit**

```bash
git add OfficeAi.Shared/TextUtil.cs OfficeAi.Shared.Tests/TextUtilTests.cs \
        WordAiAddIn/WordTools.cs ExcelAiAddIn/ExcelTools.cs PowerPointAiAddIn/PowerPointTools.cs
git commit -m "refactor(shared): extract pure text helpers to TextUtil, with tests"
```

---

### Task 2: Color conversion

**Files:** Create `OfficeAi.Shared/ColorUtil.cs`; test `OfficeAi.Shared.Tests/ColorUtilTests.cs`.

**Interfaces:** One shared `HexToOle(string) → int`. Word's call sites keep their `Word.WdColor` type by casting at the call site (`(Word.WdColor)ColorUtil.HexToOle(hex)`) — the cast is the *only* difference between the three current copies, so it stays app-side and nothing app-specific enters the shared library.

- [x] **Step 1: Create `ColorUtil.HexToOle`**

```csharp
public static class ColorUtil
{
    /// <summary>
    /// "#RRGGBB" (or "RRGGBB") to the OLE/BGR integer Office's COM APIs want.
    /// Note ColorTranslator.ToOle byte-swaps to BGR - do NOT "simplify" this
    /// to a plain (r &lt;&lt; 16) | (g &lt;&lt; 8) | b, which is the opposite order.
    /// </summary>
    public static int HexToOle(string hex)
    {
        hex = hex.TrimStart('#');
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
    }
}
```

- [x] **Step 2: Delete all three originals, repoint call sites**

`WordTools.cs` `HexToWdColor` → delete; each of its call sites becomes `(Word.WdColor)ColorUtil.HexToOle(...)`. `ExcelTools.cs` `HexToOleColor` and `PowerPointTools.cs` `HexToOle` → delete; call sites become `ColorUtil.HexToOle(...)`.

- [x] **Step 3: Write `ColorUtilTests.cs` against the moved-but-unchanged behavior**

Still a pure move at this point — Step 4 is where behavior changes, as its own commit.

- Known-value round trips: `"#FF0000"` (red) → `0x0000FF`, `"#0000FF"` (blue) → `0xFF0000`, `"#000000"` → `0`, `"#FFFFFF"` → `0xFFFFFF`. These are the tests that would actually catch a BGR/RGB swap regression.
- `"FF0000"` without the leading `#` → same as with it.
- Lowercase `"#ff0000"` → same as uppercase.
- Current failure modes, asserted as-is for now: `Assert.Throws<ArgumentOutOfRangeException>` for `"#abc"`; `Assert.Throws<FormatException>` for `"#GGGGGG"`. Step 4 rewrites these two.

- [x] **Step 4: Commit the pure move, then fix the validation as a SEPARATE commit**

```bash
git commit -m "refactor(shared): extract hex-to-OLE color conversion to ColorUtil, with tests"
```

Now the behavior fix, on its own so it reverts independently of the extraction if it causes trouble.

**The problem being fixed.** Today all three copies have no input validation, so a malformed color reaches the model as a raw .NET exception message via each tool's `catch (Exception ex)`:
- `"#abc"` (3-digit shorthand) → `"Index and length must refer to a location within the string."`
- `"#GGGGGG"` → `"Could not find any recognizable digits."`

Neither names the parameter, the offending value, or the expected format — so the model cannot tell what to send instead, and will often retry the same malformed value. This is the "silent no-op / unhelpful error" class the tool-surface audit already flags as the worst kind of tool defect.

**The fix.** Accept 3-digit shorthand (a real convenience models reach for), and reject everything else with an actionable message:

```csharp
public static int HexToOle(string hex)
{
    if (hex == null) throw new ArgumentException("Color is required, e.g. \"#RRGGBB\".", nameof(hex));
    string h = hex.Trim().TrimStart('#');

    // "abc" is the widely-used CSS shorthand for "aabbcc" - accepted because a
    // model asked for "a light grey" will often produce it, and the old code
    // failed it with an opaque Substring error rather than a usable message.
    if (h.Length == 3)
        h = new string(new[] { h[0], h[0], h[1], h[1], h[2], h[2] });

    if (h.Length != 6 || !IsHexDigits(h))
        throw new ArgumentException(
            "Invalid color \"" + hex + "\". Expected 6-digit hex \"#RRGGBB\" (or 3-digit \"#RGB\"), e.g. \"#1A73E8\".",
            nameof(hex));

    int r = Convert.ToInt32(h.Substring(0, 2), 16);
    int g = Convert.ToInt32(h.Substring(2, 2), 16);
    int b = Convert.ToInt32(h.Substring(4, 2), 16);
    return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
}

private static bool IsHexDigits(string s)
{
    foreach (char c in s)
    {
        bool ok = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
        if (!ok) return false;
    }
    return true;
}
```

Update the Step 3 tests: `"#abc"` now equals `"#aabbcc"`; `"#GGGGGG"`, `""`, `null`, `"#12345"` all throw `ArgumentException` whose message contains the offending value. Also add `" #FF0000 "` (surrounding whitespace) → same as trimmed, since `Trim()` is new here too.

**Blast radius — check before committing.** Every caller currently passes a hex string straight from model-supplied JSON, and every call site is already inside a `try/catch` that reports the exception message as a tool error, so a changed exception *type* cannot crash a tool. Confirm the call-site inventory is unchanged, though:
```bash
grep -rn "HexToOle\|HexToWdColor\|HexToOleColor" --include=*.cs .
```
Every hit should be a `ColorUtil.HexToOle` call after Step 2.

```bash
git commit -m "fix(shared): accept 3-digit hex and give an actionable error for malformed colors"
```

---

### Task 3: Tool-argument validation

**Files:** Create `OfficeAi.Shared/ToolArgs.cs`; test `OfficeAi.Shared.Tests/ToolArgsTests.cs`.

**Interfaces:** The two copies of `ValidateRequired` differ only in the dictionary they consult and the error noun. Parameterize both:

```csharp
public static void ValidateRequired(
    string kind,
    JsonElement element,
    IReadOnlyDictionary<string, string[]> requiredFields,
    string noun)   // "Command" (Word) or "Operation" (Excel)
{
    string[] required;
    if (!requiredFields.TryGetValue(kind, out required)) return;
    foreach (string f in required)
    {
        if (!element.TryGetProperty(f, out _))
            throw new ArgumentException(noun + " \"" + kind + "\" is missing required field \"" + f + "\".");
    }
}
```

The `RequiredFields` dictionaries themselves **stay in their app files** — they are app-specific data (and each is explicitly documented as mirroring that app's `entry.ts` schema, a pairing that must not be split up). Change their declared type to `Dictionary<string, string[]>` if needed so they satisfy `IReadOnlyDictionary`.

Also move `ValidateKnownFields` (`WordTools.cs:2121`) here unchanged — single copy, but pure and cheap to cover.

- [x] **Step 1:** Create `ToolArgs.cs` with both methods.
- [x] **Step 2:** Repoint Word (`noun: "Command"`) and Excel (`noun: "Operation"`). **Verify the exact error strings are unchanged** — `ApplyCommands`/`ProposeOperations` surface them verbatim to the model, and the tool-surface doc treats the specific "missing required field" wording as the contract that turns a malformed call into an actionable error.
- [x] **Step 3:** Write `ToolArgsTests.cs`: an unknown `kind` is a silent no-op (not a throw — this is deliberate current behavior for kinds with no required fields, and easy to "fix" into a regression); all-fields-present passes; one missing field throws with the exact expected message including the right noun; `ValidateKnownFields` accepts a known set and throws listing valid fields for an unknown one.
- [x] **Step 4:** Build all three, test, commit the pure move.

```bash
git commit -m "refactor(shared): extract tool-argument validation to ToolArgs, with tests"
```

- [x] **Step 5: Fix present-but-null required fields — opt-in per field, as a SEPARATE commit**

> **Do not make this a blanket rule. A naive version breaks a working feature.** The obvious fix — "treat JSON `null` as missing" — was checked against the actual handlers and is **wrong as a global rule**:
>
> | Call | Today | Verdict |
> |---|---|---|
> | `set_cell` with `value: null` | `JsonValueToObject` returns C# `null` → `Range.Value2 = null` → **clears the cell** | **Legitimate.** Rejecting it removes the only way to clear a cell via `set_cell`. |
> | `set_bold` / `set_italic` with `value: null` | `cmd.GetProperty("value").GetBoolean()` throws `InvalidOperationException`: *"The requested operation requires an element of type 'Boolean', but the target element has type 'Null'"* | **Bad.** An opaque runtime error where the clean "missing required field" message belongs. |
>
> So null is meaningful for some required fields and meaningless for others. Make rejection **opt-in per field**, not global.

Add a second, optional parameter carrying the fields for which null is invalid:

```csharp
public static void ValidateRequired(
    string kind,
    JsonElement element,
    IReadOnlyDictionary<string, string[]> requiredFields,
    string noun,
    IReadOnlyDictionary<string, string[]> nonNullFields = null)
{
    string[] required;
    if (!requiredFields.TryGetValue(kind, out required)) return;
    foreach (string f in required)
    {
        JsonElement value;
        if (!element.TryGetProperty(f, out value))
            throw new ArgumentException(noun + " \"" + kind + "\" is missing required field \"" + f + "\".");

        // Null is a MEANINGFUL value for some fields (Excel's set_cell uses it
        // to clear a cell), so it is only rejected where the owning app has
        // opted in - never globally.
        if (value.ValueKind != JsonValueKind.Null || nonNullFields == null) continue;
        string[] nonNull;
        if (nonNullFields.TryGetValue(kind, out nonNull) && System.Array.IndexOf(nonNull, f) >= 0)
            throw new ArgumentException(
                noun + " \"" + kind + "\" requires a non-null value for field \"" + f + "\".");
    }
}
```

In `WordTools.cs`, add the opt-in table next to the existing `RequiredFields` (same "edit both together" discipline that table already documents):

```csharp
// Fields where an explicit JSON null is a caller error rather than a value.
// Deliberately narrow - only add a field here after confirming null has no
// legitimate meaning for it (Excel's set_cell "value" is the counter-example).
private static readonly Dictionary<string, string[]> NonNullFields = new Dictionary<string, string[]>
{
    ["set_bold"] = new[] { "value" },
    ["set_italic"] = new[] { "value" },
    ["set_heading"] = new[] { "level" },
};
```

**Excel passes no `nonNullFields` table at all** in this step — its `set_cell`/`set_range` null semantics are load-bearing, and no other Excel op has a demonstrated null problem. Leaving it null keeps Excel's behavior bit-for-bit identical.

Tests to add: `set_bold` with `value: null` throws the new non-null message; `set_bold` with `value: false` still **passes** (the regression this risks — `false` is falsy but present and valid); a field null but *not* listed in `NonNullFields` still passes; passing `nonNullFields: null` reproduces the old behavior exactly.

```bash
git commit -m "fix(word): reject explicit nulls for required boolean/level fields with a clear message"
```

---

### Task 4: Shape-type catalog

**Files:** Create `OfficeAi.Shared/ShapeTypes.cs`; test `OfficeAi.Shared.Tests/ShapeTypesTests.cs`.

> **Design corrected by the Task 1 spike.** A `Dictionary<string, MsoAutoShapeType>` **cannot** cross the assembly boundary (`CS1769` — embedded interop type as a generic argument). The map must be `Dictionary<string, int>` in the shared library, with the app casting to `MsoAutoShapeType` at the call site. This is the same split Task 2 uses for color, and the spike confirmed it compiles in the VSTO projects *and* is consumable from the test project without that project referencing the Office PIA.

- [x] **Step 1: Diff the two maps before merging them.**

Do **not** assume the Excel and PowerPoint maps are identical — `PowerPointTools.cs`'s own comment says it was *ported* from Excel and that PowerPoint kept `rectangle`/`oval` as extra aliases for `rect`/`ellipse`. Extract both key sets and compare before choosing the merged set:

```bash
grep -oE '\["[a-zA-Z0-9]+"\] = Microsoft' -A0 ExcelAiAddIn/ExcelTools.cs | sort > /tmp/xl_keys.txt
grep -oE '\["[a-zA-Z0-9]+"\] = Microsoft' -A0 PowerPointAiAddIn/PowerPointTools.cs | sort > /tmp/pp_keys.txt
diff /tmp/xl_keys.txt /tmp/pp_keys.txt
```

The merged map must be the **union**, so no currently-valid name in either app starts erroring. Record any key present in one and not the other in the commit message. Watch for a key that exists in both but maps to *different* `MsoAutoShapeType` values — that would be a genuine conflict needing a decision, not a mechanical merge.

- [x] **Step 2: Create `ShapeTypes.cs` as an `int` map.**

```csharp
/// <summary>
/// Shape-name to MsoAutoShapeType, as int. It is int and not
/// MsoAutoShapeType deliberately: OfficeAi.Shared embeds the Office PIA, and
/// an embedded interop type used as a GENERIC TYPE ARGUMENT cannot cross an
/// assembly boundary (verified: CS1769). Callers cast:
///     (MsoAutoShapeType)ShapeTypes.ByName["rect"]
/// Same shape as ColorUtil: shared library stays int, app casts at the edge.
/// </summary>
public static readonly Dictionary<string, int> ByName =
    new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { /* union of both maps */ };
```

Keep `StringComparer.OrdinalIgnoreCase` (both current copies use it) and carry over the PIA-omission comment about `msoShapePlus`/`msoShapeMathPlus` not existing in this PIA.

- [x] **Step 3:** Delete both originals; repoint `ExcelTools.cs` and `PowerPointTools.cs` to `(Microsoft.Office.Core.MsoAutoShapeType)ShapeTypes.ByName[...]`. Both build their "unknown shapeType" error from `string.Join(", ", ShapeTypeMap.Keys)` — that message now lists the union, which is a **user-visible improvement** (each app previously advertised only its own subset), but confirm each app's `entry.ts` `enum` still matches what the handler accepts, or the schema and the error message will disagree.
- [x] **Step 4:** Write `ShapeTypesTests.cs`. Note the tests assert **keys and lookup behavior**, not enum values — asserting raw ints would be brittle and meaningless, and the key behavior is what actually carries risk: lookup is case-insensitive (`"RoundRect"` == `"roundrect"`); `rectangle`/`oval` aliases resolve to the same int as `rect`/`ellipse`; an unknown key returns false from `TryGetValue`; and the map contains every name listed in both apps' `entry.ts` shape enums — a genuine schema-vs-handler drift guard.
- [x] **Step 5:** Build all three, test, commit.

---

### Task 5: Remaining single-copy pure helpers

**Files:** Extend `OfficeAi.Shared/TextUtil.cs` (`HtmlEscape`) and add `OfficeAi.Shared/GeometryUtil.cs` (`ResolveImageSize`) + `OfficeAi.Shared/JsonUtil.cs` (`JsonValueToObject`); tests alongside.

These are single-copy, so extracting them buys testability rather than de-duplication — lower value than Tasks 1–4. **If Phase 0 is running long, this is the Task to defer**, not to rush.

- [x] **Step 1:** Move `HtmlEscape` (Word) → `TextUtil.HtmlEscape`. Test: `&` `<` `>` `"` escaping, escaping order (`&` must be escaped first or `&lt;` becomes `&amp;lt;` — the classic bug this test exists to catch), empty string, text with no special characters.
- [x] **Step 2:** Move `ResolveImageSize` (Word) → `GeometryUtil`. Test: both dimensions given → both honored verbatim; width only → height scales proportionally; height only → width scales proportionally; neither → natural size returned unchanged. This is the "never distort by defaulting the missing dimension" rule the code comment claims — worth locking down.
- [x] **Step 3:** Move `JsonValueToObject` (Excel) → `JsonUtil`. Test: string/number/true/false map correctly; **null and any other `JsonValueKind` (array, object) → `null`** — pin the current catch-all, which is what makes a nested array silently land as an empty cell rather than throwing.
- [x] **Step 4:** Build all three, test, commit.

---

### Task 6: Document the seam and update project docs

**Files:** Modify `docs/superpowers/plans/STATUS.md`, `docs/ai-tool-surface.md`.

- [x] **Step 1:** Add a dated note to `docs/ai-tool-surface.md`, matching the existing `> **Update YYYY-MM-DD (...)**` convention already used there: what moved to `OfficeAi.Shared`, that no tool *schema* changed, the `CS1769` interop constraint and the int-map pattern it forced, and the two deliberate behavior changes below.
- [x] **Step 2:** Update `STATUS.md`'s build-commands block with the new post-Phase-0 test count, and state that pure helpers now live in `OfficeAi.Shared` and belong there by default.
- [x] **Step 3:** Add a short "where does this code go?" rule near the top of `TextUtil.cs` (or a brief `OfficeAi.Shared/README.md`), covering both rules this phase established:
  1. **Anything free of COM types goes here and gets a test; anything touching `Word.*`/`Excel.*`/`PowerPoint.*` stays in its app's `*Tools.cs` until Phase 5.** Without a written rule, the next helper gets added to a `*Tools.cs` out of habit and the seam quietly stops growing.
  2. **Never expose an Office interop type as a generic type argument from this assembly** — it does not compile in the VSTO projects (`CS1769`). Carry it as `int`/`string` and cast at the app-side call site, as `ColorUtil` and `ShapeTypes` both do. This is non-obvious, costs a build cycle to rediscover, and is exactly the kind of thing the next person will otherwise hit head-on.
- [x] **Step 4:** Commit.

---

## Definition of done

- [x] `dotnet test` green, with test count meaningfully above the 23-test baseline.
- [x] All three add-ins build clean in **Debug and Release** (Release matters — `deploy/package.ps1` builds Release, and only Release signs manifests).
- [x] `grep` confirms no orphaned copies of any extracted helper remain in the three `*Tools.cs` files.
- [x] No `entry.ts`, tool *schema*, or system prompt changed anywhere in this phase.
- [x] **Exactly two behavior changes landed, each as its own revertable commit** — hex-color validation (Task 2 Step 4) and Word's non-null required fields (Task 3 Step 5). Every other commit in this phase is a pure move. If a third behavior change appears in the log, it was not planned; justify or revert it.
- [x] **Excel's `set_cell` with `value: null` still clears the cell.** This is the specific regression the opt-in null design exists to prevent — verify it directly, not by inference.
- [x] A mock-server smoke pass (`FORCE_TOOL:` against a real document per app) confirms at least one tool that consumed each extracted helper still works end-to-end — specifically `format_range` (Excel, hex color), `set_element_fill` (PowerPoint, hex color + shape types), and `apply_commands` with a deliberately missing required field (Word, validation message).
- [x] Both new error messages checked **as the model would see them** — send a malformed color and a null `set_bold` value through the mock server and read the actual tool-result text. The point of these two fixes is the message; a fix whose message is still unhelpful has not landed.

> **Why the smoke pass is not optional:** every claim in this plan is grounded in reading and diffing source, and each Task is verified by compilation plus unit tests — but neither proves the COM call *downstream* of an extracted helper still behaves the same against a live Office instance. That gap is exactly where this project's last two rounds of real bugs lived.
