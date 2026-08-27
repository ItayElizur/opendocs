# P13 · Task 1 — Split `PowerPointTools.cs`

**Part of:** `2026-08-27-phase1-3-file-split.md` (Phases 1+3).
**Prerequisite:** Task 0 done — `.split-work/ppt.before.txt` exists (65 members).

**Do this file first.** PowerPoint's members are already almost perfectly contiguous, so this is the least fiddly of the three splits. Learn the mechanics here before touching Excel or Word.

**Result:** 1 file of 1,821 lines → 10 files, largest ~300 lines. Zero logic change.

> **This exact split has already been executed and verified once** (2026-08-27, on a throwaway copy and then in-repo): it compiles, and the member set comes out identical. The config below is the one that was validated. If you follow it literally you should get the same result.

---

## Step 1 — Write the split config

Create `.split-work/ppt.json` with **exactly** this content:

```json
{
  "source": "PowerPointAiAddIn/PowerPointTools.cs",
  "encoding": "utf-8",
  "headerLines": 12,
  "tailLines": 2,
  "groups": {
    "Read":       ["ShapeText","GetDeckContext","ReadSlide","FindTextPpt","ReplaceTextPpt"],
    "Elements":   ["ResolveShape","ResolveNotesBodyPlaceholder","GetSlideNotesText","ApplyAutoDirection","ApplyBulletSetting","SetElementText","SetSlideNotes","AlignmentMap","SetElementStyle","SetElementTransform","ZOrderMap","SetElementOrder","AddTextBox","AddShape","DeleteElement"],
    "Slides":     ["AddSlide","DeleteSlide","MoveSlide","DuplicateSlide"],
    "LayoutAnim": ["SlideLayoutMap","ResolveCustomLayout","SetSlideLayout","TransitionEffectMap","SetSlideTransition","AnimationEffectMap","AnimationTriggerMap","AddAnimation","ReadAnimations","EditAnimation"],
    "Styling":    ["SetElementFill","SetElementStroke","SetSlideBackground","UngroupElement"],
    "Tables":     ["AddTable","ResolveTable","EditTableCell","EditTableStructure","EditTableStyle"],
    "Charts":     ["AddChartPpt","PptLegendPositions","EditChartPpt"],
    "SmartArt":   ["ResolveSmartArtLayout","ListSmartArtShapesOnSlide","ResolveSmartArtOnSlide","ResolveSmartArtGalleryItem","ReadSmartArt","ReadOneSmartArt","EditSmartArt","AddSmartArt"],
    "Images":     ["CropImage","ReplaceImagePpt","SetPictureOpacity"]
  }
}
```

**What the config means:**
- `headerLines: 12` — lines 1–12 are the `using`s, `namespace {`, and `public static class PowerPointTools {`. Copied into every generated file.
- `tailLines: 2` — the final `    }` and `}` that close the class and namespace.
- Members **not** listed anywhere **stay in `PowerPointTools.cs`**. Only 8 do: `ModeByDoc`, `SetMode`, `ModeFor`, `AlwaysAllowedTools`, `Execute`, `IsMutationAllowed`, `ModeLabel`, `ActivePresentation`. You do not list them.

---

## Step 2 — Dry run

```bash
cd /c/dev/officeoffice
python tools/split-partial.py .split-work/ppt.json --dry-run
```

**Expected:**

```
source        : PowerPointAiAddIn/PowerPointTools.cs (1821 lines, 65 members)
stays in core : 8 members
  .Read.cs          5 members
  .Elements.cs      15 members
  .Slides.cs        4 members
  .LayoutAnim.cs    10 members
  .Styling.cs       4 members
  .Tables.cs        5 members
  .Charts.cs        3 members
  .SmartArt.cs      8 members
  .Images.cs        3 members
```

**If instead you see `REFUSING TO SPLIT`**, the script found a name in the config that does not exist in the source, or one assigned to two groups. It prints exactly which. Fix the config; do not edit the script. This guard is the whole reason the split is scripted.

---

## Step 3 — Run it

```bash
cd /c/dev/officeoffice
python tools/split-partial.py .split-work/ppt.json
```

It writes 9 new files, rewrites `PowerPointTools.cs` down to the 8 core members, and prints the `<Compile>` lines you need next.

---

## Step 4 — Add the 9 files to the csproj

**This is the step that bites.** The `.csproj` is classic-format and lists every source file explicitly — it does **not** glob `*.cs`. A file on disk but missing here is silently not compiled, and the error you get points at a *call site in another file*, not at the file you forgot.

```bash
cd /c/dev/officeoffice
python - <<'PYEOF'
import io
p = "PowerPointAiAddIn/PowerPointAiAddIn.csproj"
t = io.open(p, encoding="utf-8-sig").read()
anchor = '    <Compile Include="PowerPointTools.cs">\n      <SubType>Code</SubType>\n    </Compile>\n'
assert anchor in t, "anchor not found - open the csproj and add the entries by hand"
parts = ["Read","Elements","Slides","LayoutAnim","Styling","Tables","Charts","SmartArt","Images"]
add = "".join(
  '    <Compile Include="PowerPointTools.%s.cs">\n'
  '      <SubType>Code</SubType>\n'
  '      <DependentUpon>PowerPointTools.cs</DependentUpon>\n'
  '    </Compile>\n' % s for s in parts)
io.open(p, "w", encoding="utf-8-sig").write(t.replace(anchor, anchor + add, 1))
print("csproj: added", len(parts), "Compile entries")
PYEOF

grep -c "<Compile" PowerPointAiAddIn/PowerPointAiAddIn.csproj
```

**Expected:** `csproj: added 9 Compile entries`, then `15` (was 6).

`<DependentUpon>` nests the parts under the parent in Visual Studio's Solution Explorer — cosmetic, but it matches how `ThisAddIn.Designer.cs` already nests.

---

## Step 5 — Build

```bash
cd /c/dev/officeoffice
"/c/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/amd64/MSBuild.exe" \
  PowerPointAiAddIn/PowerPointAiAddIn.csproj -t:Build -p:Configuration=Debug -nologo -v:minimal 2>&1 | tail -10
```

**Expected:** ends with `PowerPointAiAddIn -> ...\PowerPointAiAddIn.dll` and no `error`.

- **`warning MSB3061: Unable to delete ... WebView2Loader.dll ... locked by: Microsoft PowerPoint`** — harmless. PowerPoint is open. The build still succeeded if you see the `-> ...dll` line.
- **`error CS0103: The name 'X' does not exist`** — you almost certainly missed a file in Step 4. Re-run the `grep -c "<Compile"` check; it must be 15.

---

## Step 6 — Verify the member set (the load-bearing check)

```bash
cd /c/dev/officeoffice
inv() { grep -hnE '^        (private|public|internal)' $1 | sed -E 's/^[0-9]+:[[:space:]]*//' | sort; }
inv 'PowerPointAiAddIn/PowerPointTools*.cs' > .split-work/ppt.after.txt

diff .split-work/ppt.before.txt .split-work/ppt.after.txt && echo "MEMBER SET IDENTICAL"
```

**Expected:** `MEMBER SET IDENTICAL`, no diff output.

**If the diff is non-empty, stop and read it.** Lines prefixed `<` were lost; `>` were added. Do not proceed. Recover with:

```bash
git checkout -- PowerPointAiAddIn/PowerPointTools.cs PowerPointAiAddIn/PowerPointAiAddIn.csproj
rm -f PowerPointAiAddIn/PowerPointTools.{Read,Elements,Slides,LayoutAnim,Styling,Tables,Charts,SmartArt,Images}.cs
```

> ⚠️ **Before running that `git checkout`**, check `git status --short PowerPointAiAddIn/PowerPointAiAddIn.csproj`. If the csproj had *pre-existing* uncommitted changes, `git checkout` will destroy them. In that case restore the csproj from a copy you made instead.

---

## Step 7 — Confirm nothing else broke, then commit

```bash
cd /c/dev/officeoffice
dotnet test OfficeAi.Shared.Tests/OfficeAi.Shared.Tests.csproj --nologo -v q 2>&1 | tail -3
wc -l PowerPointAiAddIn/PowerPointTools*.cs | sort -n
```

**Expected:** `Total: 114` passing (unchanged — this task adds no testable logic), and no file over ~310 lines.

Reference sizes from the validated run:

| File | Lines |
|---|---|
| `PowerPointTools.Styling.cs` | 81 |
| `PowerPointTools.Images.cs` | 89 |
| `PowerPointTools.SmartArt.cs` | 96 |
| `PowerPointTools.Slides.cs` | 104 |
| `PowerPointTools.cs` (core) | 127 |
| `PowerPointTools.Tables.cs` | 180 |
| `PowerPointTools.Read.cs` | 253 |
| `PowerPointTools.Elements.cs` | 287 |
| `PowerPointTools.Charts.cs` | 299 |
| `PowerPointTools.LayoutAnim.cs` | 303 |
| **total** | **1,819** (from 1,683 — the +136 is 9 extra copies of the 14-line header/tail) |

Then commit:

```bash
git add PowerPointAiAddIn/PowerPointTools*.cs PowerPointAiAddIn/PowerPointAiAddIn.csproj
git commit -m "refactor(powerpoint): split PowerPointTools into partial class files by tool area

Structure only - no logic change. 1683 lines -> 10 files, largest 303.
Member set verified identical before and after.

P13 Task 1 of docs/superpowers/plans/2026-08-27-phase1-3-file-split.md."
```

> **Check the commit is proportionate:** `git show --stat HEAD`. If it lists files you did not touch in this task, the working tree had unrelated pending changes that got swept in. Not fatal, but say so rather than letting a misleading commit message stand.

---

## Definition of done

- [ ] 10 files exist; `PowerPointTools.cs` is ~127 lines.
- [ ] csproj has 15 `<Compile` entries.
- [ ] PowerPoint add-in builds clean.
- [ ] `MEMBER SET IDENTICAL`.
- [ ] `dotnet test` still 114.
- [ ] One commit, containing only this task's files.

**Next:** `2026-08-27-p13-task2-excel.md`
