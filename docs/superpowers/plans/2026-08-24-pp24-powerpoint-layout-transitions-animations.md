# PP-24: PowerPoint Slide Layout, Transitions, and Animations — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Source:** user question, 2026-08-24 ("does powerpoint support ... transitions and animations? ... change slide layout?"). Confirmed absent by direct source read (`grep -i "layout\|transition\|animation" PowerPointAiAddIn/PowerPointTools.cs` — only SmartArt-layout matches, no slide layout/transition/animation support anywhere).

**Goal:** Five new PowerPoint tools — `set_slide_layout`, `set_slide_transition`, `add_animation`, `read_animations`, `edit_animation` — closing all three gaps in one plan, since they share the same object-model neighborhood (`Slide`/`Design`/`TimeLine`) and the same curated-enum-map pattern already proven throughout this codebase.

## Grounding: verified via .NET reflection against the real referenced PIA, not guessed

Every chart-related bug this session (PP-9, PP-23) traced back to guessing a COM signature instead of checking it. This plan does the opposite from the start — every API below was confirmed by loading `Microsoft.Office.Interop.PowerPoint.dll` (the same PIA version `PowerPointAiAddIn.csproj` references) via .NET reflection and inspecting the real members, not recalled from memory or documentation.

**Slide layout** — `_Slide` (the interface behind `PowerPoint.Slide`) has:
- `get_Layout()`/`set_Layout(PpSlideLayout)` — the classic, pre-2007 layout system. `PpSlideLayout` has exactly 37 values (confirmed by enumeration): `ppLayoutTitle`, `ppLayoutText`, `ppLayoutTwoColumnText`, `ppLayoutTable`, `ppLayoutTextAndChart`, `ppLayoutChartAndText`, `ppLayoutOrgchart`, `ppLayoutChart`, `ppLayoutTextAndClipart`, `ppLayoutClipartAndText`, `ppLayoutTitleOnly`, `ppLayoutBlank`, `ppLayoutTextAndObject`, `ppLayoutObjectAndText`, `ppLayoutLargeObject`, `ppLayoutObject`, `ppLayoutTextAndMediaClip`, `ppLayoutMediaClipAndText`, `ppLayoutObjectOverText`, `ppLayoutTextOverObject`, `ppLayoutTextAndTwoObjects`, `ppLayoutTwoObjectsAndText`, `ppLayoutTwoObjectsOverText`, `ppLayoutFourObjects`, `ppLayoutVerticalText`, `ppLayoutClipArtAndVerticalText`, `ppLayoutVerticalTitleAndText`, `ppLayoutVerticalTitleAndTextOverChart`, `ppLayoutTwoObjects`, `ppLayoutObjectAndTwoObjects`, `ppLayoutTwoObjectsAndObject`, `ppLayoutCustom`, `ppLayoutSectionHeader`, `ppLayoutComparison`, `ppLayoutContentWithCaption`, `ppLayoutPictureWithCaption`, `ppLayoutMixed`.
- `get_CustomLayout()`/`set_CustomLayout(CustomLayout)` — the modern, theme-based layout system (what current PowerPoint's "Layout" gallery in the ribbon actually shows). `CustomLayout` has `.Name` (settable/gettable) and `.Index` — same "resolve by name" shape as `Word`'s SmartArt colors/quick-styles (PP-23).
- The custom layouts available to a given slide live at `slide.Design.SlideMaster.CustomLayouts` (a `CustomLayouts` collection with `.Count`/`.Item(i)`/enumerable) — `_Slide.get_Design()` and `Design.get_SlideMaster()` both confirmed to exist; `Presentation.SlideMaster` also exists directly (the presentation's primary master) but a specific slide's own `Design.SlideMaster` is the contextually-correct one to resolve against, since a deck can have more than one theme/design.

**Transitions** — `Slide.SlideShowTransition` (get-only property returning a live, mutable object) has: `get_EntryEffect()`/`set_EntryEffect(PpEntryEffect)` (189 values total — the actual named subset used below was pulled directly from the enum, not guessed: `ppEffectNone`, `ppEffectCut`, `ppEffectFade`, `ppEffectDissolve`, `ppEffectRandom`, `ppEffectWipeLeft/Up/Right/Down`, `ppEffectPushLeft/Up/Right/Down`, `ppEffectCoverLeft/Up/Right/Down`, `ppEffectUncoverLeft/Up/Right/Down`, `ppEffectZoomIn/Out/Center`, `ppEffectCircleOut`, `ppEffectDiamondOut`, `ppEffectSplitHorizontalOut/In`, `ppEffectSplitVerticalOut/In`, `ppEffectWheel1Spoke`, `ppEffectBlindsHorizontal/Vertical`, `ppEffectCheckerboardAcross/Down`), `get_Duration()`/`set_Duration(Single)` (modern, seconds-based — supersedes the older 3-value `PpTransitionSpeed`, which this plan does not expose), `get_AdvanceOnClick()`/`set_AdvanceOnClick(MsoTriState)`, `get_AdvanceOnTime()`/`set_AdvanceOnTime(MsoTriState)`, `get_AdvanceTime()`/`set_AdvanceTime(Single)`.

**Animations** — `Slide.TimeLine` (get-only) → `.MainSequence` (a `Sequence`) → `.AddEffect(Shape, MsoAnimEffect, MsoAnimateByLevel, MsoAnimTriggerType, Int32)` returns an `Effect`. Confirmed members on `Effect`: `.EffectType` (get/set, `MsoAnimEffect`), `.Exit` (get/set, bool — entrance vs. exit animation), `.Shape` (get/set), `.Timing` (get, a `Timing` object), `.Delete()`, `.MoveTo`/`.MoveBefore`/`.MoveAfter` (reordering within the sequence), `.Index`. Confirmed members on `Timing`: `.Duration` (get/set, float seconds), `.TriggerType` (get/set, `MsoAnimTriggerType`), `.TriggerDelayTime` (get/set, float seconds), `.RepeatCount`, `.AutoReverse`. `MsoAnimTriggerType` has exactly 7 values (all confirmed): `msoAnimTriggerNone`, `msoAnimTriggerOnPageClick`, `msoAnimTriggerWithPrevious`, `msoAnimTriggerAfterPrevious`, `msoAnimTriggerOnShapeClick`, `msoAnimTriggerOnMediaBookmark`, `msoAnimTriggerMixed`. `MsoAnimEffect` has 151 values; the curated subset below (`msoAnimEffectAppear`, `Fade`, `Fly`, `FlashOnce`, `Wipe`, `Zoom`, `Dissolve`, `Bounce`, `Spiral`, `Swivel`, `Wheel`, `Split`, `Box`, `Circle`, `Diamond`, `Plus`, `Checkerboard`, `RandomBars`, `GrowAndTurn`, `RiseUp`) was pulled directly from the enumerated list, not recalled.

**What's still genuinely unverified** (flagged honestly, not guessed past): `Sequence.AddEffect`'s trailing `Int32` parameter is documented (Microsoft's own VBA reference, not re-derived here) as the insertion **index** within the sequence, where `-1` appends at the end — this specific behavior was **not** independently re-derived via reflection (reflection gives signatures, not runtime semantics) and needs live-Word confirmation, same risk category as PP-9's `AddChart2` `Anchor:` parameter. Per-effect **direction** (e.g., "wipe from the left" vs. "wipe from the right" for an entrance animation) is a separate property (`Effect.EffectParameters` / `AnimationBehavior`, an object family this plan did not reflect on) — deliberately **out of scope** for `add_animation` v1, same kind of honest scope-narrowing as Word's `NUMBERED_DECIMAL_ALPHA_ROMAN` bullet preset (PP-12).

**Tech Stack:** C# 7.3 / .NET Framework 4.8, statically-typed `Microsoft.Office.Interop.PowerPoint` throughout (unlike Word's chart/SmartArt code, everything touched by this plan — `Slide`, `SlideShowTransition`, `CustomLayout`, `TimeLine`, `Sequence`, `Effect`, `Timing` — is a normal, non-`dynamic` interop type already referenced by this project; no `dynamic` needed anywhere in this plan).

## Global Constraints

- C# 7.3 / .NET Framework 4.8 only — no `using`-declarations, no target-typed `new()`, no switch expressions.
- Five **standalone top-level tools** (not `apply_commands`-style kinds) — matches how this file's other slide-level tools (`add_slide`/`delete_slide`/`move_slide`/`duplicate_slide`) are standalone.
- Curated enum maps (`SlideLayoutMap`, `TransitionEffectMap`, `AnimationEffectMap`, `TriggerTypeMap`) follow this file's own established pattern (`PptChartTypeMap`, `AlignmentMap`, `SmartArtLayoutNames`): unknown key → throw listing the valid keys, never a silent fallback to a default.
- `read_animations` is a new read tool: add to `READER_TOOLS` (`entry.ts`) and `AlwaysAllowedTools` (`PowerPointTools.cs`), matching `get_deck_context`/`read_slide`'s existing precedent.
- 0-based indices at every tool boundary (`slideIndex`, `shapeIndex`, `animationIndex`), converted to 1-based COM collection access at the point of use — the established convention throughout this file.
- No automated tests for COM executor methods (project convention, unchanged). Verification is build + the manual matrix in Task 6.
- Rebuild the bundle and re-run MSBuild after every `entry.ts` change.
- Add every new tool's `toolDisplay` entry (English + Hebrew) in the same edit that adds its schema.

---

### Task 1: `set_slide_layout`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [x] **Step 1: Curated classic-layout map**

```csharp
// PP-24: curated subset of the 37 PpSlideLayout values (confirmed via
// reflection against the real PIA - not the full set, since several
// legacy names like ppLayoutOrgchart/ppLayoutMediaClipAndText are pre-2007
// leftovers unlikely to be what a model means by a layout request). Same
// curated-map-with-throw-on-unknown pattern as PptChartTypeMap.
private static readonly Dictionary<string, PowerPoint.PpSlideLayout> SlideLayoutMap = new Dictionary<string, PowerPoint.PpSlideLayout>
{
    ["title"] = PowerPoint.PpSlideLayout.ppLayoutTitle,
    ["titleOnly"] = PowerPoint.PpSlideLayout.ppLayoutTitleOnly,
    ["blank"] = PowerPoint.PpSlideLayout.ppLayoutBlank,
    ["text"] = PowerPoint.PpSlideLayout.ppLayoutText,
    ["twoColumnText"] = PowerPoint.PpSlideLayout.ppLayoutTwoColumnText,
    ["object"] = PowerPoint.PpSlideLayout.ppLayoutObject,
    ["objectAndText"] = PowerPoint.PpSlideLayout.ppLayoutObjectAndText,
    ["textAndObject"] = PowerPoint.PpSlideLayout.ppLayoutTextAndObject,
    ["twoObjects"] = PowerPoint.PpSlideLayout.ppLayoutTwoObjects,
    ["twoObjectsAndText"] = PowerPoint.PpSlideLayout.ppLayoutTwoObjectsAndText,
    ["fourObjects"] = PowerPoint.PpSlideLayout.ppLayoutFourObjects,
    ["table"] = PowerPoint.PpSlideLayout.ppLayoutTable,
    ["chart"] = PowerPoint.PpSlideLayout.ppLayoutChart,
    ["sectionHeader"] = PowerPoint.PpSlideLayout.ppLayoutSectionHeader,
    ["comparison"] = PowerPoint.PpSlideLayout.ppLayoutComparison,
    ["contentWithCaption"] = PowerPoint.PpSlideLayout.ppLayoutContentWithCaption,
    ["pictureWithCaption"] = PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption,
};
```

- [x] **Step 2: Custom-layout-by-name resolver**, mirroring `ResolveSmartArtGalleryItem`'s substring-match-with-real-names-on-miss shape (Word `WordTools.cs`, ported here as the same pattern, not shared code across assemblies):

```csharp
private static PowerPoint.CustomLayout ResolveCustomLayout(PowerPoint.Slide slide, string query)
{
    PowerPoint.CustomLayout firstMatch = null;
    var namesSeen = new List<string>();
    foreach (PowerPoint.CustomLayout layout in slide.Design.SlideMaster.CustomLayouts)
    {
        namesSeen.Add(layout.Name);
        if (firstMatch == null && layout.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) firstMatch = layout;
    }
    if (firstMatch != null) return firstMatch;
    throw new ArgumentException("set_slide_layout: no custom layout matching '" + query + "' found in this slide's theme. Available: " + string.Join(", ", namesSeen) + ".");
}
```

- [x] **Step 3: Handler**

```csharp
private static ToolResult SetSlideLayout(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    string kind = input.TryGetProperty("kind", out var kindEl) && kindEl.ValueKind == JsonValueKind.String ? kindEl.GetString() : "classic";

    if (kind == "custom")
    {
        string layoutName = input.GetProperty("layoutName").GetString();
        slide.CustomLayout = ResolveCustomLayout(slide, layoutName);
        return new ToolResult { Output = "Slide " + slideIndex + " layout set to custom layout '" + layoutName + "'.", Mutated = true, Summary = "set_slide_layout" };
    }
    if (kind != "classic")
        throw new ArgumentException("set_slide_layout: unknown kind '" + kind + "'. Valid: classic, custom.");

    string layoutKey = input.GetProperty("layout").GetString();
    PowerPoint.PpSlideLayout layoutValue;
    if (!SlideLayoutMap.TryGetValue(layoutKey, out layoutValue))
        throw new ArgumentException("set_slide_layout: unknown layout '" + layoutKey + "'. Valid: " + string.Join(", ", SlideLayoutMap.Keys) + ".");
    slide.Layout = layoutValue;
    return new ToolResult { Output = "Slide " + slideIndex + " layout set to '" + layoutKey + "'.", Mutated = true, Summary = "set_slide_layout" };
}
```

`kind` defaults to `"classic"` when omitted — the simpler, more reliable mechanism (a real enum, not name-matching) is the natural default; `"custom"` is there for decks whose theme defines layouts the classic enum doesn't map onto well (e.g. "Two Content" vs. the theme's own "Comparison" layout may look different across templates).

- [x] **Step 4: Register in `Execute`'s switch.**

- [x] **Step 5: Schema**

```ts
{
  name: 'set_slide_layout',
  description:
    'Changes a slide\'s layout. kind:"classic" (default) uses layout, a fixed set: title, titleOnly, blank, text, twoColumnText, object, objectAndText, textAndObject, twoObjects, twoObjectsAndText, fourObjects, table, chart, sectionHeader, comparison, contentWithCaption, pictureWithCaption. ' +
    'kind:"custom" uses layoutName instead - free text, matched by substring against this presentation\'s own theme layouts (an unmatched name errors listing the real available names).',
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number' },
      kind: { type: 'string', enum: ['classic', 'custom'] },
      layout: {
        type: 'string',
        enum: ['title', 'titleOnly', 'blank', 'text', 'twoColumnText', 'object', 'objectAndText', 'textAndObject', 'twoObjects', 'twoObjectsAndText', 'fourObjects', 'table', 'chart', 'sectionHeader', 'comparison', 'contentWithCaption', 'pictureWithCaption'],
      },
      layoutName: { type: 'string' },
    },
    required: ['slideIndex'],
  },
},
```

**Verification:** builds; a real PowerPoint test applies at least `blank`, `titleOnly`, and `twoObjects`, plus one `kind:"custom"` call using a name taken from `read_slide`'s own theme (needs Task 4's `read_slide` extension, or manually inspect the Layout gallery first).

---

### Task 2: `set_slide_transition`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [x] **Step 1: Curated transition-effect map** (all names below confirmed present in the real `PpEntryEffect` enum via reflection, not recalled):

```csharp
private static readonly Dictionary<string, PowerPoint.PpEntryEffect> TransitionEffectMap = new Dictionary<string, PowerPoint.PpEntryEffect>
{
    ["none"] = PowerPoint.PpEntryEffect.ppEffectNone,
    ["cut"] = PowerPoint.PpEntryEffect.ppEffectCut,
    ["fade"] = PowerPoint.PpEntryEffect.ppEffectFade,
    ["dissolve"] = PowerPoint.PpEntryEffect.ppEffectDissolve,
    ["random"] = PowerPoint.PpEntryEffect.ppEffectRandom,
    ["wipeLeft"] = PowerPoint.PpEntryEffect.ppEffectWipeLeft,
    ["wipeRight"] = PowerPoint.PpEntryEffect.ppEffectWipeRight,
    ["wipeUp"] = PowerPoint.PpEntryEffect.ppEffectWipeUp,
    ["wipeDown"] = PowerPoint.PpEntryEffect.ppEffectWipeDown,
    ["pushLeft"] = PowerPoint.PpEntryEffect.ppEffectPushLeft,
    ["pushRight"] = PowerPoint.PpEntryEffect.ppEffectPushRight,
    ["pushUp"] = PowerPoint.PpEntryEffect.ppEffectPushUp,
    ["pushDown"] = PowerPoint.PpEntryEffect.ppEffectPushDown,
    ["coverLeft"] = PowerPoint.PpEntryEffect.ppEffectCoverLeft,
    ["coverRight"] = PowerPoint.PpEntryEffect.ppEffectCoverRight,
    ["coverUp"] = PowerPoint.PpEntryEffect.ppEffectCoverUp,
    ["coverDown"] = PowerPoint.PpEntryEffect.ppEffectCoverDown,
    ["uncoverLeft"] = PowerPoint.PpEntryEffect.ppEffectUncoverLeft,
    ["uncoverRight"] = PowerPoint.PpEntryEffect.ppEffectUncoverRight,
    ["uncoverUp"] = PowerPoint.PpEntryEffect.ppEffectUncoverUp,
    ["uncoverDown"] = PowerPoint.PpEntryEffect.ppEffectUncoverDown,
    ["zoomIn"] = PowerPoint.PpEntryEffect.ppEffectZoomIn,
    ["zoomOut"] = PowerPoint.PpEntryEffect.ppEffectZoomOut,
    ["zoomCenter"] = PowerPoint.PpEntryEffect.ppEffectZoomCenter,
    ["circle"] = PowerPoint.PpEntryEffect.ppEffectCircleOut,
    ["diamond"] = PowerPoint.PpEntryEffect.ppEffectDiamondOut,
    ["splitHorizontal"] = PowerPoint.PpEntryEffect.ppEffectSplitHorizontalOut,
    ["splitVertical"] = PowerPoint.PpEntryEffect.ppEffectSplitVerticalOut,
    ["wheel"] = PowerPoint.PpEntryEffect.ppEffectWheel1Spoke,
    ["blindsHorizontal"] = PowerPoint.PpEntryEffect.ppEffectBlindsHorizontal,
    ["blindsVertical"] = PowerPoint.PpEntryEffect.ppEffectBlindsVertical,
    ["checkerboard"] = PowerPoint.PpEntryEffect.ppEffectCheckerboardAcross,
};
```

- [x] **Step 2: Handler**

```csharp
private static ToolResult SetSlideTransition(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    PowerPoint.SlideShowTransition transition = slide.SlideShowTransition;

    string effectKey = input.GetProperty("effect").GetString();
    PowerPoint.PpEntryEffect effectValue;
    if (!TransitionEffectMap.TryGetValue(effectKey, out effectValue))
        throw new ArgumentException("set_slide_transition: unknown effect '" + effectKey + "'. Valid: " + string.Join(", ", TransitionEffectMap.Keys) + ".");
    transition.EntryEffect = effectValue;

    if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number)
        transition.Duration = (float)durEl.GetDouble();
    if (input.TryGetProperty("advanceOnClick", out var clickEl))
        transition.AdvanceOnClick = clickEl.ValueKind == JsonValueKind.True ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
    if (input.TryGetProperty("advanceAfterSeconds", out var advEl) && advEl.ValueKind == JsonValueKind.Number)
    {
        transition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
        transition.AdvanceTime = (float)advEl.GetDouble();
    }

    return new ToolResult { Output = "Slide " + slideIndex + " transition set to '" + effectKey + "'.", Mutated = true, Summary = "set_slide_transition" };
}
```

`SlideShowTransition` is a **live** COM object (confirmed via reflection: get-only property returning a mutable reference, the same shape as `Cell.Shading` in Word's table-shading work) - no `Set...` call needed on the `Slide` itself, only property assignments on the object it returns.

- [x] **Step 3: Register in `Execute`'s switch.**

- [x] **Step 4: Schema**

```ts
{
  name: 'set_slide_transition',
  description:
    'Sets or removes a slide\'s entry transition. effect:"none" removes it. durationSeconds controls how long the transition animation itself takes. ' +
    'advanceOnClick (default true in PowerPoint) and advanceAfterSeconds (sets an automatic timed advance) are independent - both can be on at once.',
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number' },
      effect: {
        type: 'string',
        enum: ['none', 'cut', 'fade', 'dissolve', 'random', 'wipeLeft', 'wipeRight', 'wipeUp', 'wipeDown', 'pushLeft', 'pushRight', 'pushUp', 'pushDown', 'coverLeft', 'coverRight', 'coverUp', 'coverDown', 'uncoverLeft', 'uncoverRight', 'uncoverUp', 'uncoverDown', 'zoomIn', 'zoomOut', 'zoomCenter', 'circle', 'diamond', 'splitHorizontal', 'splitVertical', 'wheel', 'blindsHorizontal', 'blindsVertical', 'checkerboard'],
      },
      durationSeconds: { type: 'number' },
      advanceOnClick: { type: 'boolean' },
      advanceAfterSeconds: { type: 'number' },
    },
    required: ['slideIndex', 'effect'],
  },
},
```

**Verification:** builds; real-PowerPoint test sets a few different effects, confirms `durationSeconds`/`advanceAfterSeconds` visibly apply (Transitions ribbon tab shows the same values back), and confirms `effect:"none"` actually removes a previously-set transition.

---

### Task 3: `add_animation`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [x] **Step 1: Curated animation-effect and trigger maps** (all confirmed present in the real enums via reflection):

```csharp
private static readonly Dictionary<string, PowerPoint.MsoAnimEffect> AnimationEffectMap = new Dictionary<string, PowerPoint.MsoAnimEffect>
{
    ["appear"] = PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
    ["fade"] = PowerPoint.MsoAnimEffect.msoAnimEffectFade,
    ["fly"] = PowerPoint.MsoAnimEffect.msoAnimEffectFly,
    ["flashOnce"] = PowerPoint.MsoAnimEffect.msoAnimEffectFlashOnce,
    ["wipe"] = PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
    ["zoom"] = PowerPoint.MsoAnimEffect.msoAnimEffectZoom,
    ["dissolve"] = PowerPoint.MsoAnimEffect.msoAnimEffectDissolve,
    ["bounce"] = PowerPoint.MsoAnimEffect.msoAnimEffectBounce,
    ["spiral"] = PowerPoint.MsoAnimEffect.msoAnimEffectSpiral,
    ["swivel"] = PowerPoint.MsoAnimEffect.msoAnimEffectSwivel,
    ["wheel"] = PowerPoint.MsoAnimEffect.msoAnimEffectWheel,
    ["split"] = PowerPoint.MsoAnimEffect.msoAnimEffectSplit,
    ["box"] = PowerPoint.MsoAnimEffect.msoAnimEffectBox,
    ["circle"] = PowerPoint.MsoAnimEffect.msoAnimEffectCircle,
    ["diamond"] = PowerPoint.MsoAnimEffect.msoAnimEffectDiamond,
    ["plus"] = PowerPoint.MsoAnimEffect.msoAnimEffectPlus,
    ["checkerboard"] = PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard,
    ["randomBars"] = PowerPoint.MsoAnimEffect.msoAnimEffectRandomBars,
    ["growAndTurn"] = PowerPoint.MsoAnimEffect.msoAnimEffectGrowAndTurn,
    ["riseUp"] = PowerPoint.MsoAnimEffect.msoAnimEffectRiseUp,
};

// All 7 real values, not a curated subset - small enough to expose completely.
private static readonly Dictionary<string, PowerPoint.MsoAnimTriggerType> AnimationTriggerMap = new Dictionary<string, PowerPoint.MsoAnimTriggerType>
{
    ["onClick"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick,
    ["withPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
    ["afterPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious,
};
```

- [x] **Step 2: Handler**

```csharp
private static ToolResult AddAnimation(JsonElement input)
{
    PowerPoint.Shape shape = ResolveShape(input);
    PowerPoint.Slide slide = shape.Parent as PowerPoint.Slide;
    PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;

    string effectKey = input.GetProperty("effect").GetString();
    PowerPoint.MsoAnimEffect effectValue;
    if (!AnimationEffectMap.TryGetValue(effectKey, out effectValue))
        throw new ArgumentException("add_animation: unknown effect '" + effectKey + "'. Valid: " + string.Join(", ", AnimationEffectMap.Keys) + ".");

    string triggerKey = input.TryGetProperty("trigger", out var trEl) && trEl.ValueKind == JsonValueKind.String ? trEl.GetString() : "onClick";
    PowerPoint.MsoAnimTriggerType triggerValue;
    if (!AnimationTriggerMap.TryGetValue(triggerKey, out triggerValue))
        throw new ArgumentException("add_animation: unknown trigger '" + triggerKey + "'. Valid: " + string.Join(", ", AnimationTriggerMap.Keys) + ".");

    // UNVERIFIED (see plan's own risk note): -1 is documented (not
    // independently reflection-confirmed) to mean "append at the end of
    // the sequence." If this throws or inserts at the wrong position,
    // the fallback is sequence.Count + 1 (a real, valid 1-based end
    // position) instead of -1.
    PowerPoint.Effect effect = sequence.AddEffect(shape, effectValue, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, triggerValue, -1);

    if (input.TryGetProperty("exit", out var exitEl) && exitEl.ValueKind == JsonValueKind.True)
        effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
    if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number)
        effect.Timing.Duration = (float)durEl.GetDouble();
    if (input.TryGetProperty("delaySeconds", out var delEl) && delEl.ValueKind == JsonValueKind.Number)
        effect.Timing.TriggerDelayTime = (float)delEl.GetDouble();

    int newIndex = sequence.Count - 1; // AddEffect(-1) appends; stable immediately after the call, same "report the new index" precedent as PP-9/PP-23
    return new ToolResult { Output = "Animation added at animationIndex " + newIndex + " ('" + effectKey + "', " + (input.TryGetProperty("exit", out var e2) && e2.ValueKind == JsonValueKind.True ? "exit" : "entrance") + ").", Mutated = true, Summary = "add_animation" };
}
```

`shape.Parent as PowerPoint.Slide` — `Shape.Parent` is documented to return the shape's containing slide for a normal (non-grouped) shape; this is the standard, established way to get from a `Shape` back to its `Slide` without requiring the caller to separately pass `slideIndex` (the schema still takes `slideIndex`/`shapeIndex` via `ResolveShape`, consistent with every other shape-targeting tool in this file - this cast is purely internal plumbing).

- [x] **Step 3: Register in `Execute`'s switch.**

- [x] **Step 4: Schema**

```ts
{
  name: 'add_animation',
  description:
    'Adds an entrance (default) or exit animation to a shape. effect: appear, fade, fly, flashOnce, wipe, zoom, dissolve, bounce, spiral, swivel, wheel, split, box, circle, diamond, plus, checkerboard, randomBars, growAndTurn, riseUp. ' +
    'trigger (default "onClick"): "onClick" starts on its own click during the slideshow, "withPrevious" starts together with the animation before it, "afterPrevious" starts automatically once the animation before it finishes. ' +
    'Does not support directional variants (e.g. "wipe from the left") - only the base effect. Returns the new animationIndex (0-based) for read_animations/edit_animation.',
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number' },
      shapeIndex: { type: 'number' },
      effect: { type: 'string', enum: ['appear', 'fade', 'fly', 'flashOnce', 'wipe', 'zoom', 'dissolve', 'bounce', 'spiral', 'swivel', 'wheel', 'split', 'box', 'circle', 'diamond', 'plus', 'checkerboard', 'randomBars', 'growAndTurn', 'riseUp'] },
      exit: { type: 'boolean' },
      trigger: { type: 'string', enum: ['onClick', 'withPrevious', 'afterPrevious'] },
      durationSeconds: { type: 'number' },
      delaySeconds: { type: 'number' },
    },
    required: ['slideIndex', 'shapeIndex', 'effect'],
  },
},
```

**Verification:** builds; real-PowerPoint test adds one entrance animation and confirms it plays in Slide Show view and appears in the Animation Pane with the right effect/trigger; **test the `-1` append-index behavior specifically first** (create 2-3 animations on the same slide in sequence, confirm each new one lands after the previous rather than at the start or erroring) since this is the plan's one unverified-at-write-time behavior.

---

### Task 4: `read_animations` (+ `read_slide` extension)

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [x] **Step 1: Handler**

```csharp
private static ToolResult ReadAnimations(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
    int count = sequence.Count;
    if (count == 0)
        return new ToolResult { Output = "No animations on slide " + slideIndex + ".", Summary = "read_animations" };

    var sb = new System.Text.StringBuilder();
    sb.AppendLine("Slide " + slideIndex + " has " + count + " animation(s):");
    for (int i = 1; i <= count; i++)
    {
        PowerPoint.Effect effect = sequence.Item(i);
        string effectName = null;
        foreach (var kv in AnimationEffectMap) { if (kv.Value == effect.EffectType) { effectName = kv.Key; break; } }
        string shapeName;
        try { shapeName = effect.Shape.Name; } catch { shapeName = "(unknown shape)"; }
        string triggerName = null;
        foreach (var kv in AnimationTriggerMap) { if (kv.Value == effect.Timing.TriggerType) { triggerName = kv.Key; break; } }
        sb.AppendLine("[" + (i - 1) + "] shape=\"" + shapeName + "\" effect=" + (effectName ?? ("unrecognized (" + effect.EffectType + ")")) +
            " kind=" + (effect.Exit == Microsoft.Office.Core.MsoTriState.msoTrue ? "exit" : "entrance") +
            " trigger=" + (triggerName ?? effect.Timing.TriggerType.ToString()) +
            " duration=" + effect.Timing.Duration + "s delay=" + effect.Timing.TriggerDelayTime + "s");
    }
    return new ToolResult { Output = sb.ToString().TrimEnd(), Summary = "read_animations" };
}
```

- [x] **Step 2: Extend `read_slide`'s existing per-shape/per-slide output** to mention the slide's current layout and transition, and whether it has any animations - cheap, and directly closes the same kind of "the model can add something but then can't see what it created" gap PP-23's `read_chart`/`read_table`/`read_smartart` all exist to fix:

```csharp
// Inside ReadSlide, near its existing header line:
string layoutName = null;
foreach (var kv in SlideLayoutMap) { if (kv.Value == slide.Layout) { layoutName = kv.Key; break; } }
sb.AppendLine("Layout: " + (layoutName ?? (slide.Layout == PowerPoint.PpSlideLayout.ppLayoutCustom ? "custom ('" + slide.CustomLayout.Name + "')" : slide.Layout.ToString())));
string transitionName = null;
foreach (var kv in TransitionEffectMap) { if (kv.Value == slide.SlideShowTransition.EntryEffect) { transitionName = kv.Key; break; } }
sb.AppendLine("Transition: " + (transitionName ?? (slide.SlideShowTransition.EntryEffect == PowerPoint.PpEntryEffect.ppEffectNone ? "none" : slide.SlideShowTransition.EntryEffect.ToString())));
int animCount = slide.TimeLine.MainSequence.Count;
if (animCount > 0) sb.AppendLine(animCount + " animation(s) - call read_animations to see them.");
```

- [x] **Step 3: Register `read_animations` in `Execute`'s switch and `AlwaysAllowedTools`.**

- [x] **Step 4: Schema**

```ts
{
  name: 'read_animations',
  description: 'Reads a slide\'s animations in play order, one per line (shape, effect, entrance/exit, trigger, timing). Needed before edit_animation, since animationIndex addresses this same order.',
  inputSchema: { type: 'object', properties: { slideIndex: { type: 'number' } }, required: ['slideIndex'] },
},
```

Add `read_animations` to `READER_TOOLS` (not `MUTATION_TOOLS`).

**Verification:** builds; a slide with 2+ animations of different effects/triggers reads back correctly; `read_slide` on a slide with a non-default layout/transition/animation shows all three without a separate call.

---

### Task 5: `edit_animation`

**Files:**
- Modify: `PowerPointAiAddIn/PowerPointTools.cs`, `PowerPointAiAddIn/web-src/entry.ts`

- [x] **Step 1: Handler**

```csharp
private static ToolResult EditAnimation(JsonElement input)
{
    int slideIndex = input.GetProperty("slideIndex").GetInt32();
    int animationIndex = input.GetProperty("animationIndex").GetInt32();
    PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
    PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
    if (animationIndex < 0 || animationIndex >= sequence.Count)
        throw new ArgumentOutOfRangeException("animationIndex", "animationIndex must be between 0 and " + (sequence.Count - 1) + " (" + sequence.Count + " animation(s) on this slide).");
    PowerPoint.Effect effect = sequence.Item(animationIndex + 1);

    string kind = input.GetProperty("kind").GetString();
    switch (kind)
    {
        case "delete":
            effect.Delete();
            return new ToolResult { Output = "Animation " + animationIndex + " deleted. Later animation indices have shifted - re-read (read_animations) before another edit in the same run.", Mutated = true, Summary = "edit_animation" };
        case "set_timing":
        {
            bool changed = false;
            if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number) { effect.Timing.Duration = (float)durEl.GetDouble(); changed = true; }
            if (input.TryGetProperty("delaySeconds", out var delEl) && delEl.ValueKind == JsonValueKind.Number) { effect.Timing.TriggerDelayTime = (float)delEl.GetDouble(); changed = true; }
            if (input.TryGetProperty("trigger", out var trEl) && trEl.ValueKind == JsonValueKind.String)
            {
                PowerPoint.MsoAnimTriggerType triggerValue;
                if (!AnimationTriggerMap.TryGetValue(trEl.GetString(), out triggerValue))
                    throw new ArgumentException("edit_animation: unknown trigger '" + trEl.GetString() + "'. Valid: " + string.Join(", ", AnimationTriggerMap.Keys) + ".");
                effect.Timing.TriggerType = triggerValue;
                changed = true;
            }
            if (!changed)
                throw new ArgumentException("edit_animation: set_timing requires at least one of durationSeconds, delaySeconds, trigger.");
            return new ToolResult { Output = "Animation " + animationIndex + " timing updated.", Mutated = true, Summary = "edit_animation" };
        }
        case "reorder":
        {
            int toIndex = input.GetProperty("toIndex").GetInt32();
            if (toIndex < 0 || toIndex >= sequence.Count)
                throw new ArgumentOutOfRangeException("toIndex", "toIndex must be between 0 and " + (sequence.Count - 1) + ".");
            effect.MoveTo(toIndex + 1);
            return new ToolResult { Output = "Animation moved from " + animationIndex + " to " + toIndex + ". Indices have shifted - re-read (read_animations) before another edit in the same run.", Mutated = true, Summary = "edit_animation" };
        }
        default:
            throw new ArgumentException("edit_animation: unknown kind '" + kind + "'. Valid: delete, set_timing, reorder.");
    }
}
```

`Effect.MoveTo(int)` was confirmed to exist on `Effect` via reflection (alongside `MoveBefore`/`MoveAfter`); its exact 0-based-vs-1-based indexing convention was **not** independently confirmed (only that the method exists) - flagged for the manual matrix, with `MoveBefore`/`MoveAfter` (relative-position moves against another `Effect` reference) as the fallback if `MoveTo`'s absolute-index semantics turn out to differ from this guess.

- [x] **Step 2: Register in `Execute`'s switch.**

- [x] **Step 3: Schema**

```ts
{
  name: 'edit_animation',
  description:
    'Edits an existing animation. kind: "delete", "set_timing" (durationSeconds?,delaySeconds?,trigger?), "reorder" (toIndex - 0-based new position in the play sequence). ' +
    'animationIndex addresses the animation (0-based, current play order - call read_animations first). delete/reorder shift later indices - re-read before another edit in the same run.',
  inputSchema: {
    type: 'object',
    properties: {
      slideIndex: { type: 'number' },
      animationIndex: { type: 'number' },
      kind: { type: 'string', enum: ['delete', 'set_timing', 'reorder'] },
      durationSeconds: { type: 'number' },
      delaySeconds: { type: 'number' },
      trigger: { type: 'string', enum: ['onClick', 'withPrevious', 'afterPrevious'] },
      toIndex: { type: 'number' },
    },
    required: ['slideIndex', 'animationIndex', 'kind'],
  },
},
```

**Verification:** builds; real-PowerPoint test deletes one animation (confirm the Animation Pane updates and later indices shift down), changes timing on another, and reorders two animations relative to each other - **verify `MoveTo`'s indexing convention specifically** per the risk note above.

---

### Task 6: Integration — `toolDisplay`, `AlwaysAllowedTools`/`READER_TOOLS`, system prompt

**Files:**
- Modify: `PowerPointAiAddIn/web-src/entry.ts`, `docs/ai-tool-surface.md`

- [x] **Step 1:** Add all 5 new tools' `toolDisplay` entries (English + Hebrew), matching this file's existing style (short label, one-sentence non-technical description).
- [x] **Step 2:** Confirm `read_animations` is in `READER_TOOLS` (Task 4) and the other 4 are in `MUTATION_TOOLS`.
- [x] **Step 3:** Update the PowerPoint skill's `systemPrompt` to mention slide layout/transition/animation control, and the "delete/reorder shift later animation indices - re-read before a second edit" caveat, matching the equivalent caveats already added for Word's table/SmartArt tools (PP-23) and PowerPoint's own `delete_slide` (PP-19).
- [x] **Step 4:** PowerPoint's tool count goes from 26 to 31. Update `docs/ai-tool-surface.md`'s PowerPoint section and `docs/superpowers/plans/STATUS.md`'s eventual entry for this plan with the new count.

**Verification:** `npx tsc --noEmit` clean.

---

### Task 7: Manual verification matrix

- [ ] `set_slide_layout {kind:'classic', layout:'blank'}` → slide becomes blank; `{layout:'titleOnly'}`, `{layout:'twoObjects'}` → each visibly correct.
- [ ] `set_slide_layout {kind:'custom', layoutName:'<a real theme layout name>'}` → applies that specific theme layout; an unmatched name → error listing the real available names.
- [ ] `set_slide_transition {effect:'fade', durationSeconds:1.5}` then `{effect:'wipeLeft'}` then `{effect:'none'}` → each visibly correct in the Transitions ribbon tab; `none` actually removes it.
- [ ] `set_slide_transition {effect:'fade', advanceAfterSeconds:3}` → confirm the slide auto-advances after 3s in Slide Show view, without disabling click-to-advance.
- [ ] `add_animation {effect:'fade', trigger:'onClick'}` on one shape, then a second `add_animation {effect:'fly', trigger:'afterPrevious'}` on a different shape - **confirm the second lands AFTER the first in the Animation Pane, not before or erroring** (the plan's one unverified `-1`-append assumption).
- [ ] `read_animations` after the above → both animations listed in the correct order with correct effect/trigger names.
- [ ] `edit_animation {kind:'set_timing', durationSeconds:2, trigger:'withPrevious'}` on the first animation → visibly updates in the Animation Pane.
- [ ] `edit_animation {kind:'reorder', toIndex:1}` swapping the two animations' order → **confirm `MoveTo`'s indexing does what's expected** (the plan's second unverified assumption); `read_animations` after confirms the new order.
- [ ] `edit_animation {kind:'delete'}` on one animation → removed; `read_animations` shows only the remaining one at index 0.
- [ ] Unknown `kind`/`effect`/`trigger`/`layout` on every new tool → specific error listing valid values; nothing changed.
- [ ] Out-of-range `animationIndex`/`toIndex` → specific error naming the valid range; nothing changed.
- [ ] `read_slide` on a slide with a non-default layout, a transition, and 2+ animations → all three show up in its output without needing `read_animations` separately (confirms Task 4 Step 2's extension).
- [ ] Natural language end-to-end: "make this slide blank", "add a fade transition to slide 2", "make the title fade in on click", "remove that animation" - each resolves to a correct tool call on the first attempt.
