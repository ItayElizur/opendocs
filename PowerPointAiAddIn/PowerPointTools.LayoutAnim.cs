using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeAi.Shared;

namespace PowerPointAiAddIn
{
    public static partial class PowerPointTools
    {
        // PP-24: curated subset of the 37 PpSlideLayout values (confirmed via
        // reflection against the real referenced PIA, not recalled) - the
        // pre-2007 leftovers (ppLayoutOrgchart, ppLayoutMediaClipAndText,
        // etc.) are omitted as unlikely to be what a model means by a
        // layout request. Same curated-map-with-throw-on-unknown pattern as
        // ChartTypes.ByName/AlignmentMap elsewhere in this file.
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

        // Mirrors WordTools.cs's ResolveSmartArtGalleryItem (PP-23): resolves
        // by case-insensitive substring match against this deck's own live
        // theme layouts, since custom layout names are not a fixed enum -
        // a miss lists the real available names so the caller can retry
        // correctly instead of guessing blind.
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

        // PP-24: curated subset of PpEntryEffect's 189 values, all confirmed
        // present via reflection against the real referenced PIA.
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

        // PP-24: curated subset of MsoAnimEffect's 151 values, all confirmed
        // present via reflection against the real referenced PIA.
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

        // All 7 real MsoAnimTriggerType values are not exposed - msoAnimTriggerNone
        // and msoAnimTriggerOnMediaBookmark/msoAnimTriggerMixed are not
        // meaningful choices for add_animation/edit_animation's caller.
        private static readonly Dictionary<string, PowerPoint.MsoAnimTriggerType> AnimationTriggerMap = new Dictionary<string, PowerPoint.MsoAnimTriggerType>
        {
            ["onClick"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick,
            ["withPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious,
            ["afterPrevious"] = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious,
        };

        private static ToolResult AddAnimation(JsonElement input)
        {
            PowerPoint.Shape shape = ResolveTopLevelShape(input, "add_animation");
            PowerPoint.Slide slide = (PowerPoint.Slide)shape.Parent;
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;

            string effectKey = input.GetProperty("effect").GetString();
            PowerPoint.MsoAnimEffect effectValue;
            if (!AnimationEffectMap.TryGetValue(effectKey, out effectValue))
                throw new ArgumentException("add_animation: unknown effect '" + effectKey + "'. Valid: " + string.Join(", ", AnimationEffectMap.Keys) + ".");

            string triggerKey = input.TryGetProperty("trigger", out var trEl) && trEl.ValueKind == JsonValueKind.String ? trEl.GetString() : "onClick";
            PowerPoint.MsoAnimTriggerType triggerValue;
            if (!AnimationTriggerMap.TryGetValue(triggerKey, out triggerValue))
                throw new ArgumentException("add_animation: unknown trigger '" + triggerKey + "'. Valid: " + string.Join(", ", AnimationTriggerMap.Keys) + ".");

            bool isExit = input.TryGetProperty("exit", out var exitEl) && exitEl.ValueKind == JsonValueKind.True;

            // UNVERIFIED (plan's own risk note, PP-24): -1 is documented
            // (not independently reflection-confirmed - reflection gives
            // signatures, not runtime semantics) to mean "append at the end
            // of the sequence." Test this specifically before trusting
            // multi-animation ordering.
            PowerPoint.Effect effect = sequence.AddEffect(shape, effectValue, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, triggerValue, -1);

            if (isExit)
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            if (input.TryGetProperty("durationSeconds", out var durEl) && durEl.ValueKind == JsonValueKind.Number)
                effect.Timing.Duration = (float)durEl.GetDouble();
            if (input.TryGetProperty("delaySeconds", out var delEl) && delEl.ValueKind == JsonValueKind.Number)
                effect.Timing.TriggerDelayTime = (float)delEl.GetDouble();

            int newIndex = sequence.Count - 1; // AddEffect(-1) appends; stable immediately after the call
            return new ToolResult
            {
                Output = "Animation added at animationIndex " + newIndex + " ('" + effectKey + "', " + (isExit ? "exit" : "entrance") + ").",
                Mutated = true,
                Summary = "add_animation",
            };
        }

        private static ToolResult ReadAnimations(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            int count = sequence.Count;
            if (count == 0)
                return new ToolResult { Output = "No animations on slide " + slideIndex + ".", Summary = "read_animations" };

            var sb = new StringBuilder();
            sb.AppendLine("Slide " + slideIndex + " has " + count + " animation(s):");
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Effect effect = sequence[i];
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

        private static ToolResult EditAnimation(JsonElement input)
        {
            int slideIndex = input.GetProperty("slideIndex").GetInt32();
            int animationIndex = input.GetProperty("animationIndex").GetInt32();
            PowerPoint.Slide slide = ActivePresentation.Slides[slideIndex + 1];
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            if (animationIndex < 0 || animationIndex >= sequence.Count)
                throw new ArgumentOutOfRangeException("animationIndex", "animationIndex must be between 0 and " + (sequence.Count - 1) + " (" + sequence.Count + " animation(s) on this slide).");
            PowerPoint.Effect effect = sequence[animationIndex + 1];

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
                    // UNVERIFIED (plan's own risk note, PP-24): MoveTo's exact
                    // 0-based-vs-1-based indexing convention was confirmed to
                    // EXIST via reflection but not independently confirmed
                    // for its runtime semantics. If this lands one position
                    // off, MoveBefore/MoveAfter (relative-position moves
                    // against another Effect reference) are the fallback.
                    effect.MoveTo(toIndex + 1);
                    return new ToolResult { Output = "Animation moved from " + animationIndex + " to " + toIndex + ". Indices have shifted - re-read (read_animations) before another edit in the same run.", Mutated = true, Summary = "edit_animation" };
                }
                default:
                    throw new ArgumentException("edit_animation: unknown kind '" + kind + "'. Valid: delete, set_timing, reorder.");
            }
        }

    }
}

