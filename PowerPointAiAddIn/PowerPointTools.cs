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
        // Editing-mode gating, per-document since PP-1: the client filters the tool list per
        // mode, but Execute() independently re-checks here as defense-in-depth. See
        // PowerPointTools.cs.md for the full rationale.
        private static readonly Dictionary<string, EditingMode> ModeByDoc = new Dictionary<string, EditingMode>();

        public static void SetMode(string docKey, EditingMode mode)
        {
            ModeByDoc[docKey] = mode;
        }

        private static EditingMode ModeFor(string docKey)
        {
            EditingMode m;
            return ModeByDoc.TryGetValue(docKey, out m) ? m : EditingMode.FullAutonomy;
        }

        // Tools always allowed regardless of editing mode (read-only, no document mutation).
        private static readonly System.Collections.Generic.HashSet<string> AlwaysAllowedTools =
            new System.Collections.Generic.HashSet<string> { "get_deck_context", "read_slide", "read_group", "read_animations", "find_text", "read_smartart", "read_master_elements", "list_layouts" };

        public static ToolResult Execute(string docKey, string name, JsonElement input)
        {
            try
            {
                EditingMode mode = ModeFor(docKey);
                bool isAlwaysAllowed = AlwaysAllowedTools.Contains(name);
                if (!isAlwaysAllowed && !IsMutationAllowed(mode))
                {
                    return new ToolResult
                    {
                        Output = "Blocked: editing mode is " + ModeLabel(mode) + ".",
                        IsError = true,
                        Summary = name,
                    };
                }

                // PowerPoint can coalesce several back-to-back COM mutations into one
                // undo entry with no UI tick between them (confirmed live, 2026-10-02 -
                // see PowerPointTools.cs.md). StartNewUndoEntry() forces a fresh entry
                // before each mutating call so one tool call = one undo step; unlike
                // Word's equivalent fix, undo_last_action/redo_last_action are NOT
                // excluded here - live-tested harmless against an empty redo stack.
                if (!isAlwaysAllowed)
                {
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                }

                switch (name)
                {
                    case "get_deck_context": return GetDeckContext();
                    case "read_slide": return ReadSlide(input);
                    case "read_group": return ReadGroup(input);
                    case "find_text": return FindTextPpt(input);
                    case "replace_text": return ReplaceTextPpt(input);
                    case "set_element_text": return SetElementText(input);
                    case "set_slide_notes": return SetSlideNotes(input);
                    case "set_element_style": return SetElementStyle(input);
                    case "set_element_transform": return SetElementTransform(input);
                    case "set_element_order": return SetElementOrder(input);
                    case "add_text_box": return AddTextBox(input);
                    case "add_shape": return AddShape(input);
                    case "delete_element": return DeleteElement(input);
                    case "add_slide": return AddSlide(input);
                    case "set_element_fill": return SetElementFill(input);
                    case "set_element_stroke": return SetElementStroke(input);
                    case "set_slide_background": return SetSlideBackground(input);
                    case "ungroup_element": return UngroupElement(input);
                    case "group_element": return GroupElement(input);
                    case "add_table": return AddTable(input);
                    case "edit_table_cell": return EditTableCell(input);
                    case "edit_table_structure": return EditTableStructure(input);
                    case "edit_table_style": return EditTableStyle(input);
                    case "add_chart": return AddChartPpt(input);
                    case "edit_chart": return EditChartPpt(input);
                    case "add_smartart": return AddSmartArt(input);
                    case "edit_smartart": return EditSmartArt(input);
                    case "read_smartart": return ReadSmartArt(input);
                    case "crop_image": return CropImage(input);
                    case "replace_image": return ReplaceImagePpt(input);
                    case "set_picture_opacity": return SetPictureOpacity(input);
                    case "delete_slide": return DeleteSlide(input);
                    case "move_slide": return MoveSlide(input);
                    case "duplicate_slide": return DuplicateSlide(input);
                    case "set_slide_layout": return SetSlideLayout(input);
                    case "set_slide_transition": return SetSlideTransition(input);
                    case "add_animation": return AddAnimation(input);
                    case "read_animations": return ReadAnimations(input);
                    case "edit_animation": return EditAnimation(input);
                    case "duplicate_element": return DuplicateElement(input);
                    case "copy_element": return CopyElement(input);
                    case "move_element": return MoveElement(input);
                    case "copy_element_style": return CopyElementStyle(input);
                    case "set_headers_footers": return SetHeadersFooters(input);
                    case "add_master_element": return AddMasterElement(input);
                    case "set_master_element_transform": return SetMasterElementTransform(input);
                    case "remove_master_element": return RemoveMasterElement(input);
                    case "read_master_elements": return ReadMasterElements(input);
                    case "list_layouts": return ListLayouts(input);
                    case "undo_last_action": return UndoLastAction();
                    case "redo_last_action": return RedoLastAction();
                    default: return new ToolResult { Output = "Unknown tool: " + name, IsError = true, Summary = name };
                }
            }
            catch (Exception ex)
            {
                return new ToolResult { Output = ex.Message, IsError = true, Summary = name };
            }
        }

        // Read Only and Comment Only both block all mutating tools (no comment-equivalent
        // tool exists yet, so the two currently behave identically). Track Changes is
        // simple allow/block gating, not a native revision-tracking UI. See .md for more.
        private static bool IsMutationAllowed(EditingMode mode)
        {
            return mode == EditingMode.TrackChanges || mode == EditingMode.FullAutonomy;
        }

        private static string ModeLabel(EditingMode mode)
        {
            switch (mode)
            {
                case EditingMode.ReadOnly: return "Read Only";
                case EditingMode.CommentOnly: return "Comment Only";
                case EditingMode.TrackChanges: return "Track Changes";
                default: return "Full Autonomy";
            }
        }

        // Known limitation (PP-1 Task 5 Step 5): resolves whichever presentation is
        // ACTIVE right now, not necessarily the one whose pane initiated this call.
        // See WordTools.cs's ActiveDoc for the identical rationale and out-of-scope decision.
        private static PowerPoint.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;

    }
}

