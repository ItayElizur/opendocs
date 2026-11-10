using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // Write side of the non-body objects read_object lists: one tool per family, each with a small action
        // enum over uniform params (a single generic add_object/edit_object was rejected - the per-kind payloads
        // share almost nothing). Indices are read_object's per-kind indices. See WordTools.ObjectEdits.cs.md.

        private static string OptStr(JsonElement i, string n)
        {
            return i.TryGetProperty(n, out var e) && e.ValueKind == JsonValueKind.String ? e.GetString() : null;
        }

        private static int? OptInt(JsonElement i, string n)
        {
            return i.TryGetProperty(n, out var e) && e.ValueKind == JsonValueKind.Number ? (int?)e.GetInt32() : null;
        }

        private static float? OptFloat(JsonElement i, string n)
        {
            return i.TryGetProperty(n, out var e) && e.ValueKind == JsonValueKind.Number ? (float?)e.GetDouble() : null;
        }

        private static bool? OptBool(JsonElement i, string n)
        {
            if (!i.TryGetProperty(n, out var e)) return null;
            if (e.ValueKind == JsonValueKind.True) return true;
            if (e.ValueKind == JsonValueKind.False) return false;
            return null;
        }

        private static string RequireAction(JsonElement input, string tool, params string[] valid)
        {
            string action = OptStr(input, "action");
            if (action == null || Array.IndexOf(valid, action) < 0)
                throw new ArgumentException(tool + ": action is required. Valid: " + string.Join(", ", valid) + ".");
            return action;
        }

        private static int RequireIndex(JsonElement input, string tool, string what)
        {
            int? index = OptInt(input, "index");
            if (!index.HasValue)
                throw new ArgumentException(tool + ": index is required to " + what + " (use read_object to list the indices).");
            return index.Value;
        }

        // End of a paragraph's TEXT (before its paragraph mark): where a footnote reference or content control
        // belongs. RangeAfterBlock collapses past the mark, i.e. into the NEXT paragraph, which is wrong for these.
        private static Word.Range RangeAtEndOfBlockText(int blockIndex)
        {
            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            if (blockIndex < 0 || blockIndex > paragraphs.Count - 1)
                throw new ArgumentOutOfRangeException("afterBlockIndex", "afterBlockIndex must be between 0 and " + (paragraphs.Count - 1) + ".");
            Word.Range r = paragraphs[blockIndex + 1].Range;
            r.MoveEnd(Word.WdUnits.wdCharacter, -1);
            r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            return r;
        }

        // doc.ContentControls.Add(type, range) failed while the user's cursor was inside a content control
        // ("the selection partially covers a plain text content control") - Word consulted the selection, not the
        // range. Bind the add to the range itself; if Word still refuses, move the selection to the range for the
        // call and put it back afterwards. See WordTools.ObjectEdits.cs.md.
        private static dynamic AddContentControl(Word.Range at, int code)
        {
            try
            {
                return ((dynamic)at).ContentControls.Add(code);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                int start = sel.Start, end = sel.End;
                try
                {
                    at.Select();
                    return ((dynamic)sel.Range).ContentControls.Add(code);
                }
                finally
                {
                    sel.SetRange(start, end);
                }
            }
        }

        // The last paragraph's text end - the default insertion point when no afterBlockIndex is given.
        private static Word.Range RangeAtEndOfDocumentText()
        {
            return RangeAtEndOfBlockText(ActiveDoc.Paragraphs.Count - 1);
        }

        private static string Paras(string text)
        {
            return (text ?? "").Replace("\r\n", "\r").Replace('\n', '\r');
        }

        // ---------------------------------------------------------------- content controls

        private static readonly Dictionary<string, int> ContentControlTypeCodes = new Dictionary<string, int>
        {
            { "rich_text", 0 }, { "plain_text", 1 }, { "combo_box", 3 }, { "dropdown", 4 }, { "date", 6 }, { "checkbox", 8 },
        };

        private static dynamic NthContentControl(Word.Document doc, int index)
        {
            int count = doc.ContentControls.Count;
            if (index < 0 || index >= count)
                throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (count - 1) + " (" + count + " content control(s) in the document).");
            int i = 0;
            foreach (dynamic cc in doc.ContentControls)
            {
                if (i++ == index) return cc;
            }
            throw new ArgumentOutOfRangeException("index");
        }

        private static int ContentControlIndexOf(Word.Document doc, dynamic target)
        {
            string id = (string)target.ID;
            int i = 0;
            foreach (dynamic cc in doc.ContentControls)
            {
                if ((string)cc.ID == id) return i;
                i++;
            }
            return -1;
        }

        // Rejects fields the control type cannot hold before anything is created or written, so a bad call leaves the
        // document untouched (a control deleted afterwards leaves its text behind as a tracked deletion).
        private static void CheckContentControlFields(int typeCode, JsonElement input)
        {
            if (OptBool(input, "checked").HasValue && typeCode != 8)
                throw new ArgumentException("edit_content_control: checked only applies to a checkbox content control.");
            if (input.TryGetProperty("entries", out var entries) && entries.ValueKind == JsonValueKind.Array && typeCode != 3 && typeCode != 4)
                throw new ArgumentException("edit_content_control: entries only applies to a dropdown or combo_box content control.");
        }

        private static void ApplyContentControlFields(dynamic cc, JsonElement input)
        {
            string title = OptStr(input, "title");
            string tag = OptStr(input, "tag");
            string text = OptStr(input, "text");
            bool? isChecked = OptBool(input, "checked");
            if (title != null) cc.Title = title;
            if (tag != null) cc.Tag = tag;
            if (text != null)
            {
                cc.Range.Text = text;
                ApplyAutoDirection((Word.Range)cc.Range);
            }
            if (isChecked.HasValue) cc.Checked = isChecked.Value;
            if (input.TryGetProperty("entries", out var entries) && entries.ValueKind == JsonValueKind.Array)
            {
                cc.DropdownListEntries.Clear();
                foreach (JsonElement e in entries.EnumerateArray())
                {
                    if (e.ValueKind == JsonValueKind.String) cc.DropdownListEntries.Add(e.GetString(), e.GetString());
                }
            }
        }

        private static ToolResult EditContentControl(JsonElement input)
        {
            string action = RequireAction(input, "edit_content_control", "add", "edit", "delete");
            Word.Document doc = ActiveDoc;

            if (action == "add")
            {
                string type = OptStr(input, "type");
                int code;
                if (type == null || !ContentControlTypeCodes.TryGetValue(type, out code))
                    throw new ArgumentException("edit_content_control: type is required for add. Valid: " + string.Join(", ", ContentControlTypeCodes.Keys) + ".");
                int? after = OptInt(input, "afterBlockIndex");
                CheckContentControlFields(code, input);
                Word.Range at = after.HasValue ? RangeAtEndOfBlockText(after.Value) : RangeAtEndOfDocumentText();
                dynamic cc = AddContentControl(at, code);
                try
                {
                    ApplyContentControlFields(cc, input);
                }
                catch
                {
                    // Otherwise the control exists, the call reports an error, and a retry adds a duplicate. Tracking is off for
                    // the delete so the control's text is removed instead of left behind as a tracked deletion.
                    bool tracking = doc.TrackRevisions;
                    try
                    {
                        doc.TrackRevisions = false;
                        cc.Delete(true);
                    }
                    catch { }
                    finally { doc.TrackRevisions = tracking; }
                    throw;
                }
                return new ToolResult
                {
                    Output = "Added " + type + " content control (index " + ContentControlIndexOf(doc, cc) + "). Indices of later content controls may have shifted.",
                    Mutated = true, Summary = "edit_content_control add",
                };
            }

            int index = RequireIndex(input, "edit_content_control", action);
            dynamic target = NthContentControl(doc, index);
            if (action == "edit")
            {
                CheckContentControlFields((int)target.Type, input);
                ApplyContentControlFields(target, input);
                return new ToolResult { Output = "Updated content control " + index + ".", Mutated = true, Summary = "edit_content_control edit" };
            }
            // delete: keep the content as ordinary text unless deleteContents is set.
            bool deleteContents = OptBool(input, "deleteContents") ?? false;
            target.Delete(deleteContents);
            return new ToolResult
            {
                Output = "Deleted content control " + index + (deleteContents ? " and its contents" : " (its content was kept as ordinary text)") + ". Later content control indices have shifted.",
                Mutated = true, Summary = "edit_content_control delete",
            };
        }

        // ---------------------------------------------------------------- headers / footers

        private static ToolResult EditHeaderFooter(JsonElement input)
        {
            string action = RequireAction(input, "edit_header_footer", "set", "clear");
            string which = OptStr(input, "which");
            if (which != "header" && which != "footer")
                throw new ArgumentException("edit_header_footer: which is required. Valid: header, footer.");
            string type = OptStr(input, "type") ?? "primary";
            int typeCode;
            switch (type)
            {
                case "primary": typeCode = 1; break;
                case "first_page": typeCode = 2; break;
                case "even_pages": typeCode = 3; break;
                default: throw new ArgumentException("edit_header_footer: unknown type '" + type + "'. Valid: primary, first_page, even_pages.");
            }
            string text = OptStr(input, "text");
            if (action == "set" && text == null)
                throw new ArgumentException("edit_header_footer: text is required for set.");

            Word.Document doc = ActiveDoc;
            int sectionCount = doc.Sections.Count;
            int section = OptInt(input, "section") ?? 1;
            if (section < 1 || section > sectionCount)
                throw new ArgumentOutOfRangeException("section", "section must be between 1 and " + sectionCount + ".");

            dynamic sec = doc.Sections[section];
            // First-page / even-page parts are only shown when the section enables them, so turn that on when
            // writing one. OddAndEvenPagesHeaderFooter is effectively document-wide in Word.
            if (action == "set" && typeCode == 2) sec.PageSetup.DifferentFirstPageHeaderFooter = -1;
            if (action == "set" && typeCode == 3) sec.PageSetup.OddAndEvenPagesHeaderFooter = -1;

            dynamic hf = (which == "header" ? sec.Headers : sec.Footers).Item(typeCode);
            // A part linked to the previous section would write through to the previous section; break the link so
            // the edit stays in the requested section.
            bool unlinked = false;
            if (section > 1 && (bool)hf.LinkToPrevious) { hf.LinkToPrevious = false; unlinked = true; }
            int replacedFields = (int)hf.Range.Fields.Count;

            // The old content goes out first (a tracked deletion under Track Changes), then the new content is built left to
            // right before the part's final paragraph mark: text, a field for each {PAGE}-style token, more text, and so on.
            // Nothing is searched for afterwards, so deleted text that is still in the range under Track Changes can never be
            // matched a second time.
            hf.Range.Text = "";
            int insertedFields = 0;
            if (action == "set")
            {
                foreach (FieldSegment part in FieldTokens.Split(Paras(text)))
                {
                    Word.Range r = hf.Range;
                    int pos = r.End - 1;
                    r.SetRange(pos, pos);
                    if (part.Field == null)
                    {
                        r.InsertAfter(part.Text);
                    }
                    else
                    {
                        ((dynamic)r).Fields.Add(r, -1, part.Field, false);
                        insertedFields++;
                    }
                }
                ApplyAutoDirection((Word.Range)hf.Range);
            }

            return new ToolResult
            {
                Output = (action == "clear" ? "Cleared " : "Set ") + "section " + section + " " + which + " (" + type.Replace('_', ' ') + ")."
                    + (insertedFields > 0 ? " Inserted " + insertedFields + " field(s)." : "")
                    + (replacedFields > 0 ? " Replaced " + replacedFields + " field(s) that were in it." : "")
                    + (unlinked ? " It was linked to the previous section; the link is now broken so the edit stays in this section." : "")
                    + (action == "set" && typeCode == 3 ? " Even-page headers/footers are now switched on for the whole document." : "")
                    + (action == "set" && typeCode == 2 ? " A different first page is now switched on for this section." : ""),
                Mutated = true, Summary = "edit_header_footer " + action,
            };
        }

        // ---------------------------------------------------------------- footnotes / endnotes

        private static dynamic NotesOf(Word.Document doc, string noteKind)
        {
            return noteKind == "footnote" ? (dynamic)doc.Footnotes : doc.Endnotes;
        }

        private static ToolResult EditNote(JsonElement input)
        {
            string action = RequireAction(input, "edit_note", "add", "edit", "delete");
            string noteKind = OptStr(input, "noteKind");
            if (noteKind != "footnote" && noteKind != "endnote")
                throw new ArgumentException("edit_note: noteKind is required. Valid: footnote, endnote.");
            Word.Document doc = ActiveDoc;
            dynamic notes = NotesOf(doc, noteKind);

            if (action == "add")
            {
                string text = OptStr(input, "text");
                if (text == null) throw new ArgumentException("edit_note: text is required for add.");
                int? after = OptInt(input, "afterBlockIndex");
                if (!after.HasValue) throw new ArgumentException("edit_note: afterBlockIndex (the paragraph that carries the reference mark) is required for add.");

                Word.Range at;
                string afterText = OptStr(input, "afterText");
                if (afterText != null)
                {
                    // Reference mark goes right after afterText, which must match exactly once in that paragraph;
                    // several matches are an error that shows each one numbered, so the model picks a longer phrase or
                    // passes match: N (only meaningful after seeing that list).
                    // The search runs on the paragraph text with footnote marks, field marks etc. removed
                    // (TextAnchors), because those sit between words and break Word's own Find.
                    RangeAtEndOfBlockText(after.Value); // validates afterBlockIndex, same error as the branch below
                    Word.Range para = ActiveDoc.Paragraphs[after.Value + 1].Range;
                    int paraStart = para.Start, paraEnd = para.End;
                    para.TextRetrievalMode.IncludeFieldCodes = true;
                    para.TextRetrievalMode.IncludeHiddenText = true;
                    string rawText = para.Text ?? "";
                    if (!TextAnchors.PositionsAlign(rawText, paraEnd - paraStart))
                        throw new InvalidOperationException("edit_note: the text positions of paragraph " + after.Value + " cannot be mapped reliably; omit afterText to put the mark at the end of the paragraph.");
                    CleanText clean = TextAnchors.Clean(rawText);
                    List<int> starts = TextAnchors.FindAll(clean.Text, afterText, 20);
                    if (starts.Count == 0)
                        throw new ArgumentException("edit_note: afterText '" + afterText + "' was not found in paragraph " + after.Value + ".");
                    int? match = OptInt(input, "match");
                    if (match.HasValue && (match.Value < 1 || match.Value > starts.Count))
                        throw new ArgumentException("edit_note: match " + match.Value + " does not exist - afterText '" + afterText + "' matches " + starts.Count + " time(s) in paragraph " + after.Value + ".");
                    if (starts.Count > 1 && !match.HasValue)
                    {
                        var shown = new List<string>();
                        for (int i = 0; i < starts.Count && i < 5; i++)
                            shown.Add((i + 1) + ") \"" + TextAnchors.Snippet(clean.Text, starts[i], afterText.Length, 20) + "\"");
                        throw new ArgumentException("edit_note: afterText '" + afterText + "' matches " + starts.Count + " times in paragraph " + after.Value
                            + " (" + string.Join("; ", shown) + "). Each match is marked [[like this]]. Retry with a longer, unique phrase, or pass match: N for the numbered match you want.");
                    }
                    int markPos = paraStart + clean.Map[starts[(match ?? 1) - 1] + afterText.Length - 1] + 1;
                    at = ActiveDoc.Range(markPos, markPos);
                    at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    at = RangeAtEndOfBlockText(after.Value);
                }
                dynamic note = notes.Add(at, Type.Missing, text);
                ApplyAutoDirection((Word.Range)note.Range);
                return new ToolResult
                {
                    Output = "Added " + noteKind + " (index " + ((int)note.Index - 1) + "). Later " + noteKind + " indices may have shifted.",
                    Mutated = true, Summary = "edit_note add",
                };
            }

            int index = RequireIndex(input, "edit_note", action);
            int count = notes.Count;
            if (index < 0 || index >= count)
                throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (count - 1) + " (" + count + " " + noteKind + "(s) in the document).");
            dynamic target = notes.Item(index + 1);
            if (action == "edit")
            {
                string text = OptStr(input, "text");
                if (text == null) throw new ArgumentException("edit_note: text is required for edit.");
                target.Range.Text = text;
                ApplyAutoDirection((Word.Range)target.Range);
                return new ToolResult { Output = "Updated " + noteKind + " " + index + ".", Mutated = true, Summary = "edit_note edit" };
            }
            target.Delete();
            return new ToolResult { Output = "Deleted " + noteKind + " " + index + ". Later " + noteKind + " indices have shifted.", Mutated = true, Summary = "edit_note delete" };
        }

        // ---------------------------------------------------------------- table of contents

        private static ToolResult EditToc(JsonElement input)
        {
            string action = RequireAction(input, "edit_toc", "add", "edit", "update", "delete");
            Word.Document doc = ActiveDoc;
            int? upper = OptInt(input, "upperLevel");
            int? lower = OptInt(input, "lowerLevel");
            foreach (int? level in new[] { upper, lower })
            {
                if (level.HasValue && (level.Value < 1 || level.Value > 9))
                    throw new ArgumentOutOfRangeException("upperLevel/lowerLevel", "heading levels must be between 1 and 9.");
            }

            if (action == "add")
            {
                if ((upper ?? 1) > (lower ?? 3))
                    throw new ArgumentException("edit_toc: upperLevel (" + (upper ?? 1) + ") must not be greater than lowerLevel (" + (lower ?? 3) + ").");
                // Same convention as insert_content: afterBlockIndex -1 / omitted = start of the document.
                Word.Range at = RangeAfterBlock(OptInt(input, "afterBlockIndex") ?? -1);
                doc.TablesOfContents.Add(at, true, upper ?? 1, lower ?? 3);
                return new ToolResult
                {
                    Output = "Added a table of contents (headings " + (upper ?? 1) + "-" + (lower ?? 3) + "). Paragraph indices after this point have shifted - re-read before addressing them.",
                    Mutated = true, Summary = "edit_toc add",
                };
            }

            int index = RequireIndex(input, "edit_toc", action);
            int count = doc.TablesOfContents.Count;
            if (index < 0 || index >= count)
                throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (count - 1) + " (" + count + " TOC(s) in the document).");
            dynamic target = doc.TablesOfContents[index + 1];

            switch (action)
            {
                case "edit":
                    if (!upper.HasValue && !lower.HasValue)
                        throw new ArgumentException("edit_toc: pass upperLevel and/or lowerLevel for edit.");
                    int newUpper = upper ?? (int)target.UpperHeadingLevel;
                    int newLower = lower ?? (int)target.LowerHeadingLevel;
                    if (newUpper > newLower)
                        throw new ArgumentException("edit_toc: upperLevel (" + newUpper + ") must not be greater than lowerLevel (" + newLower + ").");
                    // Word refuses an upper level below the current lower one, so write the pair in an order that stays valid.
                    int curLower = (int)target.LowerHeadingLevel;
                    DebugLog.Write("edit_toc edit: levels " + (int)target.UpperHeadingLevel + "-" + curLower + " -> " + newUpper + "-" + newLower);
                    if (newUpper > curLower)
                    {
                        target.LowerHeadingLevel = newLower;
                        DebugLog.Write("edit_toc edit: lower written");
                        target.UpperHeadingLevel = newUpper;
                    }
                    else
                    {
                        target.UpperHeadingLevel = newUpper;
                        DebugLog.Write("edit_toc edit: upper written");
                        target.LowerHeadingLevel = newLower;
                    }
                    DebugLog.Write("edit_toc edit: both written");
                    target.Update();
                    return new ToolResult { Output = "Updated TOC " + index + " heading levels and refreshed it. Paragraph indices after it may have shifted.", Mutated = true, Summary = "edit_toc edit" };
                case "update":
                    if (OptBool(input, "pageNumbersOnly") ?? false) target.UpdatePageNumbers(); else target.Update();
                    return new ToolResult { Output = "Refreshed TOC " + index + ".", Mutated = true, Summary = "edit_toc update" };
                default:
                    target.Delete();
                    return new ToolResult { Output = "Deleted TOC " + index + ". Paragraph indices after it have shifted.", Mutated = true, Summary = "edit_toc delete" };
            }
        }

        // ---------------------------------------------------------------- text boxes

        private static ToolResult EditTextbox(JsonElement input)
        {
            string action = RequireAction(input, "edit_textbox", "add", "edit", "delete");
            Word.Document doc = ActiveDoc;
            string text = OptStr(input, "text");
            float? x = OptFloat(input, "x"), y = OptFloat(input, "y"), w = OptFloat(input, "w"), h = OptFloat(input, "h");

            if (action == "add")
            {
                int? after = OptInt(input, "afterBlockIndex");
                dynamic anchor = after.HasValue ? (dynamic)RangeAtEndOfBlockText(after.Value) : Type.Missing;
                dynamic shape = ((dynamic)doc).Shapes.AddTextbox(1 /* wdTextOrientationHorizontal */, x ?? 72f, y ?? 72f, w ?? 200f, h ?? 80f, anchor);
                try
                {
                    string name = OptStr(input, "name");
                    if (name != null) shape.Name = name;
                    if (text != null)
                    {
                        shape.TextFrame.TextRange.Text = text;
                        ApplyAutoDirection((Word.Range)shape.TextFrame.TextRange);
                    }
                    // AddTextbox positions are relative to the anchor paragraph; x/y are documented as points from the page corner.
                    shape.RelativeHorizontalPosition = 1; // wdRelativeHorizontalPositionPage
                    shape.RelativeVerticalPosition = 1;   // wdRelativeVerticalPositionPage
                    shape.Left = x ?? 72f;
                    shape.Top = y ?? 72f;
                }
                catch
                {
                    // Otherwise the box exists, the call reports an error, and a retry adds a duplicate.
                    try { shape.Delete(); } catch { }
                    throw;
                }

                string finalName = (string)shape.Name;
                int newId = (int)shape.ID; // names can repeat, the ID cannot
                int index = ListShapes(doc).FindIndex(si => { try { return (int)si.Shape.ID == newId; } catch { return false; } });
                return new ToolResult
                {
                    Output = "Added text box \"" + finalName + "\"" + (index >= 0 ? " (shape index " + index + ")." : " (list the shapes with read_object kind \"shape\" to find its index)."),
                    Mutated = true, Summary = "edit_textbox add",
                };
            }

            int target = RequireIndex(input, "edit_textbox", action);
            List<ShapeInfo> shapes = ListShapes(doc);
            if (target < 0 || target >= shapes.Count)
                throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " shape(s) in the document; indices are read_object kind \"shape\").");
            ShapeInfo info = shapes[target];
            // The index space is every shape in the document, pictures and lines included; a wrong index must not delete or
            // move one of those through a tool named for text boxes.
            if (info.Type != 17 && info.Type != 1)
                throw new ArgumentException("edit_textbox: shape " + target + " \"" + info.Name + "\" is a " + info.TypeName + ", not a text box or autoshape - this tool only edits and deletes those.");

            if (action == "delete")
            {
                info.Shape.Delete();
                return new ToolResult { Output = "Deleted shape " + target + " \"" + info.Name + "\". Later shape indices have shifted.", Mutated = true, Summary = "edit_textbox delete" };
            }

            if (text == null && !x.HasValue && !y.HasValue && !w.HasValue && !h.HasValue)
                throw new ArgumentException("edit_textbox: pass text and/or x, y, w, h for edit.");
            if (text != null)
            {
                try
                {
                    info.Shape.TextFrame.TextRange.Text = text;
                    ApplyAutoDirection((Word.Range)info.Shape.TextFrame.TextRange);
                }
                catch (Exception ex) { throw new ArgumentException("edit_textbox: shape " + target + " (" + info.TypeName + ") has no editable text frame: " + ex.Message); }
            }
            if (x.HasValue) { info.Shape.RelativeHorizontalPosition = 1; info.Shape.Left = x.Value; } // points from the page corner
            if (y.HasValue) { info.Shape.RelativeVerticalPosition = 1; info.Shape.Top = y.Value; }
            if (w.HasValue) info.Shape.Width = w.Value;
            if (h.HasValue) info.Shape.Height = h.Value;
            return new ToolResult { Output = "Updated shape " + target + " \"" + info.Name + "\".", Mutated = true, Summary = "edit_textbox edit" };
        }
    }
}
