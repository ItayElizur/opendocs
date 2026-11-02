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

        private static void ApplyContentControlFields(dynamic cc, JsonElement input)
        {
            string title = OptStr(input, "title");
            string tag = OptStr(input, "tag");
            string text = OptStr(input, "text");
            bool? isChecked = OptBool(input, "checked");
            if (title != null) cc.Title = title;
            if (tag != null) cc.Tag = tag;
            if (text != null) cc.Range.Text = text;
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
                Word.Range at = after.HasValue ? RangeAtEndOfBlockText(after.Value) : RangeAtEndOfDocumentText();
                dynamic cc = AddContentControl(at, code);
                ApplyContentControlFields(cc, input);
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
            if (section > 1 && (bool)hf.LinkToPrevious) hf.LinkToPrevious = false;
            int replacedFields = (int)hf.Range.Fields.Count;
            hf.Range.Text = action == "clear" ? "" : Paras(text);
            // {PAGE}, {NUMPAGES}, {DATE}, {TIME}, {FILENAME}, {AUTHOR}, {TITLE} in the text become real fields.
            int insertedFields = 0;
            if (action == "set")
            {
                foreach (string fieldName in FieldTokens.Find(text))
                {
                    Word.Range r = hf.Range;
                    r.Find.ClearFormatting();
                    r.Find.Text = "{" + fieldName + "}";
                    r.Find.Forward = true;
                    r.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    if (!r.Find.Execute()) continue;
                    ((dynamic)r).Fields.Add(r, -1, fieldName, false);
                    insertedFields++;
                }
            }

            return new ToolResult
            {
                Output = (action == "clear" ? "Cleared " : "Set ") + "section " + section + " " + which + " (" + type.Replace('_', ' ') + ")."
                    + (insertedFields > 0 ? " Inserted " + insertedFields + " field(s)." : "")
                    + (replacedFields > 0 ? " Replaced " + replacedFields + " field(s) that were in it." : ""),
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
                    // Reference mark goes right after the Nth (default 1st) occurrence of afterText inside that paragraph.
                    int occurrence = OptInt(input, "occurrence") ?? 1;
                    if (occurrence < 1) throw new ArgumentException("edit_note: occurrence is 1-based.");
                    Word.Range para = ActiveDoc.Paragraphs[after.Value + 1].Range;
                    int paraEnd = para.End;
                    Word.Range search = para.Duplicate;
                    Word.Range nth = null;
                    int found = 0;
                    while (true)
                    {
                        search.Find.ClearFormatting();
                        search.Find.Text = afterText;
                        search.Find.Forward = true;
                        search.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        if (!search.Find.Execute() || search.End > paraEnd) break;
                        found++;
                        if (found == occurrence) nth = search.Duplicate;
                        search = ActiveDoc.Range(search.End, paraEnd);
                    }
                    if (found == 0)
                        throw new ArgumentException("edit_note: afterText '" + afterText + "' was not found in paragraph " + after.Value + ".");
                    if (nth == null)
                        throw new ArgumentException("edit_note: afterText '" + afterText + "' occurs " + found + " time(s) in paragraph " + after.Value + ", so occurrence " + occurrence + " does not exist.");
                    at = nth;
                    at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    at = RangeAtEndOfBlockText(after.Value);
                }
                dynamic note = notes.Add(at, Type.Missing, text);
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
                // Same convention as insert_content: afterBlockIndex -1 / omitted = start of the document.
                Word.Range at = RangeAfterBlock(OptInt(input, "afterBlockIndex") ?? -1);
                dynamic toc = doc.TablesOfContents.Add(at, true, upper ?? 1, lower ?? 3);
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
                    if (upper.HasValue) target.UpperHeadingLevel = upper.Value;
                    if (lower.HasValue) target.LowerHeadingLevel = lower.Value;
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
                string name = OptStr(input, "name");
                if (name != null) shape.Name = name;
                if (text != null) shape.TextFrame.TextRange.Text = text;

                string finalName = (string)shape.Name;
                int index = ListShapes(doc).FindIndex(s => s.Name == finalName);
                return new ToolResult
                {
                    Output = "Added text box \"" + finalName + "\" (shape index " + index + ").",
                    Mutated = true, Summary = "edit_textbox add",
                };
            }

            int target = RequireIndex(input, "edit_textbox", action);
            List<ShapeInfo> shapes = ListShapes(doc);
            if (target < 0 || target >= shapes.Count)
                throw new ArgumentOutOfRangeException("index", "index must be between 0 and " + (shapes.Count - 1) + " (" + shapes.Count + " shape(s) in the document; indices are read_object kind \"shape\").");
            ShapeInfo info = shapes[target];

            if (action == "delete")
            {
                info.Shape.Delete();
                return new ToolResult { Output = "Deleted shape " + target + " \"" + info.Name + "\". Later shape indices have shifted.", Mutated = true, Summary = "edit_textbox delete" };
            }

            if (text == null && !x.HasValue && !y.HasValue && !w.HasValue && !h.HasValue)
                throw new ArgumentException("edit_textbox: pass text and/or x, y, w, h for edit.");
            if (text != null)
            {
                if (info.Type == 6 || info.Type == 20)
                    throw new ArgumentException("edit_textbox: shape " + target + " is a " + info.TypeName + ", not a text box - its text lives in child shapes, which this tool cannot address.");
                try { info.Shape.TextFrame.TextRange.Text = text; }
                catch (Exception ex) { throw new ArgumentException("edit_textbox: shape " + target + " (" + info.TypeName + ") has no editable text frame: " + ex.Message); }
            }
            if (x.HasValue) info.Shape.Left = x.Value;
            if (y.HasValue) info.Shape.Top = y.Value;
            if (w.HasValue) info.Shape.Width = w.Value;
            if (h.HasValue) info.Shape.Height = h.Value;
            return new ToolResult { Output = "Updated shape " + target + " \"" + info.Name + "\".", Mutated = true, Summary = "edit_textbox edit" };
        }
    }
}
