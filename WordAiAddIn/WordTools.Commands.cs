using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Spike 3: real COM tool execution against the live Word document, called
    // from the WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        // PP-5: mirrors WORD_COMMAND_SCHEMAS's `required` arrays in
        // WordAiAddIn/web-src/entry.ts exactly (minus "kind" itself, which is
        // validated separately in ApplyCommands before this table is
        // consulted) - the two must be edited together. This is the actual
        // guarantee: the TS schema is documentation the model reads, not a
        // validator that runs (not every provider enforces oneOf/const), so
        // this precheck is what turns a missing field into a specific,
        // per-command error instead of a raw COM/NullReference exception.
        private static readonly Dictionary<string, string[]> RequiredFields = new Dictionary<string, string[]>
        {
            ["set_bold"] = new[] { "startIndex", "endIndex", "value" },
            ["set_italic"] = new[] { "startIndex", "endIndex", "value" },
            ["set_heading"] = new[] { "index", "level" },
            ["find_replace"] = new[] { "find", "replace" },
            ["updateTextStyle"] = new[] { "target", "style", "fields" },
            ["updateParagraphStyle"] = new[] { "target", "style", "fields" },
            ["deleteBlocks"] = new[] { "target" },
            ["moveBlocks"] = new[] { "blockIndexes", "afterBlockIndex" },
            ["createParagraphBullets"] = new[] { "target" },
            ["set_bullet"] = new[] { "target" },   // alias - see the switch case
            ["deleteParagraphBullets"] = new[] { "target" },
            ["updateImageProperties"] = new[] { "imageIndex", "properties", "fields" },
            ["insertToc"] = new[] { "afterBlockIndex" },
        };

        // Fields where an explicit JSON null is a caller error rather than a
        // value - e.g. set_bold's "value": null previously reached
        // GetBoolean() and threw an opaque InvalidOperationException where
        // this clean "missing required field" message belongs. Deliberately
        // narrow: only add a field here after confirming null has no
        // legitimate meaning for it (Excel's set_cell "value" is the
        // counter-example - null there means "clear the cell").
        private static readonly Dictionary<string, string[]> NonNullFields = new Dictionary<string, string[]>
        {
            ["set_bold"] = new[] { "value" },
            ["set_italic"] = new[] { "value" },
            ["set_heading"] = new[] { "level" },
        };

        // Single source for the "what can I send?" answer in the
        // unknown-kind error. Mirrors ApplyCommands' switch - edit both
        // together (set_bullet is an accepted alias, listed so the model
        // discovers it).
        private static readonly string[] KnownCommandKinds =
        {
            "set_bold", "set_italic", "set_heading", "set_bullet", "find_replace",
            "updateTextStyle", "updateParagraphStyle", "deleteBlocks", "moveBlocks",
            "createParagraphBullets", "deleteParagraphBullets", "updateImageProperties", "insertToc",
        };

        // PP-12 Task 3 (the half PP-5 Task 4 Step 1 did not cover): each
        // result line is prefixed with the command's 0-based position in the
        // batch, and a summary header states how many succeeded/failed - with
        // partial batches now the norm (no rollback - Word COM offers no
        // batch transaction, and a hand-rolled undo would be less reliable
        // than this honest report; the user retains Word's own Ctrl+Z), the
        // model needs to know WHICH command in a batch of several identical
        // kinds failed, not just that "one of them" did.
        private static ToolResult ApplyCommands(JsonElement input)
        {
            var lines = new System.Text.StringBuilder();
            bool anyMutated = false;
            int failedCount = 0;
            int commandIndex = 0;
            int totalCount = 0;
            foreach (JsonElement cmd in input.GetProperty("commands").EnumerateArray())
            {
                totalCount++;
                string kind = null;
                try
                {
                    JsonElement kindEl;
                    if (!cmd.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("Command is missing a string \"kind\" field.");
                    kind = kindEl.GetString();
                    ToolArgs.ValidateRequired(kind, cmd, RequiredFields, "Command", NonNullFields);
                    switch (kind)
                    {
                        case "set_bold":
                            SetRunProperty(cmd, (range, value) => range.Bold = value ? 1 : 0);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "set_italic":
                            SetRunProperty(cmd, (range, value) => range.Italic = value ? 1 : 0);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "set_heading":
                            SetHeading(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "find_replace":
                            int replacements = FindReplace(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: {replacements} replacement(s)");
                            if (replacements > 0) anyMutated = true;
                            break;
                        case "updateTextStyle":
                            UpdateTextStyle(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "updateParagraphStyle":
                            UpdateParagraphStyle(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "deleteBlocks":
                            DeleteBlocksCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "moveBlocks":
                            MoveBlocksCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "createParagraphBullets":
                        {
                            string report = CreateParagraphBullets(cmd);
                            lines.AppendLine($"[{commandIndex}] {report}");
                            if (report.StartsWith("createParagraphBullets: 0 applied")) { /* nothing changed */ } else anyMutated = true;
                            break;
                        }
                        case "deleteParagraphBullets":
                        {
                            string report = DeleteParagraphBullets(cmd);
                            lines.AppendLine($"[{commandIndex}] {report}");
                            if (report.StartsWith("deleteParagraphBullets: 0 removed")) { /* nothing changed */ } else anyMutated = true;
                            break;
                        }
                        // Post-hoc addition (2026-08-27, user-reported): a model
                        // sent kind:"set_bullet" and got a dead-end "unknown
                        // command kind". It is an entirely reasonable guess -
                        // the neighbouring commands are set_bold/set_italic/
                        // set_heading, so a snake_case set_X for bullets reads
                        // as the obvious name, while the real ones are
                        // camelCase createParagraphBullets/deleteParagraphBullets.
                        // Rather than expect the model to memorise an
                        // inconsistency, accept the guess: set_bullet takes the
                        // same target as the two it delegates to, plus a
                        // value:true|false picking which.
                        case "set_bullet":
                        {
                            bool on = !cmd.TryGetProperty("value", out var bulletVal) || bulletVal.ValueKind != JsonValueKind.False;
                            string report = on ? CreateParagraphBullets(cmd) : DeleteParagraphBullets(cmd);
                            lines.AppendLine($"[{commandIndex}] set_bullet -> {report}");
                            bool noop = report.StartsWith("createParagraphBullets: 0 applied")
                                     || report.StartsWith("deleteParagraphBullets: 0 removed");
                            if (!noop) anyMutated = true;
                            break;
                        }
                        case "updateImageProperties":
                            UpdateImageProperties(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        case "insertToc":
                            InsertTocCmd(cmd);
                            lines.AppendLine($"[{commandIndex}] {kind}: ok"); anyMutated = true; break;
                        default:
                            // List what IS valid. A bare "unknown command kind"
                            // is a dead end - the model has no way to correct
                            // itself and typically retries the same wrong name.
                            lines.AppendLine($"[{commandIndex}] {kind}: unknown command kind. Valid kinds: " +
                                             string.Join(", ", KnownCommandKinds) + ".");
                            failedCount++; break;
                    }
                }
                catch (Exception ex)
                {
                    lines.AppendLine($"[{commandIndex}] " + (kind ?? "(unknown kind)") + ": ERROR - " + ex.Message);
                    failedCount++;
                }
                commandIndex++;
            }
            string summary = $"Applied {totalCount - failedCount} of {totalCount} command(s)" + (failedCount > 0 ? $" ({failedCount} failed)." : ".");
            return new ToolResult { Output = summary + "\n" + lines, Mutated = anyMutated, IsError = failedCount > 0, Summary = "apply_commands" };
        }

        private static void SetRunProperty(JsonElement cmd, Action<Word.Range, bool> apply)
        {
            int startIndex = cmd.GetProperty("startIndex").GetInt32();
            int endIndex = cmd.GetProperty("endIndex").GetInt32();
            bool value = cmd.GetProperty("value").GetBoolean();
            Word.Document doc = ActiveDoc;
            Word.Paragraphs paragraphs = doc.Paragraphs;
            endIndex = Math.Min(endIndex, paragraphs.Count - 1);
            Word.Range range = doc.Range(paragraphs[startIndex + 1].Range.Start, paragraphs[endIndex + 1].Range.End);
            apply(range, value);
        }

        private static void SetHeading(JsonElement cmd)
        {
            int index = cmd.GetProperty("index").GetInt32();
            int level = cmd.GetProperty("level").GetInt32();
            Word.Paragraph p = ActiveDoc.Paragraphs[index + 1];
            p.Range.set_Style(level == 0 ? "Normal" : "Heading " + level);
        }

        private static int FindReplace(JsonElement cmd)
        {
            string find = cmd.GetProperty("find").GetString();
            string replace = cmd.GetProperty("replace").GetString();
            bool matchCase = cmd.TryGetProperty("matchCase", out var mc) && mc.GetBoolean();
            Word.Find findObj = ActiveDoc.Content.Find;
            findObj.ClearFormatting();
            findObj.Text = find;
            findObj.Replacement.ClearFormatting();
            findObj.Replacement.Text = replace;
            findObj.MatchCase = matchCase;
            // wdReplaceAll only reports whether ANYTHING matched (true/false),
            // not how many - the caller's "N replacement(s)" message used to
            // always say 0 or 1 regardless of the real count. Loop
            // wdReplaceOne instead (the same one-at-a-time advance Word's own
            // "Replace All" button does internally) so the count is accurate.
            int count = 0;
            while (findObj.Execute(Replace: Word.WdReplace.wdReplaceOne))
            {
                count++;
                if (count > 10000) break; // safety net, not expected to ever trigger
            }
            return count;
        }

        // Returns each matched paragraph's 0-based index AND its already-
        // resolved Paragraph object (not just the index) - every caller used
        // to turn around and re-look up paragraphs[i + 1] per match, paying
        // Word's slow positional-indexing cost a second time. Returning the
        // object we already have in hand during the walk removes that
        // second lookup entirely.
        private static List<(int Index, Word.Paragraph Paragraph)> ResolveTargetParagraphs(JsonElement target)
        {
            string nodeType = target.TryGetProperty("nodeType", out var nt) && nt.ValueKind == JsonValueKind.String ? nt.GetString() : null;
            int? headingLevel = target.TryGetProperty("headingLevel", out var hl) && hl.ValueKind == JsonValueKind.Number ? hl.GetInt32() : (int?)null;
            string containsText = target.TryGetProperty("containsText", out var ct) && ct.ValueKind == JsonValueKind.String ? ct.GetString() : null;
            bool matchCase = target.TryGetProperty("matchCase", out var mc) && mc.ValueKind == JsonValueKind.True;
            HashSet<int> blockIndexes = null;
            if (target.TryGetProperty("blockIndexes", out var bi) && bi.ValueKind == JsonValueKind.Array)
            {
                blockIndexes = new HashSet<int>();
                foreach (JsonElement e in bi.EnumerateArray()) blockIndexes.Add(e.GetInt32());
            }
            string scope = target.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String ? sc.GetString() : "document";

            if (nodeType == null && containsText == null && blockIndexes == null)
            {
                throw new ArgumentException("Target must specify at least one of nodeType, containsText, or blockIndexes.");
            }

            Word.Paragraphs paragraphs = ActiveDoc.Paragraphs;
            int selStart = -1, selEnd = -1;
            if (scope == "selection")
            {
                Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                if (sel.Type != Word.WdSelectionType.wdNoSelection)
                {
                    selStart = sel.Range.Start;
                    selEnd = sel.Range.End;
                }
            }

            // Walks forward via the collection's own enumerator instead of
            // positional paragraphs[i + 1] indexing - Paragraphs is not a
            // real array in Word's COM object model, so indexing it by
            // position re-walks the document from the start on EVERY single
            // access, turning a full scan into roughly O(n^2) internally
            // (confirmed root cause of a real reported freeze in find_text/
            // get_headings, fixed there the same way). Every command that
            // funnels through this one function - updateTextStyle,
            // updateParagraphStyle, deleteBlocks, createParagraphBullets,
            // deleteParagraphBullets - inherits the fix.
            var result = new List<(int, Word.Paragraph)>();
            int i = 0;
            foreach (Word.Paragraph p in paragraphs)
            {
                // The body has several early `continue`s (skip this
                // paragraph, keep walking) - wrapping it in try/finally
                // guarantees i still advances exactly once per paragraph on
                // every path, since `continue` inside a try still runs its
                // finally before moving to the next iteration.
                try
                {
                    if (scope == "selection")
                    {
                        if (selStart == -1) continue;
                        if (p.Range.Start > selEnd || p.Range.End < selStart) continue;
                    }

                    if (blockIndexes != null && !blockIndexes.Contains(i)) continue;

                    string styleName = p.Range.get_Style().NameLocal;
                    bool isHeading = styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
                    bool isListItem = p.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;

                    if (nodeType == "heading" && !isHeading) continue;
                    if (nodeType == "paragraph" && (isHeading || isListItem)) continue;
                    if (nodeType == "listItem" && !isListItem) continue;

                    if (nodeType == "heading" && headingLevel.HasValue)
                    {
                        string levelDigits = new string(styleName.Where(char.IsDigit).ToArray());
                        if (!int.TryParse(levelDigits, out int actualLevel) || actualLevel != headingLevel.Value) continue;
                    }

                    if (containsText != null)
                    {
                        string text = p.Range.Text ?? "";
                        bool found = matchCase
                            ? text.Contains(containsText)
                            : text.IndexOf(containsText, StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!found) continue;
                    }

                    result.Add((i, p));
                }
                finally
                {
                    i++;
                }
            }
            return result;
        }

    }
}

