using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeAi.Shared;

namespace ExcelAiAddIn
{
    public static partial class ExcelTools
    {
        private static void SetFilter(JsonElement op)
        {
            string range = op.GetProperty("range").GetString();
            Sheet(op).Range[range].AutoFilter();
        }

        private static void ClearFilter(JsonElement op)
        {
            Excel.Worksheet sheet = Sheet(op);
            if (sheet.AutoFilterMode)
            {
                sheet.AutoFilterMode = false;
            }
        }

        private static void SetFilterCriteria(JsonElement op)
        {
            Excel.Worksheet sheet = Sheet(op);
            if (sheet.AutoFilter == null)
            {
                throw new InvalidOperationException("set_filter_criteria: no AutoFilter is active on this sheet - call set_filter first.");
            }
            int column = op.GetProperty("column").GetInt32(); // 0-based, relative to the AutoFilter range's first column
            Excel.Range filterRange = sheet.AutoFilter.Range;
            int fieldIndex = column + 1; // AutoFilter's Field parameter is 1-based, relative to the filter range - a common COM gotcha

            if (!op.TryGetProperty("values", out var values) || values.ValueKind == JsonValueKind.Null)
            {
                filterRange.AutoFilter(Field: fieldIndex); // toggling with no Criteria1 clears that column's filter
                return;
            }
            var criteria = new List<string>();
            foreach (JsonElement v in values.EnumerateArray()) criteria.Add(v.GetString());
            filterRange.AutoFilter(Field: fieldIndex, Criteria1: criteria.ToArray(), Operator: Excel.XlAutoFilterOperator.xlFilterValues);
        }

        private static Excel.XlFormatConditionOperator MapCfOperator(string op)
        {
            switch (op)
            {
                case "greaterThan": return Excel.XlFormatConditionOperator.xlGreater;
                case "lessThan": return Excel.XlFormatConditionOperator.xlLess;
                case "equal": return Excel.XlFormatConditionOperator.xlEqual;
                case "notEqual": return Excel.XlFormatConditionOperator.xlNotEqual;
                case "greaterEqual": return Excel.XlFormatConditionOperator.xlGreaterEqual;
                case "lessEqual": return Excel.XlFormatConditionOperator.xlLessEqual;
                case "between": return Excel.XlFormatConditionOperator.xlBetween;
                case "notBetween": return Excel.XlFormatConditionOperator.xlNotBetween;
                default:
                    throw new ArgumentException("add_conditional_format: unknown operator '" + op +
                        "'. Valid: greaterThan, lessThan, equal, notEqual, greaterEqual, lessEqual, between, notBetween.");
            }
        }

        private static double GetCfNumber(JsonElement rule, string field)
        {
            JsonElement v = rule.GetProperty(field);
            if (v.ValueKind == JsonValueKind.Number) return v.GetDouble();
            if (v.ValueKind == JsonValueKind.String)
            {
                double parsed;
                if (double.TryParse(v.GetString(), out parsed)) return parsed;
            }
            throw new ArgumentException("add_conditional_format: '" + field + "' must be a number (or a numeric string).");
        }

        private static Excel.XlContainsOperator MapCfTextMatch(string match)
        {
            switch (match)
            {
                case "contains": return Excel.XlContainsOperator.xlContains;
                case "notContains": return Excel.XlContainsOperator.xlDoesNotContain;
                case "beginsWith": return Excel.XlContainsOperator.xlBeginsWith;
                case "endsWith": return Excel.XlContainsOperator.xlEndsWith;
                default:
                    throw new ArgumentException("add_conditional_format: unknown text match '" + match +
                        "'. Valid: contains, notContains, beginsWith, endsWith.");
            }
        }

        // Returns a short description of what was created, for the batch result
        // line (PP-14 Task 5 Step 4).
        private static string AddConditionalFormat(JsonElement op)
        {
            JsonElement rangeEl;
            if (!op.TryGetProperty("range", out rangeEl) || rangeEl.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_conditional_format: missing required field \"range\".");
            string range = rangeEl.GetString();
            Excel.Range target;
            try
            {
                target = Sheet(op).Range[range];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new ArgumentException("add_conditional_format: '" + range + "' is not a valid range address.");
            }

            JsonElement rule;
            if (!op.TryGetProperty("rule", out rule) || rule.ValueKind != JsonValueKind.Object)
                throw new ArgumentException("add_conditional_format: missing required field \"rule\".");
            JsonElement kindEl;
            if (!rule.TryGetProperty("kind", out kindEl) || kindEl.ValueKind != JsonValueKind.String)
                throw new ArgumentException("add_conditional_format: rule is missing a string \"kind\" field.");
            string kind = kindEl.GetString();
            Excel.FormatCondition fc = null;
            string detail = "";

            switch (kind)
            {
                case "number":
                {
                    string oper = rule.GetProperty("operator").GetString();
                    Excel.XlFormatConditionOperator mappedOp = MapCfOperator(oper); // throws on unknown
                    double value = GetCfNumber(rule, "value");
                    bool needsSecond = mappedOp == Excel.XlFormatConditionOperator.xlBetween || mappedOp == Excel.XlFormatConditionOperator.xlNotBetween;
                    string formula2 = null;
                    if (needsSecond)
                    {
                        if (!rule.TryGetProperty("value2", out _))
                            throw new ArgumentException("add_conditional_format: 'value2' is required when operator is '" + oper + "'.");
                        formula2 = GetCfNumber(rule, "value2").ToString();
                    }
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, mappedOp, value.ToString(), formula2);
                    detail = "number " + oper + " " + value + (formula2 != null ? ".." + formula2 : "");
                    break;
                }
                case "text":
                {
                    string text = rule.GetProperty("text").GetString();
                    string matchName = rule.TryGetProperty("match", out var matchEl) && matchEl.ValueKind == JsonValueKind.String ? matchEl.GetString() : "contains";
                    Excel.XlContainsOperator matchOp = MapCfTextMatch(matchName); // throws on unknown
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlTextString, String: text, TextOperator: matchOp);
                    detail = "text " + matchName + " \"" + text + "\"";
                    break;
                }
                case "blank":
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlBlanksCondition);
                    detail = "blank";
                    break;
                case "duplicate":
                {
                    string modeName = rule.TryGetProperty("mode", out var modeEl) && modeEl.ValueKind == JsonValueKind.String ? modeEl.GetString() : "duplicate";
                    Excel.XlDupeUnique dupeUnique;
                    if (modeName == "duplicate") dupeUnique = Excel.XlDupeUnique.xlDuplicate;
                    else if (modeName == "unique") dupeUnique = Excel.XlDupeUnique.xlUnique;
                    else throw new ArgumentException("add_conditional_format: unknown mode '" + modeName + "'. Valid: duplicate, unique.");
                    fc = target.FormatConditions.AddUniqueValues();
                    ((Excel.UniqueValues)fc).DupeUnique = dupeUnique;
                    detail = modeName;
                    break;
                }
                case "top10":
                {
                    int rank = rule.TryGetProperty("rank", out var r) ? r.GetInt32() : 10;
                    bool percent = rule.TryGetProperty("percent", out var p) && p.ValueKind == JsonValueKind.True;
                    bool bottom = rule.TryGetProperty("bottom", out var b) && b.ValueKind == JsonValueKind.True;
                    Excel.Top10 top10 = target.FormatConditions.AddTop10();
                    top10.Rank = rank;
                    top10.Percent = percent;
                    top10.TopBottom = bottom ? Excel.XlTopBottom.xlTop10Bottom : Excel.XlTopBottom.xlTop10Top;
                    if (rule.TryGetProperty("format", out var top10Format))
                    {
                        if (top10Format.TryGetProperty("bold", out var bold)) top10.Font.Bold = bold.ValueKind == JsonValueKind.True;
                        if (top10Format.TryGetProperty("fontColor", out var fontColor)) top10.Font.Color = ColorUtil.HexToOle(fontColor.GetString());
                        if (top10Format.TryGetProperty("fillColor", out var fillColor)) top10.Interior.Color = ColorUtil.HexToOle(fillColor.GetString());
                    }
                    return "top10 range=" + range + " kind=top10 rank=" + rank + (percent ? "%" : "") + (bottom ? " bottom" : " top");
                    // Top10 doesn't implement FormatCondition in this PIA (confirmed via reflection) - format applied directly above, mirroring colorScale/dataBar's early-return pattern
                }
                case "formula":
                    fc = target.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: rule.GetProperty("formula").GetString());
                    detail = "formula";
                    break;
                case "colorScale":
                {
                    Excel.ColorScale scale = target.FormatConditions.AddColorScale(3);
                    if (rule.TryGetProperty("minColor", out var minC)) scale.ColorScaleCriteria[1].FormatColor.Color = ColorUtil.HexToOle(minC.GetString());
                    if (rule.TryGetProperty("midColor", out var midC)) scale.ColorScaleCriteria[2].FormatColor.Color = ColorUtil.HexToOle(midC.GetString());
                    if (rule.TryGetProperty("maxColor", out var maxC)) scale.ColorScaleCriteria[3].FormatColor.Color = ColorUtil.HexToOle(maxC.GetString());
                    return "colorScale range=" + range; // ColorScale/DataBar carry their own visual - no separate "format" object to apply below
                }
                case "dataBar":
                {
                    Excel.Databar bar = target.FormatConditions.AddDatabar();
                    if (rule.TryGetProperty("color", out var barColor))
                    {
                        bar.BarColor.Color = ColorUtil.HexToOle(barColor.GetString());
                    }
                    return "dataBar range=" + range;
                }
                default:
                    throw new ArgumentException("add_conditional_format: unknown rule kind '" + kind +
                        "'. Valid: number, text, blank, duplicate, top10, formula, colorScale, dataBar.");
            }

            if (fc != null && rule.TryGetProperty("format", out var format))
            {
                if (format.TryGetProperty("bold", out var bold)) fc.Font.Bold = bold.ValueKind == JsonValueKind.True;
                if (format.TryGetProperty("fontColor", out var fontColor)) fc.Font.Color = ColorUtil.HexToOle(fontColor.GetString());
                if (format.TryGetProperty("fillColor", out var fillColor)) fc.Interior.Color = ColorUtil.HexToOle(fillColor.GetString());
            }
            return kind + " range=" + range + " (" + detail + ")";
        }

        // PP-13: horizontal/vertical alignment name -> XlHAlign/XlVAlign, mirroring
        // this file's existing ChartTypes.ByName pattern (ShapeTypes moved to
        // OfficeAi.Shared in Phase 0).
        private static readonly Dictionary<string, Excel.XlHAlign> HAlignMap = new Dictionary<string, Excel.XlHAlign>
        {
            ["general"] = Excel.XlHAlign.xlHAlignGeneral,
            ["left"] = Excel.XlHAlign.xlHAlignLeft,
            ["center"] = Excel.XlHAlign.xlHAlignCenter,
            ["right"] = Excel.XlHAlign.xlHAlignRight,
            ["fill"] = Excel.XlHAlign.xlHAlignFill,
            ["justify"] = Excel.XlHAlign.xlHAlignJustify,
            ["centerAcrossSelection"] = Excel.XlHAlign.xlHAlignCenterAcrossSelection,
            ["distributed"] = Excel.XlHAlign.xlHAlignDistributed,
        };

        private static readonly Dictionary<string, Excel.XlVAlign> VAlignMap = new Dictionary<string, Excel.XlVAlign>
        {
            ["top"] = Excel.XlVAlign.xlVAlignTop,
            ["center"] = Excel.XlVAlign.xlVAlignCenter,
            ["bottom"] = Excel.XlVAlign.xlVAlignBottom,
            ["justify"] = Excel.XlVAlign.xlVAlignJustify,
            ["distributed"] = Excel.XlVAlign.xlVAlignDistributed,
        };

        // PP-13: border edge name -> XlBordersIndex, and style name -> (LineStyle, Weight).
        private static readonly Dictionary<string, Excel.XlBordersIndex> BorderEdgeMap = new Dictionary<string, Excel.XlBordersIndex>
        {
            ["left"] = Excel.XlBordersIndex.xlEdgeLeft,
            ["top"] = Excel.XlBordersIndex.xlEdgeTop,
            ["bottom"] = Excel.XlBordersIndex.xlEdgeBottom,
            ["right"] = Excel.XlBordersIndex.xlEdgeRight,
            ["insideHorizontal"] = Excel.XlBordersIndex.xlInsideHorizontal,
            ["insideVertical"] = Excel.XlBordersIndex.xlInsideVertical,
            ["diagonalDown"] = Excel.XlBordersIndex.xlDiagonalDown,
            ["diagonalUp"] = Excel.XlBordersIndex.xlDiagonalUp,
        };

        private static readonly string[] OutlineEdges = { "left", "top", "bottom", "right" };
        private static readonly string[] AllEdges = { "left", "top", "bottom", "right", "insideHorizontal", "insideVertical" };

        // Returns a note to append to the batch result line (e.g. borders skipped
        // on a single cell), or null when there is nothing to report.
        private static string FormatRange(JsonElement op)
        {
            string note = null;
            Excel.Range range = Sheet(op).Range[op.GetProperty("address").GetString()];
            if (op.TryGetProperty("bold", out var bold)) range.Font.Bold = bold.GetBoolean();
            if (op.TryGetProperty("italic", out var italic)) range.Font.Italic = italic.GetBoolean();
            if (op.TryGetProperty("numberFormat", out var nf)) range.NumberFormat = nf.GetString();
            if (op.TryGetProperty("fillColor", out var fc) && fc.ValueKind == JsonValueKind.String)
            {
                range.Interior.Color = ColorUtil.HexToOle(fc.GetString());
            }

            if (op.TryGetProperty("fontName", out var fn) && fn.ValueKind == JsonValueKind.String) range.Font.Name = fn.GetString();
            if (op.TryGetProperty("fontSize", out var fs) && fs.ValueKind == JsonValueKind.Number) range.Font.Size = fs.GetDouble();
            if (op.TryGetProperty("fontColor", out var fcol) && fcol.ValueKind == JsonValueKind.String) range.Font.Color = ColorUtil.HexToOle(fcol.GetString());
            if (op.TryGetProperty("strikethrough", out var st)) range.Font.Strikethrough = st.ValueKind == JsonValueKind.True;

            if (op.TryGetProperty("underline", out var underline))
            {
                if (underline.ValueKind == JsonValueKind.True || underline.ValueKind == JsonValueKind.False)
                {
                    range.Font.Underline = underline.ValueKind == JsonValueKind.True
                        ? Excel.XlUnderlineStyle.xlUnderlineStyleSingle
                        : Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                }
                else
                {
                    string u = underline.GetString();
                    switch (u)
                    {
                        case "none": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone; break;
                        case "single": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle; break;
                        case "double": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleDouble; break;
                        case "singleAccounting": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting; break;
                        case "doubleAccounting": range.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting; break;
                        default:
                            throw new ArgumentException("format_range: unknown underline '" + u + "'. Valid: none, single, double, singleAccounting, doubleAccounting (or a boolean).");
                    }
                }
            }

            if (op.TryGetProperty("horizontalAlignment", out var hAlign) && hAlign.ValueKind == JsonValueKind.String)
            {
                Excel.XlHAlign mapped;
                if (!HAlignMap.TryGetValue(hAlign.GetString(), out mapped))
                    throw new ArgumentException("format_range: unknown horizontalAlignment '" + hAlign.GetString() + "'. Valid: " + string.Join(", ", HAlignMap.Keys) + ".");
                range.HorizontalAlignment = mapped;
            }
            if (op.TryGetProperty("verticalAlignment", out var vAlign) && vAlign.ValueKind == JsonValueKind.String)
            {
                Excel.XlVAlign mapped;
                if (!VAlignMap.TryGetValue(vAlign.GetString(), out mapped))
                    throw new ArgumentException("format_range: unknown verticalAlignment '" + vAlign.GetString() + "'. Valid: " + string.Join(", ", VAlignMap.Keys) + ".");
                range.VerticalAlignment = mapped;
            }

            if (op.TryGetProperty("wrapText", out var wt)) range.WrapText = wt.ValueKind == JsonValueKind.True;
            if (op.TryGetProperty("textRotation", out var tr) && tr.ValueKind == JsonValueKind.Number)
            {
                int deg = tr.GetInt32();
                if (deg < -90 || deg > 90)
                    throw new ArgumentOutOfRangeException("textRotation", "textRotation must be between -90 and 90 degrees.");
                range.Orientation = deg;
            }
            if (op.TryGetProperty("indent", out var ind) && ind.ValueKind == JsonValueKind.Number)
            {
                int lvl = ind.GetInt32();
                if (lvl < 0 || lvl > 15)
                    throw new ArgumentOutOfRangeException("indent", "indent must be between 0 and 15.");
                range.IndentLevel = lvl;
            }

            if (op.TryGetProperty("borders", out var borders) && borders.ValueKind == JsonValueKind.Object)
            {
                note = ApplyBorders(range, borders);
            }
            return note;
        }

        // Returns a note when interior-border edges were silently skipped on a
        // single-cell range (Excel throws for insideHorizontal/insideVertical
        // there), or null otherwise.
        private static string ApplyBorders(Excel.Range range, JsonElement borders)
        {
            var edges = new List<string>();
            if (borders.TryGetProperty("preset", out var preset) && preset.ValueKind == JsonValueKind.String)
            {
                switch (preset.GetString())
                {
                    case "none":
                        foreach (Excel.XlBordersIndex idx in BorderEdgeMap.Values) range.Borders[idx].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                        return null; // clearing is the whole request; edges/style below don't apply to "none"
                    case "outline": edges.AddRange(OutlineEdges); break;
                    case "all": edges.AddRange(AllEdges); break;
                    case "thick-outline": edges.AddRange(OutlineEdges); break;
                    default:
                        throw new ArgumentException("format_range: unknown borders.preset '" + preset.GetString() + "'. Valid: none, outline, all, thick-outline.");
                }
            }
            if (borders.TryGetProperty("edges", out var edgesEl) && edgesEl.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement e in edgesEl.EnumerateArray())
                {
                    string edge = e.GetString();
                    if (!BorderEdgeMap.ContainsKey(edge))
                        throw new ArgumentException("format_range: unknown borders.edges value '" + edge + "'. Valid: " + string.Join(", ", BorderEdgeMap.Keys) + ".");
                    if (!edges.Contains(edge)) edges.Add(edge);
                }
            }
            if (edges.Count == 0) edges.AddRange(OutlineEdges); // borders object given with no preset/edges - sane default

            string styleName = borders.TryGetProperty("style", out var styleEl) && styleEl.ValueKind == JsonValueKind.String
                ? styleEl.GetString()
                : (preset.ValueKind == JsonValueKind.String && preset.GetString() == "thick-outline" ? "thick" : "thin");
            Excel.XlLineStyle lineStyle;
            Excel.XlBorderWeight weight;
            switch (styleName)
            {
                case "thin": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlThin; break;
                case "medium": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlMedium; break;
                case "thick": lineStyle = Excel.XlLineStyle.xlContinuous; weight = Excel.XlBorderWeight.xlThick; break;
                case "double": lineStyle = Excel.XlLineStyle.xlDouble; weight = Excel.XlBorderWeight.xlThick; break;
                case "dotted": lineStyle = Excel.XlLineStyle.xlDot; weight = Excel.XlBorderWeight.xlThin; break;
                case "dashed": lineStyle = Excel.XlLineStyle.xlDash; weight = Excel.XlBorderWeight.xlThin; break;
                case "none": lineStyle = Excel.XlLineStyle.xlLineStyleNone; weight = Excel.XlBorderWeight.xlThin; break;
                default:
                    throw new ArgumentException("format_range: unknown borders.style '" + styleName + "'. Valid: thin, medium, thick, double, dotted, dashed, none.");
            }

            int? oleColor = null;
            if (borders.TryGetProperty("color", out var colorEl) && colorEl.ValueKind == JsonValueKind.String)
                oleColor = ColorUtil.HexToOle(colorEl.GetString());

            bool singleCell = range.Cells.Count == 1;
            bool skippedInterior = false;
            foreach (string edge in edges)
            {
                if (singleCell && (edge == "insideHorizontal" || edge == "insideVertical"))
                {
                    skippedInterior = true;
                    continue; // Excel throws for interior borders on a single cell - silently skip, noted to the caller below
                }
                Excel.Border border = range.Borders[BorderEdgeMap[edge]];
                border.LineStyle = lineStyle;
                if (lineStyle != Excel.XlLineStyle.xlLineStyleNone)
                {
                    border.Weight = weight;
                    if (oleColor.HasValue) border.Color = oleColor.Value;
                }
            }
            return skippedInterior
                ? "insideHorizontal/insideVertical skipped (single-cell range has no interior edges)"
                : null;
        }

    }
}

