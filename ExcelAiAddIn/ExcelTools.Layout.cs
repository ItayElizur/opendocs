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
        private static void InsertDeleteRows(JsonElement op, bool insert)
        {
            int startRow = op.GetProperty("startRow").GetInt32();
            int count = op.GetProperty("count").GetInt32();
            Excel.Range rows = Sheet(op).Range[$"{startRow}:{startRow + count - 1}"];
            if (insert) rows.EntireRow.Insert(); else rows.EntireRow.Delete();
        }

        private static void InsertDeleteCols(JsonElement op, bool insert)
        {
            int startCol = op.GetProperty("startCol").GetInt32();
            int count = op.GetProperty("count").GetInt32();
            string startLetter = TextUtil.ColumnLetter(startCol);
            string endLetter = TextUtil.ColumnLetter(startCol + count - 1);
            Excel.Range cols = Sheet(op).Range[$"{startLetter}:{endLetter}"];
            if (insert) cols.EntireColumn.Insert(); else cols.EntireColumn.Delete();
        }

        private static void SortRange(JsonElement op)
        {
            string range = op.GetProperty("range").GetString();
            string byColumn = op.GetProperty("byColumn").GetString();
            string order = op.GetProperty("order").GetString();
            bool hasHeader = op.TryGetProperty("hasHeader", out var hh) && hh.ValueKind == JsonValueKind.True;
            Excel.Range target = Sheet(op).Range[range];
            Excel.Range key = Sheet(op).Range[byColumn + "1"];
            target.Sort(
                Key1: key,
                Order1: order == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending,
                Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
        }

        private static void SetRowHeight(JsonElement op)
        {
            int row = op.GetProperty("row").GetInt32();
            int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
            double heightPoints = op.GetProperty("heightPoints").GetDouble();
            Sheet(op).Range[$"{row}:{row + count - 1}"].RowHeight = heightPoints;
        }

        // Character-width units, NOT pixels: Excel's native ColumnWidth property has
        // no pixel unit at the COM layer. Converts genoffice's px schema using the
        // standard Calibri-11 approximation (px - 5) / 7 - a documented, deliberate
        // approximation, not exact for other default fonts.
        private static void SetColWidth(JsonElement op)
        {
            int col = op.GetProperty("column").GetInt32();
            int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
            double widthPx = op.GetProperty("widthPx").GetDouble();
            double charWidth = Math.Max(0, (widthPx - 5) / 7.0);
            string startLetter = TextUtil.ColumnLetter(col);
            string endLetter = TextUtil.ColumnLetter(col + count - 1);
            Sheet(op).Range[$"{startLetter}:{endLetter}"].ColumnWidth = charWidth;
        }

        private static void SetRowsHidden(JsonElement op)
        {
            int row = op.GetProperty("row").GetInt32();
            int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
            bool hidden = op.GetProperty("hidden").GetBoolean();
            Sheet(op).Range[$"{row}:{row + count - 1}"].EntireRow.Hidden = hidden;
        }

        private static void SetColsHidden(JsonElement op)
        {
            int col = op.GetProperty("column").GetInt32();
            int count = op.TryGetProperty("count", out var c) ? c.GetInt32() : 1;
            bool hidden = op.GetProperty("hidden").GetBoolean();
            string startLetter = TextUtil.ColumnLetter(col);
            string endLetter = TextUtil.ColumnLetter(col + count - 1);
            Sheet(op).Range[$"{startLetter}:{endLetter}"].EntireColumn.Hidden = hidden;
        }

        private static void SetFreeze(JsonElement op)
        {
            int rows = op.GetProperty("rows").GetInt32();
            int columns = op.GetProperty("columns").GetInt32();
            Excel.Worksheet sheet = Sheet(op);
            sheet.Activate();
            Excel.Window window = Globals.ThisAddIn.Application.ActiveWindow;
            window.FreezePanes = false;
            if (rows == 0 && columns == 0) return; // unfreeze only
            sheet.Cells[rows + 1, columns + 1].Select();
            window.FreezePanes = true;
        }

        private static void SetPageSetup(JsonElement op)
        {
            Excel.PageSetup setup = Sheet(op).PageSetup;
            if (op.TryGetProperty("orientation", out var orient) && orient.ValueKind == JsonValueKind.String)
                setup.Orientation = orient.GetString() == "landscape" ? Excel.XlPageOrientation.xlLandscape : Excel.XlPageOrientation.xlPortrait;

            bool hasScale = op.TryGetProperty("scale", out var scaleEl) && scaleEl.ValueKind == JsonValueKind.Number;
            bool hasFitWidth = op.TryGetProperty("fitToWidth", out var ftwEl) && ftwEl.ValueKind == JsonValueKind.Number;
            bool hasFitHeight = op.TryGetProperty("fitToHeight", out var fthEl) && fthEl.ValueKind == JsonValueKind.Number;
            bool hasFit = hasFitWidth || hasFitHeight;
            if (hasScale && hasFit)
                throw new ArgumentException(
                    "set_page_setup: 'scale' and 'fitToWidth'/'fitToHeight' are mutually exclusive in Excel " +
                    "(matching its own Page Setup UI). Pass either scale, or one/both fit values - not both.");

            if (hasFit)
            {
                setup.Zoom = false; // Zoom and FitToPages are mutually exclusive in Excel's own UI
                if (hasFitWidth) setup.FitToPagesWide = (int)ftwEl.GetDouble();
                // fitToHeight: 0 means "unlimited tall" - Excel expresses that as
                // FitToPagesTall = false, the single most common real-world
                // page-setup request ("fit on one page wide") that was otherwise
                // unexpressible.
                if (hasFitHeight)
                {
                    double fth = fthEl.GetDouble();
                    if (fth == 0) setup.FitToPagesTall = false;
                    else setup.FitToPagesTall = (int)fth;
                }
            }
            else if (hasScale)
            {
                int scaleVal = (int)scaleEl.GetDouble();
                if (scaleVal < 10 || scaleVal > 400)
                    throw new ArgumentOutOfRangeException("scale", "set_page_setup: scale must be between 10 and 400.");
                setup.Zoom = scaleVal;
            }

            if (op.TryGetProperty("printGridlines", out var pg))
                setup.PrintGridlines = pg.ValueKind == JsonValueKind.True;
            if (op.TryGetProperty("printHeadings", out var ph))
                setup.PrintHeadings = ph.ValueKind == JsonValueKind.True;
            if (op.TryGetProperty("printArea", out var pa) && pa.ValueKind == JsonValueKind.String)
                setup.PrintArea = pa.GetString();
            if (op.TryGetProperty("margins", out var margins) && margins.ValueKind == JsonValueKind.String)
            {
                // Point values match Excel's own ribbon presets (Normal/Wide/Narrow).
                double top, bottom, leftRight, header, footer;
                switch (margins.GetString())
                {
                    case "wide": top = 72; bottom = 72; leftRight = 72; header = 36; footer = 36; break;
                    case "narrow": top = 27; bottom = 27; leftRight = 18; header = 13.5; footer = 13.5; break;
                    default: top = 54; bottom = 54; leftRight = 43.2; header = 22.5; footer = 22.5; break; // "normal"
                }
                setup.TopMargin = top; setup.BottomMargin = bottom;
                setup.LeftMargin = leftRight; setup.RightMargin = leftRight;
                setup.HeaderMargin = header; setup.FooterMargin = footer;
            }
        }

        private static void AddSheet(JsonElement op)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet newSheet = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            newSheet.Name = op.GetProperty("name").GetString();
        }

        private static void DeleteSheet(JsonElement op)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            bool prevAlerts = app.DisplayAlerts;
            app.DisplayAlerts = false;
            try
            {
                Sheet(op).Delete();
            }
            finally
            {
                app.DisplayAlerts = prevAlerts;
            }
        }

        private static void DuplicateSheet(JsonElement op)
        {
            Excel.Worksheet source = Sheet(op);
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            source.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
            Excel.Worksheet copy = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            if (op.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
            {
                copy.Name = name.GetString();
            }
        }

        private static void SetSheetHidden(JsonElement op)
        {
            bool hidden = op.GetProperty("hidden").GetBoolean();
            Sheet(op).Visible = hidden ? Excel.XlSheetVisibility.xlSheetHidden : Excel.XlSheetVisibility.xlSheetVisible;
        }

        private static void MoveSheet(JsonElement op)
        {
            int position = op.GetProperty("position").GetInt32(); // 1-based, desired final position
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet target = Sheet(op);
            int count = wb.Worksheets.Count;
            int currentIndex = target.Index; // 1-based
            int clamped = Math.Max(1, Math.Min(position, count));
            if (clamped == currentIndex) return; // already there
            if (clamped > currentIndex)
            {
                int afterShift = clamped + 1;
                if (afterShift > count) target.Move(After: wb.Worksheets[count]);
                else target.Move(Before: wb.Worksheets[afterShift]);
            }
            else
            {
                target.Move(Before: wb.Worksheets[clamped]);
            }
        }

        private static void ProtectSheet(JsonElement op)
        {
            bool isProtected = op.GetProperty("protected").GetBoolean();
            Excel.Worksheet sheet = Sheet(op);
            if (isProtected) sheet.Protect();
            else sheet.Unprotect();
        }

    }
}

