using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeAi.Shared.AttachmentText
{
    // Text-only extraction. No styling, no images, no layout - just the words,
    // in document order, so the model can read what an attachment says.
    internal static class OpenXmlReader
    {
        public static string ReadWord(string path)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, false))
            {
                W.Body body = doc.MainDocumentPart != null && doc.MainDocumentPart.Document != null
                    ? doc.MainDocumentPart.Document.Body
                    : null;
                if (body == null) return "";

                StringBuilder sb = new StringBuilder();
                foreach (W.Paragraph para in body.Descendants<W.Paragraph>())
                {
                    sb.AppendLine(string.Concat(para.Descendants<W.Text>().Select(t => t.Text)));
                }
                return sb.ToString().Trim();
            }
        }

        public static string ReadExcel(string path)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                if (wbPart == null) return "";

                string[] shared = wbPart.SharedStringTablePart != null
                    && wbPart.SharedStringTablePart.SharedStringTable != null
                    ? wbPart.SharedStringTablePart.SharedStringTable
                        .Elements<X.SharedStringItem>().Select(s => s.InnerText).ToArray()
                    : new string[0];

                StringBuilder sb = new StringBuilder();
                foreach (X.Sheet sheet in wbPart.Workbook.Descendants<X.Sheet>())
                {
                    WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
                    sb.AppendLine("# " + sheet.Name);

                    X.SheetData sheetData = wsPart.Worksheet.GetFirstChild<X.SheetData>();
                    if (sheetData != null)
                    {
                        foreach (X.Row row in sheetData.Elements<X.Row>())
                        {
                            sb.AppendLine(string.Join("\t",
                                row.Elements<X.Cell>().Select(c => CellText(c, shared))));
                        }
                    }
                    sb.AppendLine();
                }
                return sb.ToString().Trim();
            }
        }

        private static string CellText(X.Cell cell, string[] shared)
        {
            string v = cell.CellValue != null ? cell.CellValue.InnerText : "";
            if (cell.DataType != null && cell.DataType.Value == X.CellValues.SharedString)
            {
                int idx;
                if (int.TryParse(v, out idx) && idx >= 0 && idx < shared.Length) return shared[idx];
                return "";
            }
            if (cell.DataType != null && cell.DataType.Value == X.CellValues.InlineString)
            {
                return cell.InnerText;
            }
            return v;
        }

        public static string ReadPowerPoint(string path)
        {
            using (PresentationDocument doc = PresentationDocument.Open(path, false))
            {
                PresentationPart pPart = doc.PresentationPart;
                if (pPart == null) return "";

                StringBuilder sb = new StringBuilder();
                int n = 0;
                foreach (SlidePart slide in SlidesInOrder(pPart))
                {
                    n++;
                    sb.AppendLine("# Slide " + n);
                    foreach (A.Text t in slide.Slide.Descendants<A.Text>())
                    {
                        if (!string.IsNullOrWhiteSpace(t.Text)) sb.AppendLine(t.Text);
                    }
                    sb.AppendLine();
                }
                return sb.ToString().Trim();
            }
        }

        private static IEnumerable<SlidePart> SlidesInOrder(PresentationPart pPart)
        {
            P.SlideIdList idList = pPart.Presentation != null ? pPart.Presentation.SlideIdList : null;
            if (idList == null) yield break;

            foreach (P.SlideId slideId in idList.Elements<P.SlideId>())
            {
                yield return (SlidePart)pPart.GetPartById(slideId.RelationshipId);
            }
        }
    }
}
