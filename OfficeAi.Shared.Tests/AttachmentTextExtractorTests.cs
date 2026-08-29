using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using OfficeAi.Shared.AttachmentText;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using X = DocumentFormat.OpenXml.Spreadsheet;

public class AttachmentTextExtractorTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "AttTests_" + Guid.NewGuid().ToString("N"));

    public AttachmentTextExtractorTests() { Directory.CreateDirectory(_dir); }

    public void Dispose()
    {
        try { if (Directory.Exists(_dir)) Directory.Delete(_dir, true); } catch { }
    }

    private string Path2(string name) { return Path.Combine(_dir, name); }

    [Fact]
    public void PlainText_IsReturnedVerbatim()
    {
        string p = Path2("note.txt");
        File.WriteAllText(p, "hello\nworld");
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out bool truncated);
        Assert.True(ok);
        Assert.False(truncated);
        Assert.Equal("hello\nworld", text);
    }

    [Fact]
    public void Html_IsTagStrippedAndEntityDecoded()
    {
        string p = Path2("page.html");
        File.WriteAllText(p, "<html><head><style>x{}</style></head><body><h1>Title</h1><p>a &amp; b</p></body></html>");
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out _);
        Assert.True(ok);
        Assert.Contains("Title", text);
        Assert.Contains("a & b", text);
        Assert.DoesNotContain("<", text);
        Assert.DoesNotContain("x{}", text);
    }

    [Fact]
    public void Pdf_AndImages_ReturnFalse()
    {
        string pdf = Path2("doc.pdf");
        File.WriteAllBytes(pdf, new byte[] { 0x25, 0x50, 0x44, 0x46 });
        Assert.False(AttachmentTextExtractor.TryExtract(pdf, out _, out _));

        string png = Path2("pic.png");
        File.WriteAllBytes(png, new byte[] { 0x89, 0x50, 0x4E, 0x47 });
        Assert.False(AttachmentTextExtractor.TryExtract(png, out _, out _));
    }

    [Fact]
    public void MissingFile_ReturnsFalse()
    {
        Assert.False(AttachmentTextExtractor.TryExtract(Path2("nope.txt"), out _, out _));
    }

    [Fact]
    public void Truncation_FlagIsSetPastTheCap()
    {
        string p = Path2("big.txt");
        File.WriteAllText(p, new string('a', AttachmentTextExtractor.MaxChars + 50));
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out bool truncated);
        Assert.True(ok);
        Assert.True(truncated);
        Assert.Equal(AttachmentTextExtractor.MaxChars, text.Length);
    }

    [Fact]
    public void Docx_ParagraphTextIsExtracted()
    {
        string p = Path2("d.docx");
        using (WordprocessingDocument doc = WordprocessingDocument.Create(p, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = doc.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("First line."))),
                new W.Paragraph(new W.Run(new W.Text("Second line.")))));
            main.Document.Save();
        }
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out _);
        Assert.True(ok);
        Assert.Contains("First line.", text);
        Assert.Contains("Second line.", text);
    }

    [Fact]
    public void Xlsx_CellValuesAndSharedStringsAreExtracted()
    {
        string p = Path2("s.xlsx");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(p, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new X.Workbook();
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();

            X.SheetData data = new X.SheetData(
                new X.Row(
                    new X.Cell { CellReference = "A1", DataType = X.CellValues.String, CellValue = new X.CellValue("Name") },
                    new X.Cell { CellReference = "B1", DataType = X.CellValues.Number, CellValue = new X.CellValue("42") }));
            wsPart.Worksheet = new X.Worksheet(data);

            X.Sheets sheets = wbPart.Workbook.AppendChild(new X.Sheets());
            sheets.Append(new X.Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1U, Name = "Sheet1" });
            wbPart.Workbook.Save();
        }
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out _);
        Assert.True(ok);
        Assert.Contains("Sheet1", text);
        Assert.Contains("Name", text);
        Assert.Contains("42", text);
    }

    [Fact]
    public void Pptx_SlideTextIsExtracted()
    {
        string p = Path2("p.pptx");
        using (PresentationDocument doc = PresentationDocument.Create(p, PresentationDocumentType.Presentation))
        {
            PresentationPart pPart = doc.AddPresentationPart();
            pPart.Presentation = new P.Presentation();

            SlidePart slidePart = pPart.AddNewPart<SlidePart>();
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.Shape(
                            new P.TextBody(
                                new A.BodyProperties(),
                                new A.Paragraph(new A.Run(new A.Text("Hello from the deck"))))))));

            P.SlideIdList idList = pPart.Presentation.AppendChild(new P.SlideIdList());
            idList.Append(new P.SlideId { Id = 256U, RelationshipId = pPart.GetIdOfPart(slidePart) });
            pPart.Presentation.Save();
        }
        bool ok = AttachmentTextExtractor.TryExtract(p, out string text, out _);
        Assert.True(ok);
        Assert.Contains("Hello from the deck", text);
    }
}
