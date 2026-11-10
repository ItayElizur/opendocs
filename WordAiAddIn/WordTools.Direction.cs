using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    public static partial class WordTools
    {
        // A COM text write never flips a paragraph's reading order for Hebrew, so each write sets it from the
        // paragraph's own text. ReadingOrder also mirrors a left/right alignment and keeps centered or justified,
        // so alignment is never written. See WordTools.Direction.cs.md.
        private static void ApplyAutoDirection(Word.Range range)
        {
            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                switch (TextUtil.DetectDirection(paragraph.Range.Text))
                {
                    case TextDirection.Rtl:
                        paragraph.ReadingOrder = Word.WdReadingOrder.wdReadingOrderRtl;
                        break;
                    case TextDirection.Ltr:
                        paragraph.ReadingOrder = Word.WdReadingOrder.wdReadingOrderLtr;
                        break;
                }
            }
        }

        // SmartArt node text is TextFrame2, which ApplyAutoDirection (Word.Range) cannot take. Sets only the paragraph
        // direction, not alignment: the layout owns alignment. Left-to-right text only resets a node that is currently
        // right-to-left.
        private static void SetNodeText(dynamic node, string text)
        {
            node.TextFrame2.TextRange.Text = text;
            dynamic paragraph = node.TextFrame2.TextRange.ParagraphFormat;
            switch (TextUtil.DetectDirection(text))
            {
                case TextDirection.Rtl:
                    paragraph.TextDirection = (int)Microsoft.Office.Core.MsoTextDirection.msoTextDirectionRightToLeft;
                    break;
                case TextDirection.Ltr:
                    if ((int)paragraph.TextDirection == (int)Microsoft.Office.Core.MsoTextDirection.msoTextDirectionRightToLeft)
                        paragraph.TextDirection = (int)Microsoft.Office.Core.MsoTextDirection.msoTextDirectionLeftToRight;
                    break;
            }
        }
    }
}
