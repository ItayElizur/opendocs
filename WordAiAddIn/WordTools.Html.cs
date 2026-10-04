using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using OfficeAi.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace WordAiAddIn
{
    // Real COM tool execution against the live Word document, called from the
    // WebView2-hosted AgentLoop via the JSON WebMessage bridge.
    public static partial class WordTools
    {
        // Restricted-HTML insertion: a fixed, small supported set (no tables, images,
        // attributes, or nested lists). Parsed via XElement.Parse rather than a regex/
        // hand scanner for well-formedness checking, and validated against the
        // supported tag set before any Word write happens, so a bad fragment can't
        // leave a partial insert.
        private static readonly HashSet<string> HtmlBlockTags = new HashSet<string> { "p", "h1", "h2", "h3", "ul", "ol" };
        private static readonly HashSet<string> HtmlInlineTags = new HashSet<string> { "b", "strong", "i", "em", "u", "br" };

        private static System.Xml.Linq.XElement ParseHtmlFragment(string html)
        {
            // Normalize the most likely void-element mistake before parsing,
            // rather than failing the call over it.
            string normalized = System.Text.RegularExpressions.Regex.Replace(
                html, "<br\\s*>", "<br/>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            System.Xml.Linq.XElement root;
            try
            {
                root = System.Xml.Linq.XElement.Parse("<root>" + normalized + "</root>");
            }
            catch (System.Xml.XmlException ex)
            {
                throw new ArgumentException("Malformed HTML fragment (must be well-formed XHTML - closed tags, <br/> not <br>): " + ex.Message);
            }
            ValidateHtmlTags(root, true);
            return root;
        }

        private static void ValidateHtmlTags(System.Xml.Linq.XElement el, bool isBlockContext)
        {
            foreach (System.Xml.Linq.XElement child in el.Elements())
            {
                string tag = child.Name.LocalName.ToLowerInvariant();
                if (isBlockContext)
                {
                    if (tag == "li") continue; // only valid directly inside ul/ol, checked below
                    if (!HtmlBlockTags.Contains(tag))
                        throw new ArgumentException("Unsupported HTML tag '<" + tag + ">'. Supported: " +
                            string.Join(", ", HtmlBlockTags) + ", li (inside ul/ol), " + string.Join(", ", HtmlInlineTags) + ".");
                    if (tag == "ul" || tag == "ol")
                    {
                        foreach (System.Xml.Linq.XElement liOrOther in child.Elements())
                        {
                            if (liOrOther.Name.LocalName.ToLowerInvariant() != "li")
                                throw new ArgumentException("<" + tag + "> may only contain <li> children, found <" + liOrOther.Name.LocalName + ">.");
                            ValidateHtmlTags(liOrOther, false);
                        }
                    }
                    else
                    {
                        ValidateHtmlTags(child, false);
                    }
                }
                else
                {
                    if (!HtmlInlineTags.Contains(tag))
                        throw new ArgumentException("Unsupported HTML tag '<" + tag + ">' in inline content. Supported inline: " +
                            string.Join(", ", HtmlInlineTags) + ".");
                    if (tag != "br") ValidateHtmlTags(child, false);
                }
            }
        }

        // Writes one paragraph's inline content into `cursor` (collapsed at the start
        // of an empty paragraph). Word.Range is a COM reference type, so Collapse/Text
        // mutate the caller's own range - no ref parameter is needed for the
        // recursion to see the cursor advance.
        private static void WriteInlineNodes(Word.Range cursor, IEnumerable<System.Xml.Linq.XNode> nodes, bool bold, bool italic, bool underline)
        {
            foreach (System.Xml.Linq.XNode node in nodes)
            {
                System.Xml.Linq.XText textNode = node as System.Xml.Linq.XText;
                if (textNode != null)
                {
                    string text = textNode.Value;
                    if (text.Length == 0) continue;
                    cursor.Text = text;
                    cursor.Font.Bold = bold ? 1 : 0;
                    cursor.Font.Italic = italic ? 1 : 0;
                    cursor.Font.Underline = underline ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                System.Xml.Linq.XElement el = node as System.Xml.Linq.XElement;
                if (el == null) continue;
                string tag = el.Name.LocalName.ToLowerInvariant();
                if (tag == "br")
                {
                    // A soft line break within the same paragraph, not a new paragraph.
                    cursor.InsertBreak(Word.WdBreakType.wdLineBreak);
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    continue;
                }
                WriteInlineNodes(cursor, el.Nodes(),
                    bold || tag == "b" || tag == "strong",
                    italic || tag == "i" || tag == "em",
                    underline || tag == "u");
            }
        }

        // Inserts a validated HTML fragment (see ParseHtmlFragment) starting at `at`.
        // Each block element becomes its own new paragraph (same InsertParagraphAfter
        // idiom InsertContent uses) - the paragraph `at` points into is never merged
        // into, only new paragraphs after it are created.
        private static void InsertHtmlFragment(Word.Range at, string html)
        {
            System.Xml.Linq.XElement root = ParseHtmlFragment(html);
            Word.Range cursor = at;
            foreach (System.Xml.Linq.XElement block in root.Elements())
            {
                string tag = block.Name.LocalName.ToLowerInvariant();
                if (tag == "ul" || tag == "ol")
                {
                    foreach (System.Xml.Linq.XElement li in block.Elements())
                    {
                        cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        cursor.InsertParagraphAfter();
                        cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        Word.Paragraph para = cursor.Paragraphs[1];
                        WriteInlineNodes(cursor, li.Nodes(), false, false, false);
                        if (tag == "ul") para.Range.ListFormat.ApplyBulletDefault();
                        else para.Range.ListFormat.ApplyNumberDefault();
                    }
                }
                else
                {
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    cursor.InsertParagraphAfter();
                    cursor.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Paragraph para = cursor.Paragraphs[1];
                    WriteInlineNodes(cursor, block.Nodes(), false, false, false);
                    if (tag.Length == 2 && tag[0] == 'h')
                    {
                        para.Range.set_Style("Heading " + tag[1]);
                    }
                }
            }
        }

    }
}

