using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class QuoteElementCreator
    {
        public static Paragraph CreateQuoteElement(Paragraph quoteParagraph, Style style, int leftIndentation)
        {
            if (quoteParagraph == null) throw new ArgumentNullException(nameof(quoteParagraph), "Cannot use null for this parameter.");
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Paragraph) throw new InvalidStyleTypeException(style);
            if (leftIndentation < 0) throw new ArgumentOutOfRangeException(nameof(leftIndentation), leftIndentation, "The left indentation is must be greater than equal 0.");

            quoteParagraph.ParagraphProperties.ParagraphStyleId.Val.Value = style.StyleId.Value;

            if (leftIndentation > 0)
            {
                quoteParagraph.ParagraphProperties.Append(
                    new Indentation()
                    {
                        Left = leftIndentation.ToString(),
                    }
                );
            }

            return quoteParagraph;
        }
    }
}
