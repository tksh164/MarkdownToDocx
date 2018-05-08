using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class ParagraphElementCreator
    {
        public static Paragraph CreateParagraphElement(IEnumerable<OpenXmlElement> inlineElements, Style style, int leftIndentation)
        {
            if (inlineElements == null) throw new ArgumentNullException(nameof(inlineElements), "Cannot use null for this parameter.");
            if (style != null && style.Type != StyleValues.Paragraph) throw new InvalidStyleTypeException(style);

            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId()
                    {
                        Val = style?.StyleId?.Value ?? "",
                    }
                )
            );

            if (leftIndentation > 0)
            {
                paragraph.ParagraphProperties.Append(
                    new Indentation()
                    {
                        Left = leftIndentation.ToString(),
                    }
                );
            }

            paragraph.Append(inlineElements);

            return paragraph;
        }
    }
}
