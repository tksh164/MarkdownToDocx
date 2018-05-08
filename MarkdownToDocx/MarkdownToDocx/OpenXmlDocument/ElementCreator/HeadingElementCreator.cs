using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class HeadingElementCreator
    {
        public static Paragraph CreateHeadingElement(IEnumerable<OpenXmlElement> inlineElements, Style style)
        {
            if (inlineElements == null) throw new ArgumentNullException(nameof(inlineElements), "Cannot use null for this parameter.");
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Paragraph) throw new InvalidStyleTypeException(style);

            var header = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId()
                    {
                        Val = style.StyleId.Value,
                    }
                )
            );

            header.Append(inlineElements);

            return header;
        }
    }
}
