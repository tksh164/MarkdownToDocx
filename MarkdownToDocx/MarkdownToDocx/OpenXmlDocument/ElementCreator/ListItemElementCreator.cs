using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class ListItemElementCreator
    {
        public static Paragraph CreateUnorderedListItemElement(IEnumerable<OpenXmlElement> inlineElements, Style style, int numberingId, int nestLevel, int leftIndentation)
        {
            if (inlineElements == null) throw new ArgumentNullException(nameof(inlineElements), "Cannot use null for this parameter.");
            if (nestLevel < 0) throw new ArgumentOutOfRangeException(nameof(nestLevel), nestLevel, "The list indentation is must be greater than equal 0.");
            if (numberingId < 0) throw new ArgumentOutOfRangeException(nameof(numberingId), numberingId, "The numbering ID is must be greater than equal 0.");
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Paragraph) throw new InvalidStyleTypeException(style);

            var listItem = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId()
                    {
                        Val = style.StyleId.Value,
                    },
                    new NumberingProperties(
                        new NumberingLevelReference()
                        {
                            Val = nestLevel,
                        },
                        new NumberingId()
                        {
                            Val = numberingId,
                        }
                    )
                )
            );

            listItem.Append(inlineElements);

            return listItem;
        }
    }
}
