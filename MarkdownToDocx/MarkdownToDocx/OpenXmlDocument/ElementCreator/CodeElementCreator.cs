using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class CodeElementCreator
    {
        public static Paragraph CreateCodeElement(string codeLineText, Style style, int leftIndentation)
        {
            if (codeLineText == null) throw new ArgumentNullException(nameof(codeLineText), "Cannot use null for this parameter.");
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Paragraph) throw new InvalidStyleTypeException(style);
            if (leftIndentation < 0) throw new ArgumentOutOfRangeException(nameof(leftIndentation), leftIndentation, "The left indentation is must be greater than equal 0.");

            var codeParagraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId()
                    {
                        Val = style.StyleId.Value,
                    }
                ),
                new Run(
                    new Text(codeLineText)
                    {
                        Space = SpaceProcessingModeValues.Preserve,
                    }
                )
            );

            if (leftIndentation > 0)
            {
                codeParagraph.ParagraphProperties.Append(
                    new Indentation()
                    {
                        Left = leftIndentation.ToString(),
                    }
                );
            }

            return codeParagraph;
        }
    }
}
