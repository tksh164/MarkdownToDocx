using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class HyperlinkElementCreator
    {
        public static Hyperlink CreateHyperlinkElement(string linkText, string hyperlinkRelationshipId, bool isBold, bool isItalic, Style style)
        {
            if (linkText == null) throw new ArgumentNullException(nameof(linkText), "Cannot use null for this parameter.");
            if (string.IsNullOrWhiteSpace(hyperlinkRelationshipId)) throw new ArgumentOutOfRangeException(nameof(hyperlinkRelationshipId), hyperlinkRelationshipId, "The hyperlink relationship ID is not valid.");
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Character) throw new InvalidStyleTypeException(style);

            var run = RunElementCreator.CreateRunElement(linkText, isBold, isItalic);
            if (run.RunProperties == null)
            {
                run.RunProperties = new RunProperties();
            }
            run.RunProperties.Append(
                new RunStyle()
                {
                    Val = style.StyleId.Value,
                }
            );

            return new Hyperlink(run)
            {
                Id = hyperlinkRelationshipId,
            };
        }
    }
}
