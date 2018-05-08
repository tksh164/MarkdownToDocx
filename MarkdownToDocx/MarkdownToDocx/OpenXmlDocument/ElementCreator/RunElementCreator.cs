using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class RunElementCreator
    {
        public static Run CreateRunElement(string text, bool isBold, bool isItalic)
        {
            if (text == null) throw new ArgumentNullException(nameof(text), "Cannot use null for this parameter.");

            var emphasises = new List<OpenXmlElement>();
            if (isBold) emphasises.Add(new Bold());
            if (isItalic) emphasises.Add(new Italic());

            var runElement = new Run();
            if (isBold || isItalic)
            {
                runElement.AppendChild(new RunProperties(emphasises));
            }

            var textElement = new Text(text);
            if (text.StartsWith(" ") || text.EndsWith(" "))
            {
                textElement.Space = SpaceProcessingModeValues.Preserve;
            }
            runElement.AppendChild(textElement);

            return runElement;
        }
    }
}
