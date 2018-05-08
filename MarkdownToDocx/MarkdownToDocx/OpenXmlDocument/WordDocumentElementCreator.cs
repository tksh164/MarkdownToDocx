using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownToDocx.OpenXmlDocument.ElementCreator;

namespace MarkdownToDocx.OpenXmlDocument
{
    internal class WordDocumentElementCreator
    {
        private WordDocumentNumberingManager NumberingManager { get; set; }
        private WordDocumentStyleManager StyleManager { get; set; }

        public WordDocumentElementCreator(WordDocumentNumberingManager numberingManager, WordDocumentStyleManager styleManager)
        {
            NumberingManager = numberingManager;
            StyleManager = styleManager;
        }

        public Run CreateRunElement(string text, bool isBold, bool isItalic)
        {
            return RunElementCreator.CreateRunElement(text, isBold, isItalic);
        }

        public Paragraph CreateParagraphElement(IEnumerable<OpenXmlElement> inlineElements, string styleId, int numberingId, int nestLevel)
        {
            var style = StyleManager.FindStyleById(styleId);
            var leftIndentation = NumberingManager.GetNumberingLeftIndentation(numberingId, nestLevel);
            return ParagraphElementCreator.CreateParagraphElement(inlineElements, style, leftIndentation);
        }

        public Break CreateBreakElement()
        {
            return BreakElementCreator.CreateBreakElement();
        }

        public Paragraph CreateHeadingElement(IEnumerable<OpenXmlElement> inlineElements, string styleId)
        {
            var style = StyleManager.FindStyleById(styleId);
            return HeadingElementCreator.CreateHeadingElement(inlineElements, style);
        }

        public Paragraph CreateListItemElement(IEnumerable<OpenXmlElement> inlineElements, char bulletTypeChar, string styleId, int numberingId, int nestLevel)
        {
            var bulletType = NumberingManager.GetBulletType(bulletTypeChar);
            NumberingManager.SetNumberingFormat(numberingId, nestLevel, bulletType);

            var style = StyleManager.FindStyleById(styleId);
            var leftIndentation = NumberingManager.GetNumberingLeftIndentation(numberingId, nestLevel);
            return ListItemElementCreator.CreateUnorderedListItemElement(inlineElements, style, numberingId, nestLevel, leftIndentation);
        }

        public Run CreateImageElement(string imageRelationshipId, long iamgeWidthInEmus, long imageHeightInEmus, string fileName, string description)
        {
            return ImageElementCreator.CreateImageElement(imageRelationshipId, iamgeWidthInEmus, imageHeightInEmus, fileName, description);
        }

        public Hyperlink CreateHyperlinkElement(string linkText, string hyperlinkRelationshipId, bool isBold, bool isItalic, string styleId)
        {
            var style = StyleManager.FindStyleById(styleId);
            return HyperlinkElementCreator.CreateHyperlinkElement(linkText, hyperlinkRelationshipId, isBold, isItalic, style);
        }

        public Paragraph CreateCodeElement(string codeLineText, string styleId, int numberingId, int nestLevel)
        {
            var style = StyleManager.FindStyleById(styleId);
            var leftIndentation = NumberingManager.GetNumberingLeftIndentation(numberingId, nestLevel);
            return CodeElementCreator.CreateCodeElement(codeLineText, style, leftIndentation);
        }

        public Paragraph CreateQuoteElement(Paragraph quoteParagraph, string styleId, int numberingId, int nestLevel)
        {
            var style = StyleManager.FindStyleById(styleId);
            var leftIndentation = NumberingManager.GetNumberingLeftIndentation(numberingId, nestLevel);
            return QuoteElementCreator.CreateQuoteElement(quoteParagraph, style, leftIndentation);
        }

        public Table CreateTableElement(string styleId, int numberingId, int nestLevel)
        {
            var style = StyleManager.FindStyleById(styleId);
            var leftIndentation = NumberingManager.GetNumberingLeftIndentation(numberingId, nestLevel);
            return TableElementCreator.CreateTableElement(style, leftIndentation);
        }

        public TableRow CreateTableRowElement()
        {
            return TableElementCreator.CreateTableRowElement();
        }

        public TableCell CreateTableCellElement()
        {
            return TableElementCreator.CreateTableCellElement();
        }
    }
}
