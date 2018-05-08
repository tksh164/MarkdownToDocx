using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using MDT = Markdig.Extensions.Tables;
using MarkdownToDocx.OpenXmlDocument;

namespace MarkdownToDocx
{
    internal sealed class BlockConverter
    {
        private WordDocumentManipulator Manipulator { get; set; }
        private UserSettingStyleMap UserSettingStyleMap { get; set; }
        private string BaseFolderPathForRelativePath { get; set; }

        public BlockConverter(WordDocumentManipulator manipulator, UserSettingStyleMap userSettingStyleMap, string baseFolderPathForRelativePath)
        {
            Manipulator = manipulator;
            UserSettingStyleMap = userSettingStyleMap;
            BaseFolderPathForRelativePath = baseFolderPathForRelativePath;
        }

        public OpenXmlElement[] Convert(Block block)
        {
            try
            {
                return ConvertBlock(block, WordDocumentNumberingManager.OutsideOfListNumberingId, 0);
            }
            catch (Exception ex)
            {
                ex.Data.Add("Block:LineInMarkdown", block.Line);
                ex.Data.Add("Block:ColumnInMarkdown", block.Column);
                throw;
            }
        }

        private OpenXmlElement[] ConvertBlock(Block block, int numberingId, int nestLevel)
        {
            switch (block)
            {
                case ParagraphBlock paragraphBlock:
                    return new OpenXmlElement[] { ConvertParagraphBlock(paragraphBlock, numberingId, nestLevel) };

                case ListBlock listBlock:
                    return ConvertListBlock(listBlock, numberingId, nestLevel);

                case HeadingBlock headingBlock:
                    return new OpenXmlElement[] { ConvertHeadingBlock(headingBlock) };

                case QuoteBlock quoteBlock:
                    return ConvertQuoteBlock(quoteBlock, numberingId, nestLevel);

                case CodeBlock codeBlock:
                    return ConvertCodeBlock(codeBlock, numberingId, nestLevel);

                case MDT.Table tableBlock:
                    return new OpenXmlElement[] { ConvertTableBlock(tableBlock, numberingId, nestLevel) };

                case HtmlBlock htmlBlock:
                    // TODO
                    return new OpenXmlElement[] { };

                default:
                    throw new NotImplementedException(string.Format("Unknown block type: {0}", block.GetType().FullName));
            }
        }

        private Paragraph ConvertParagraphBlock(ParagraphBlock paragraphBlock, int numberingId, int nestLevel)
        {
            var inlineConverter = new InlineConverter(Manipulator, UserSettingStyleMap, BaseFolderPathForRelativePath);
            var oxmlInlineElements = inlineConverter.Convert(paragraphBlock.Inline);

            var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Paragraph, null);
            var oxmlParagraph = Manipulator.ElementCreator.CreateParagraphElement(oxmlInlineElements, styleId, numberingId, nestLevel);

            Manipulator.AdjustImageDimension(oxmlParagraph);

            return oxmlParagraph;
        }

        private OpenXmlElement[] ConvertListBlock(ListBlock listBlock, int numberingId, int nestLevel)
        {
            if (Manipulator.NumberingManager.IsOutsideOfListNumberingId(numberingId))
            {
                numberingId = Manipulator.NumberingManager.AddNumberingDefinition();
            }

            var oxmlElements = new List<OpenXmlElement>();
            foreach (ListItemBlock listItemBlock in listBlock)
            {
                if (listItemBlock.Count > 0)
                {
                    // The paragraph of the list item itself.
                    oxmlElements.Add(CreateListItemElement((ParagraphBlock)listItemBlock[0], numberingId, nestLevel, listBlock.BulletType));

                    // The nested blocks.
                    for (int i = 1; i < listItemBlock.Count; i++)
                    {
                        if (listItemBlock[i].GetType() == typeof(ListBlock))
                        {
                            oxmlElements.AddRange(ConvertBlock(listItemBlock[i], numberingId, nestLevel + 1));
                        }
                        else
                        {
                            oxmlElements.AddRange(ConvertBlock(listItemBlock[i], numberingId, nestLevel));
                        }
                    }
                }
            }

            return oxmlElements.ToArray();
        }

        private Paragraph CreateListItemElement(ParagraphBlock paragraphBlock, int numberingId, int nestLevel, char bulletTypeChar)
        {
            var inlineConverter = new InlineConverter(Manipulator, UserSettingStyleMap, BaseFolderPathForRelativePath);
            var oxmlInlineElements = inlineConverter.Convert(paragraphBlock.Inline);

            var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.ListItem, null);
            var oxmlParagraph = Manipulator.ElementCreator.CreateListItemElement(oxmlInlineElements, bulletTypeChar, styleId, numberingId, nestLevel);

            Manipulator.AdjustImageDimension(oxmlParagraph);

            return oxmlParagraph;
        }

        private Paragraph ConvertHeadingBlock(HeadingBlock headingBlock)
        {
            var inlineConverter = new InlineConverter(Manipulator, UserSettingStyleMap, BaseFolderPathForRelativePath);
            var oxmlInlineElements = inlineConverter.Convert(headingBlock.Inline);

            var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Heading, new UserSettingStyleMap.HeadingStyleMapArgs() { Level = headingBlock.Level });
            var oxmlHeading = Manipulator.ElementCreator.CreateHeadingElement(oxmlInlineElements, styleId);

            Manipulator.AdjustImageDimension(oxmlHeading);

            return oxmlHeading;
        }

        private OpenXmlElement[] ConvertCodeBlock(CodeBlock codeBlock, int numberingId, int nestLevel)
        {
            string styleId;
            if (codeBlock.GetType() == typeof(FencedCodeBlock))
            {
                // For the code block starting with ``` line and ending with ``` line.
                styleId = GetFencedCodeBlockStyleId(((FencedCodeBlock)codeBlock).Info);
            }
            else
            {
                // For the code block by indent.
                styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Code, null);
            }

            var oxmlParagraphs = new List<OpenXmlElement>();
            for (int i = 0; i < codeBlock.Lines.Count; i++)
            {
                var lineText = codeBlock.Lines.Lines[i].Slice.ToString();
                var oxmlParagraph = Manipulator.ElementCreator.CreateCodeElement(lineText, styleId, numberingId, nestLevel);
                oxmlParagraphs.Add(oxmlParagraph);
            }

            return oxmlParagraphs.ToArray();
        }

        private string GetFencedCodeBlockStyleId(string adhocStyleId)
        {
            if (string.IsNullOrWhiteSpace(adhocStyleId))
            {
                return UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Code, null);
            }

            if (Manipulator.StyleManager.ContainsStyleId(adhocStyleId))
            {
                return adhocStyleId;
            }
            else
            {
                var ex = new LackOfStyleWithinBaseFileException(string.Format("The base file does not contain the style ID: {0}", adhocStyleId));
                ex.Data.Add("LackedStyleIds", new string[] { adhocStyleId });
                throw ex;
            }
        }

        private OpenXmlElement[] ConvertQuoteBlock(QuoteBlock quoteBlock, int numberingId, int nestLevel)
        {
            var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Quote, null);

            var oxmlElements = new List<OpenXmlElement>();
            foreach (var block in quoteBlock)
            {
                switch (block)
                {
                    case ParagraphBlock paragraphBlock:
                        var oxmlParagraph = ConvertParagraphBlock(paragraphBlock, WordDocumentNumberingManager.OutsideOfListNumberingId, 0);
                        var oxmlQuoteParagraph = Manipulator.ElementCreator.CreateQuoteElement(oxmlParagraph, styleId, numberingId, nestLevel);

                        Manipulator.AdjustImageDimension(oxmlQuoteParagraph);

                        oxmlElements.Add(oxmlQuoteParagraph);
                        break;

                    case ListBlock listBlock:
                        var oxmlListItems = ConvertListBlock(listBlock, WordDocumentNumberingManager.OutsideOfListNumberingId, 0);
                        foreach (var oxmlListItem in oxmlListItems)
                        {
                            if (oxmlListItem.GetType() == typeof(Paragraph))
                            {
                                var oxmlListItemParagraph = (Paragraph)oxmlListItem;
                                var oxmlListItemQuoteParagraph = Manipulator.ElementCreator.CreateQuoteElement(oxmlListItemParagraph, styleId, numberingId, nestLevel);

                                Manipulator.AdjustImageDimension(oxmlListItemQuoteParagraph);

                                oxmlElements.Add(oxmlListItemQuoteParagraph);
                            }
                        }
                        break;

                    default:
                        throw new NotImplementedException(string.Format("Unknown block type within the quote block: {0}", block.GetType().FullName));
                }
            }

            return oxmlElements.ToArray();
        }

        private Table ConvertTableBlock(MDT.Table tableBlock, int numberingId, int nestLevel)
        {
            var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Table, null);

            var oxmlTable = Manipulator.ElementCreator.CreateTableElement(styleId, numberingId, nestLevel);

            foreach (MDT.TableRow tableRow in tableBlock)
            {
                var oxmlTableRow = Manipulator.ElementCreator.CreateTableRowElement();
                foreach (MDT.TableCell tableCell in tableRow)
                {
                    var oxmlTableCell = Manipulator.ElementCreator.CreateTableCellElement();
                    foreach (ParagraphBlock paragraphBlock in tableCell)
                    {
                        var oxmlParagraph = ConvertParagraphBlock(paragraphBlock, WordDocumentNumberingManager.OutsideOfListNumberingId, 0);
                        oxmlTableCell.Append(oxmlParagraph);
                    }

                    oxmlTableRow.Append(oxmlTableCell);
                }

                oxmlTable.Append(oxmlTableRow);
            }

            return oxmlTable;
        }
    }
}
