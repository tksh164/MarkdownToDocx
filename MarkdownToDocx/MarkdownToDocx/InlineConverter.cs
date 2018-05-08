using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using Markdig.Syntax.Inlines;
using MarkdownToDocx.OpenXmlDocument;

namespace MarkdownToDocx
{
    internal sealed class InlineConverter
    {
        private WordDocumentManipulator Manipulator { get; set; }
        private UserSettingStyleMap UserSettingStyleMap { get; set; }
        private string BaseFolderPathForRelativePath { get; set; }

        public InlineConverter(WordDocumentManipulator manipulator, UserSettingStyleMap userSettingStyleMap, string baseFolderPathForRelativePath)
        {
            Manipulator = manipulator;
            UserSettingStyleMap = userSettingStyleMap;
            BaseFolderPathForRelativePath = baseFolderPathForRelativePath;
        }

        public OpenXmlElement[] Convert(ContainerInline containerInline)
        {
            try
            {
                var inlineElements = new List<OpenXmlElement>();

                foreach (var inline in containerInline)
                {
                    switch (inline)
                    {
                        case LiteralInline literalInline:
                            inlineElements.Add(ConvertLiteralInline(literalInline));
                            break;

                        case EmphasisInline emphasisInline:
                            inlineElements.AddRange(ConvertEmphasisInline(emphasisInline));
                            break;

                        case LinkInline linkInline:
                            inlineElements.Add(ConvertLinkInline(linkInline));
                            break;

                        case CodeInline codeInline:
                            // TODO
                            break;

                        case LineBreakInline lineBreakInline:
                            inlineElements.AddRange(ConvertLineBreakInline(lineBreakInline));
                            break;

                        default:
                            throw new NotImplementedException(string.Format("Unknown inline type: {0}", inline.GetType().FullName));
                    }
                }

                return inlineElements.ToArray();
            }
            catch (Exception ex)
            {
                ex.Data.Add("Inline:LineInMarkdown", containerInline.Line);
                ex.Data.Add("Inline:ColumnInMarkdown", containerInline.Column);
                throw;
            }
        }

        private OpenXmlElement ConvertLiteralInline(LiteralInline literalInline, bool isBoldInherited = false, bool isItalicInherited = false)
        {
            var normalText = literalInline.Content.ToString();
            return Manipulator.ElementCreator.CreateRunElement(normalText, isBoldInherited, isItalicInherited);
        }

        private OpenXmlElement[] ConvertEmphasisInline(EmphasisInline emphasisInline, bool isBoldInherited = false, bool isItalicInherited = false)
        {
            bool isBold = isBoldInherited || emphasisInline.IsDouble;
            bool isItatic = isItalicInherited || !emphasisInline.IsDouble;

            var elements = new List<OpenXmlElement>();
            foreach (var inline in emphasisInline)
            {
                switch (inline)
                {
                    case LiteralInline childLiteralInline:
                        elements.Add(ConvertLiteralInline(childLiteralInline, isBold, isItatic));
                        break;

                    case EmphasisInline childEmphasisInline:
                        elements.AddRange(ConvertEmphasisInline(childEmphasisInline, isBold, isItatic));
                        break;

                    case LinkInline linkInline:
                        elements.Add(ConvertLinkInline(linkInline, isBold, isItatic));
                        break;

                    default:
                        throw new NotImplementedException(string.Format("Unknown inline type: {0}", inline.GetType().FullName));
                }
            }

            return elements.ToArray();
        }

        private OpenXmlElement ConvertLinkInline(LinkInline linkInline, bool isBoldInherited = false, bool isItalicInherited = false)
        {
            if (linkInline.IsImage)
            {
                var explicitAbsoluteImagePath = GetExplicitAbsoluteImagePath(linkInline.Url, BaseFolderPathForRelativePath);
                var relationshipId = AddImagePartFromFile(explicitAbsoluteImagePath);

                // At this time, temporary uses the original image dimension.
                // In later, adjust the image dimension using the page settings and inherited indentation.
                (var originalImageWidthInch, var originalImageHeightInch) = Manipulator.GetImageDimensionInInch(relationshipId);
                var originalImageWidthEmu = WordDocumentManipulator.UnitConverter.InchToEmu(originalImageWidthInch);
                var originalImageHeightEmu = WordDocumentManipulator.UnitConverter.InchToEmu(originalImageHeightInch);

                var fileName = Path.GetFileName(explicitAbsoluteImagePath);
                var altText = GetLinkText(linkInline);
                return Manipulator.ElementCreator.CreateImageElement(relationshipId, (long)originalImageWidthEmu, (long)originalImageHeightEmu, fileName, altText);
            }
            else
            {
                var linkText = GetLinkText(linkInline);

                var hyperlinkUri = GetLinkTargetUri(linkInline.Url);
                var hyperlinkRelationshipId = Manipulator.AddHyperlinkRelationship(hyperlinkUri);

                var styleId = UserSettingStyleMap.GetStyleId(UserSettingStyleMap.StyleMapKeyType.Hyperlink, null);
                return Manipulator.ElementCreator.CreateHyperlinkElement(linkText, hyperlinkRelationshipId, isBoldInherited, isItalicInherited, styleId);
            }
        }

        private static string GetExplicitAbsoluteImagePath(string imagePath, string baseFolderPathForRelativePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath)) throw new ArgumentOutOfRangeException(nameof(imagePath), imagePath, "Invalid path.");
            if (string.IsNullOrWhiteSpace(baseFolderPathForRelativePath)) throw new ArgumentOutOfRangeException(nameof(baseFolderPathForRelativePath), baseFolderPathForRelativePath, "Invalid path.");

            // Web URI.
            if (imagePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || imagePath.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            {
                throw new NotImplementedException(string.Format("Cannot use URI for image file path in currently: {0}", imagePath));
            }

            var normalizedImagePath = imagePath.Replace('/', Path.DirectorySeparatorChar);

            // UNC path.
            if (normalizedImagePath.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase))
            {
                return normalizedImagePath;
            }

            // Absolute path
            if (Path.IsPathRooted(normalizedImagePath))
            {
                var root = Path.GetPathRoot(normalizedImagePath);
                if (string.Compare("\\", root, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    // Absolute path without volume label.
                    var explicitRoot = Path.GetPathRoot(baseFolderPathForRelativePath);
                    return Path.Combine(explicitRoot, normalizedImagePath.TrimStart(new char[] { '\\' }));
                }

                // Absolute path with volume label.
                return normalizedImagePath;
            }

            // Relative path.
            else
            {
                if (normalizedImagePath.StartsWith(".\\", StringComparison.OrdinalIgnoreCase))
                {
                    normalizedImagePath = normalizedImagePath.Substring(2);
                }
                return Path.Combine(baseFolderPathForRelativePath, normalizedImagePath);
            }
        }

        private string AddImagePartFromFile(string absoluteImageFilePath)
        {
            using (var stream = new FileStream(absoluteImageFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return Manipulator.AddImagePart(stream);
            }
        }

        private static Uri GetLinkTargetUri(string linkTarget)
        {
            if (string.IsNullOrWhiteSpace(linkTarget)) throw new ArgumentOutOfRangeException(nameof(linkTarget), linkTarget, "Invalid link target.");

            // Web URI.
            if (linkTarget.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || linkTarget.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            {
                return new Uri(linkTarget);
            }

            // Other link targets.
            return new Uri("file:///" + linkTarget);
        }

        private static string GetLinkText(LinkInline linkInline)
        {
            var inline = linkInline.FirstChild as LiteralInline;
            return inline == null ? "" : inline.Content.ToString();
        }

        private OpenXmlElement[] ConvertLineBreakInline(LineBreakInline lineBreakInline)
        {
            if (lineBreakInline.IsHard)
            {
                return new OpenXmlElement[] { Manipulator.ElementCreator.CreateBreakElement() };
            }

            return new OpenXmlElement[] { };
        }
    }
}
