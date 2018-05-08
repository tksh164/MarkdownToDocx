using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument
{
    internal sealed class WordDocumentManipulator : IDisposable
    {
        private Dictionary<string, string> ExistingImageRelationshipIdCache { get; set; }
        private WordprocessingDocument WordDocPackage { get; set; }

        public WordDocumentNumberingManager NumberingManager { get; private set; }
        public WordDocumentStyleManager StyleManager { get; private set; }
        public WordDocumentElementCreator ElementCreator { get; private set; }

        public void Dispose()
        {
            if (WordDocPackage != null)
            {
                WordDocPackage.Close();
                WordDocPackage.Dispose();
            }
        }

        public WordDocumentManipulator(string baseWordDocFilePath, string outputWordDocFilePath, bool overwriteExistingFile)
        {
            ExistingImageRelationshipIdCache = new Dictionary<string, string>();

            File.Copy(baseWordDocFilePath, outputWordDocFilePath, overwriteExistingFile);
            WordDocPackage = OpenEditableWordDocumentPackage(outputWordDocFilePath);

            NumberingManager = new WordDocumentNumberingManager(WordDocPackage.MainDocumentPart);
            StyleManager = new WordDocumentStyleManager(WordDocPackage.MainDocumentPart.StyleDefinitionsPart);
            ElementCreator = new WordDocumentElementCreator(NumberingManager, StyleManager);

            CleanupUnnecessaryElementsAsBaseFile(WordDocPackage);
        }

        private static WordprocessingDocument OpenEditableWordDocumentPackage(string wordDocFilePath)
        {
            var settings = new OpenSettings()
            {
                AutoSave = false,
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.NoProcess, FileFormatVersions.Office2013),
                MaxCharactersInPart = 0,
            };

            return WordprocessingDocument.Open(wordDocFilePath, true, settings);
        }

        private static void CleanupUnnecessaryElementsAsBaseFile(WordprocessingDocument wordDoc)
        {
            // Remove the existing document structure elements from the copied base file.
            wordDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
            wordDoc.MainDocumentPart.Document.Body.RemoveAllChildren<Table>();

            // Remove the existing numbering definitions from the copied base file.
            if (wordDoc.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                var numberingDefinitionsPart = wordDoc.MainDocumentPart.NumberingDefinitionsPart;
                if (numberingDefinitionsPart.Numbering != null)
                {
                    numberingDefinitionsPart.Numbering.RemoveAllChildren();
                }
                else
                {
                    // Purge the numbering definitions part because invalid package structure.
                    wordDoc.MainDocumentPart.DeletePart(wordDoc.MainDocumentPart.NumberingDefinitionsPart);
                }
            }

            // Remove the all existing image parts from the copied base file.
            wordDoc.MainDocumentPart.DeleteParts(wordDoc.MainDocumentPart.ImageParts);
        }

        public void Save()
        {
            WordDocPackage.MainDocumentPart.Document.Save();
            WordDocPackage.Save();
        }

        public void AppendElementsToDocumentBody(IEnumerable<OpenXmlElement> elements)
        {
            WordDocPackage.MainDocumentPart.Document.Body.Append(elements);
        }

        public string AddHyperlinkRelationship(Uri hyperlinkUri)
        {
            var mainDocPart = WordDocPackage.MainDocumentPart;
            var relationship = mainDocPart.AddHyperlinkRelationship(hyperlinkUri, true);
            return relationship.Id;
        }

        public string AddImagePart(Stream sourceImage)
        {
            var sourceImageHash = ImageHelper.CalculateSourceImageHash(sourceImage);
            if (ExistingImageRelationshipIdCache.ContainsKey(sourceImageHash))
            {
                // Get the relationship ID from the table. Herewith avoid the image duplication.
                return ExistingImageRelationshipIdCache[sourceImageHash];
            }
            else
            {
                // Add a new image part and feed the image data.
                var mainDocPart = WordDocPackage.MainDocumentPart;
                var imagePart = mainDocPart.AddImagePart(ImageHelper.DetectImagePartType(sourceImage));
                imagePart.FeedData(sourceImage);

                // Get the relationship ID for the image part and store to the table.
                var relationshipId = mainDocPart.GetIdOfPart(imagePart);
                ExistingImageRelationshipIdCache.Add(sourceImageHash, relationshipId);

                return relationshipId;
            }
        }

        public (double widthInch, double heightInch) GetImageDimensionInInch(string relationshipId)
        {
            var imagePart = WordDocPackage.MainDocumentPart.GetPartById(relationshipId);

            // Get the BitmapImage from the image part.
            var bitmapImage = new BitmapImage();
            using (var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
            {
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = stream;
                bitmapImage.EndInit();
            }

            var widthInch = bitmapImage.PixelWidth / bitmapImage.DpiX;
            var heightInch = bitmapImage.PixelHeight / bitmapImage.DpiY;

            return (widthInch, heightInch);
        }

        public void AdjustImageDimension(Paragraph paragraph)
        {
            // The cumulative left indentation by nesting.
            var cumulativeLeftIndentationEmu = GetCumulativeLeftIndentationInEmu(paragraph);

            (var maxContentAreaWidthEmu, var maxContentAreaHeightEmu) = GetMaxContentAreaDimensionInEmu();
            var availableContentAreaWidthEmu = maxContentAreaWidthEmu - cumulativeLeftIndentationEmu;

            var drawings = paragraph.Descendants<Drawing>();
            foreach (var drawing in drawings)
            {
                var dwExtent = drawing.Descendants<DW.Extent>().FirstOrDefault();
                if (dwExtent != null)
                {
                    // The original image dimension.
                    var originalImageWidthEmu = (double)dwExtent.Cx.Value;
                    var originalImageHeightEmu = (double)dwExtent.Cy.Value;

                    // Adjusted image dimension starts from the original image dimension.
                    var adjustedImageWidthEmu = originalImageWidthEmu;
                    var adjustedImageHeightEmu = originalImageHeightEmu;

                    // Adjust the image width.
                    if (availableContentAreaWidthEmu < originalImageWidthEmu)
                    {
                        var ratio = availableContentAreaWidthEmu / originalImageWidthEmu;
                        adjustedImageWidthEmu = availableContentAreaWidthEmu;
                        adjustedImageHeightEmu = adjustedImageHeightEmu * ratio;
                    }

                    // Adjust the image height.
                    if (maxContentAreaHeightEmu < adjustedImageHeightEmu)
                    {
                        var ratio = maxContentAreaHeightEmu / adjustedImageHeightEmu;
                        adjustedImageWidthEmu = adjustedImageWidthEmu * ratio;
                        adjustedImageHeightEmu = maxContentAreaHeightEmu;
                    }

                    dwExtent.Cx.Value = (long)adjustedImageWidthEmu;
                    dwExtent.Cy.Value = (long)adjustedImageHeightEmu;

                    var dExtents = drawing.Descendants<D.Extents>().FirstOrDefault();
                    if (dExtents != null)
                    {
                        dExtents.Cx.Value = (long)adjustedImageWidthEmu;
                        dExtents.Cy.Value = (long)adjustedImageHeightEmu;
                    }
                }
            }
        }

        private static double GetCumulativeLeftIndentationInEmu(Paragraph paragraph)
        {
            if (!Int32.TryParse(paragraph.ParagraphProperties?.Indentation?.Left?.Value, out int inheritedLeftIndentationDxa))
            {
                inheritedLeftIndentationDxa = 0;
            }
            return UnitConverter.DxaToEmu(inheritedLeftIndentationDxa);
        }

        private (double maxContentAreaWidthEmu, double maxContentAreaHeightEmu) GetMaxContentAreaDimensionInEmu()
        {
            var sectionProperties = WordDocPackage.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            var pageMargin = sectionProperties.GetFirstChild<PageMargin>();

            var maxContentAreaWidthDxa = pageSize.Width - pageMargin.Left - pageMargin.Right;
            var maxContentAreaWidthEmu = UnitConverter.DxaToEmu(maxContentAreaWidthDxa);

            var maxContentAreaHeightDxa = pageSize.Height - pageMargin.Top - pageMargin.Bottom;
            var maxContentAreaHeightEmu = UnitConverter.DxaToEmu(maxContentAreaHeightDxa);

            return (maxContentAreaWidthEmu, maxContentAreaHeightEmu);
        }

        internal static class ImageHelper
        {
            public static string CalculateSourceImageHash(Stream sourceImage)
            {
                sourceImage.Position = 0;  // Set the position to the beginning of the stream.

                SHA1CryptoServiceProvider sha1Provider = new SHA1CryptoServiceProvider();
                var sourceImageHash = sha1Provider.ComputeHash(sourceImage);

                sourceImage.Position = 0;  // Reset the stream position.

                return BitConverter.ToString(sourceImageHash);
            }

            public static ImagePartType DetectImagePartType(Stream img)
            {
                byte[] signature;
                img.Position = 0;  // Set the position to the beginning of the stream.
                using (var reader = new BinaryReader(img, Encoding.ASCII, true))
                {
                    signature = reader.ReadBytes(16);
                }
                img.Position = 0;  // Reset the stream position.

                // PNG
                if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a }))
                {
                    return ImagePartType.Png;
                }

                // JPEG (JFIF format)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0xff, 0xd8, 0xff, 0xe0 }) && StartsWithByteArraySequence(signature, 6, new byte[] { 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01 }))
                {
                    return ImagePartType.Jpeg;
                }

                // JPEG (Exif format)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0xff, 0xd8, 0xff, 0xe1 }) && StartsWithByteArraySequence(signature, 6, new byte[] { 0x45, 0x78, 0x69, 0x66, 0x00, 0x00 }))
                {
                    return ImagePartType.Jpeg;
                }

                // JPEG (Raw format)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0xff, 0xd8, 0xff, 0xdb }))
                {
                    return ImagePartType.Jpeg;
                }

                // GIF (GIF89a)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }))
                {
                    return ImagePartType.Gif;
                }

                // GIF (GIF87a)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }))
                {
                    return ImagePartType.Gif;
                }

                // TIFF (Little endian format)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x49, 0x49, 0x2A, 0x00 }))
                {
                    return ImagePartType.Tiff;
                }

                // TIFF (Big endian format)
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x4d, 0x4d, 0x00, 0x2a }))
                {
                    return ImagePartType.Tiff;
                }

                // BMP
                else if (StartsWithByteArraySequence(signature, 0, new byte[] { 0x42, 0x4d }))
                {
                    return ImagePartType.Bmp;
                }

                // Unknown
                else
                {
                    throw new ImageTypeNotSupportedException(string.Format("First 16 bytes are {0}", BitConverter.ToString(signature).ToLower().Replace("-", ", ")));
                }
            }

            private static bool StartsWithByteArraySequence(byte[] entireArray, int startIndex, byte[] partArray)
            {
                if (entireArray == null) throw new ArgumentNullException(nameof(entireArray), "Cannot be set null for this parameter.");
                if (partArray == null) throw new ArgumentNullException(nameof(partArray), "Cannot be set null for this parameter.");
                if (startIndex < 0) throw new ArgumentOutOfRangeException(nameof(startIndex), startIndex, "The acceptable value is greater than equal to 0.");

                if ((entireArray.Length - startIndex) < partArray.Length)
                {
                    return false;
                }

                for (int i = 0; i < partArray.Length; i++)
                {
                    if (entireArray[startIndex + i] != partArray[i])
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        public static class UnitConverter
        {
            public static double InchToEmu(double valueInches)
            {
                //
                // Convert the unit from inches to EMUs.
                // https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
                // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
                //
                const double EMUsPerInch = 914400;
                return valueInches * EMUsPerInch;
            }

            public static double DxaToInch(double valueDXAs)
            {
                //
                // Convert the unit from DXAs to Inches.
                // https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
                // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
                //
                const double DXAsPerInch = 1440;  // 1 inch = 72 points = 72 * 20 DXAs
                return valueDXAs / DXAsPerInch;
            }

            public static double DxaToEmu(double valueDXAs)
            {
                return InchToEmu(DxaToInch(valueDXAs));
            }
        }
    }
}
