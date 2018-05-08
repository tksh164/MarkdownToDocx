using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownToDocx.OpenXmlDocument.ElementCreator;

namespace MarkdownToDocx.OpenXmlDocument
{
    internal class WordDocumentNumberingManager
    {
        public enum BulletType
        {
            Ordered,
            Unordered,
        }

        private MainDocumentPart MainDocumentPart { get; set; }

        public WordDocumentNumberingManager(MainDocumentPart mainDocumentPart)
        {
            MainDocumentPart = mainDocumentPart;
        }

        private NumberingDefinitionsPart GetNumberingDefinitionsPart()
        {
            var numberingPart = MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart = MainDocumentPart.NumberingDefinitionsPart;
                numberingPart.Numbering = new Numbering();
            }
            return numberingPart;
        }

        public const int OutsideOfListNumberingId = -1;

        public bool IsOutsideOfListNumberingId(int numberingId)
        {
            return numberingId <= OutsideOfListNumberingId;
        }

        public int AddNumberingDefinition()
        {
            (var oxmlAbstractNum, var oxmlNumberingInstance) = NumberingElementCreator.CreateNumberingElementPair();

            var oxmlNumbering = GetNumberingDefinitionsPart().Numbering;
            oxmlNumbering.Append(oxmlAbstractNum, oxmlNumberingInstance);

            return oxmlNumberingInstance.NumberID;
        }

        public BulletType GetBulletType(char bulletTypeChar)
        {
            switch (bulletTypeChar)
            {
                case '-':
                case '*':
                    return BulletType.Unordered;

                case '1':
                    return BulletType.Ordered;

                default:
                    throw new NotImplementedException(string.Format("Unknown bullet type: {0}", bulletTypeChar));
            }
        }

        public void SetNumberingFormat(int numberingId, int levelIndex, BulletType bulletType)
        {
            if (IsOutsideOfListNumberingId(numberingId)) throw new InvalidOperationException("Outside of the list.");
            if (levelIndex < 0) throw new ArgumentOutOfRangeException(nameof(levelIndex), levelIndex, "The level index is must be greater than equal 0.");

            var oxmlLevel = GetLevelElement(numberingId, levelIndex);
            switch (bulletType)
            {
                case BulletType.Ordered:
                    oxmlLevel.NumberingFormat.Val.Value = NumberFormatValues.Decimal;
                    oxmlLevel.LevelText.Val.Value = "%1.";
                    oxmlLevel.NumberingSymbolRunProperties.RunFonts.Ascii = null;
                    oxmlLevel.NumberingSymbolRunProperties.RunFonts.HighAnsi = null;
                    oxmlLevel.NumberingSymbolRunProperties.RunFonts.ComplexScript = null;
                    break;

                case BulletType.Unordered:
                    // Use default values.
                    break;

                default:
                    throw new NotImplementedException(string.Format("Unknown bullet type: {0}", bulletType.ToString()));
            }
        }

        public int GetNumberingLeftIndentation(int numberingId, int levelIndex)
        {
            // It does not have the left indentation, because it does not currently in the list.
            if (IsOutsideOfListNumberingId(numberingId) && levelIndex == 0)
            {
                return 0;
            }

            var oxmlLevel = GetLevelElement(numberingId, levelIndex);
            return Int32.Parse(oxmlLevel.PreviousParagraphProperties.Indentation.Left.Value);
        }

        private Level GetLevelElement(int numberingId, int levelIndex)
        {
            var oxmlNumbering = GetNumberingDefinitionsPart().Numbering;

            var oxmlNumberingInstance = GetNumberingInstanceElement(oxmlNumbering, numberingId);

            int abstractNumId = oxmlNumberingInstance.AbstractNumId.Val.Value;
            var oxmlAbstractNum = GetAbstractNumElement(oxmlNumbering, abstractNumId);

            return GetLevelElement(oxmlAbstractNum, levelIndex);
        }

        private static NumberingInstance GetNumberingInstanceElement(Numbering numbering, int numberingId)
        {
            var oxmlNumberingInstance = numbering.Elements<NumberingInstance>().FirstOrDefault(ni =>
            {
                return ni.NumberID.Value == numberingId;
            });
            if (oxmlNumberingInstance == null)
            {
                throw new OpenXmlPackageException(string.Format("Couldn't find the NumberingInstance element by the Number ID (ID:{0})", numberingId));
            }

            return oxmlNumberingInstance;
        }

        private static AbstractNum GetAbstractNumElement(Numbering numbering, int abstractNumId)
        {
            var oxmlAbstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(an =>
            {
                return an.AbstractNumberId.Value == abstractNumId;
            });
            if (oxmlAbstractNum == null)
            {
                throw new OpenXmlPackageException(string.Format("Couldn't find the AbstractNum element by the AbstractNum ID (ID:{0})", abstractNumId));
            }

            return oxmlAbstractNum;
        }

        private static Level GetLevelElement(AbstractNum abstractNum, int levelIndex)
        {
            var oxmlLevel = abstractNum.Elements<Level>().FirstOrDefault(lv =>
            {
                return lv.LevelIndex.Value == levelIndex;
            });
            if (oxmlLevel == null)
            {
                throw new OpenXmlPackageException(string.Format("Couldn't find the Level element by the level index (LevelIndex:{0})", levelIndex));
            }

            return oxmlLevel;
        }
    }
}
