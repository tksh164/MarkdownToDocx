using System;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class NumberingElementCreator
    {
        public static (AbstractNum abstractNum, NumberingInstance numberingInstance) CreateNumberingElementPair()
        {
            var abstractNumId = AbstractNumIdGenerator.GetNewId();
            var oxmlAbstractNum = CreateAbstractNumElement(abstractNumId);

            var numberingInstanceId = NumberingInstanceIdGenerator.GetNewId();
            var oxmlNumberingInstance = CreateNumberingInstanceElement(numberingInstanceId, abstractNumId);

            return (oxmlAbstractNum, oxmlNumberingInstance);
        }

        private static class AbstractNumIdGenerator
        {
            private static int NextAbstractNumId = 0;  // This ID start from 0.

            public static int GetNewId()
            {
                return NextAbstractNumId++;
            }
        }

        private static class NumberingInstanceIdGenerator
        {
            private static int NextNumberingInstanceId = 1;  // This ID start from 1.

            public static int GetNewId()
            {
                return NextNumberingInstanceId++;
            }
        }

        private static AbstractNum CreateAbstractNumElement(int abstractNumId)
        {
            if (abstractNumId < 0) throw new ArgumentOutOfRangeException(nameof(abstractNumId), abstractNumId, "The prameter is must be greater than equal 0.");

            var abstractNum = new AbstractNum(
                new MultiLevelType()
                {
                    Val = MultiLevelValues.HybridMultilevel,
                }
            )
            {
                AbstractNumberId = abstractNumId,
            };

            // Create the Level elements from default values.
            foreach (var defaultValue in NumberingLevelDefaultValues)
            {
                var level = CreateLevelElement(defaultValue.LevelIndex, defaultValue.NumberingFormat, defaultValue.LevelText, defaultValue.IndentationLeft, defaultValue.IndentationHanging, defaultValue.FontName);
                abstractNum.Append(level);
            }

            return abstractNum;
        }

        private static Level CreateLevelElement(int levelIndex, NumberFormatValues numberFormat, string levelText, int indentationLeft, int indentationHanging, string bulletFontName)
        {
            if (levelIndex < 0) throw new ArgumentOutOfRangeException(nameof(levelIndex), levelIndex, "The prameter is must be greater than equal 0.");
            if (levelText == null) throw new ArgumentNullException(nameof(levelText), "Cannot use null for this parameter.");
            if (indentationLeft < 0) throw new ArgumentOutOfRangeException(nameof(indentationLeft), indentationLeft, "The prameter is must be greater than equal 0.");
            if (indentationHanging < 0) throw new ArgumentOutOfRangeException(nameof(indentationHanging), indentationHanging, "The prameter is must be greater than equal 0.");
            if (string.IsNullOrWhiteSpace(bulletFontName)) throw new ArgumentOutOfRangeException(nameof(bulletFontName), bulletFontName, "The bullet font name is not valid.");

            return new Level(
                new StartNumberingValue()
                {
                    Val = 1,  // The start number of the list.
                },
                new NumberingFormat()
                {
                    Val = numberFormat,  // The bullet style.
                },
                new LevelText()
                {
                    Val = levelText,  // The bullet character (text).
                },
                new LevelJustification()
                {
                    Val = LevelJustificationValues.Left,
                },
                new PreviousParagraphProperties(
                    new Indentation()
                    {
                        Left = indentationLeft.ToString(),
                        Hanging = indentationHanging.ToString(),
                    }
                ),
                new NumberingSymbolRunProperties(
                    // The font for the bullet.
                    new RunFonts()
                    {
                        Ascii = bulletFontName,
                        HighAnsi = bulletFontName,
                        ComplexScript = bulletFontName,
                        Hint = FontTypeHintValues.Default,
                    }
                )
            )
            {
                LevelIndex = levelIndex,
            };
        }

        private static NumberingInstance CreateNumberingInstanceElement(int numberingId, int abstractNumId)
        {
            if (numberingId < 0) throw new ArgumentOutOfRangeException(nameof(numberingId), numberingId, "The prameter is must be greater than equal 0.");
            if (abstractNumId < 0) throw new ArgumentOutOfRangeException(nameof(abstractNumId), abstractNumId, "The prameter is must be greater than equal 0.");

            return new NumberingInstance(
                new AbstractNumId()
                {
                    Val = abstractNumId,
                }
            )
            {
                NumberID = numberingId,
            };
        }

        private struct NumberingLevelDefaultValue
        {
            public int LevelIndex { get; set; }
            public int StartNumbering { get; set; }
            public NumberFormatValues NumberingFormat { get; set; }
            public string LevelText { get; set; }
            public LevelJustificationValues LevelJustification { get; set; }
            public int IndentationLeft { get; set; }
            public int IndentationHanging { get; set; }
            public string FontName { get; set; }
        }

        // A step on the ruler in Microsoft Word is equivalent 180 in the default setting. (and the default unit is inches)
        private const int IndentationLeftBase = 180 * 3;

        private static NumberingLevelDefaultValue[] NumberingLevelDefaultValues = new NumberingLevelDefaultValue[]
        {
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 0,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xb7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 1,  // Word's default: 720
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Symbol",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 1,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0x6f }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 2,  // Word's default: 1440,
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Courier New",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 2,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xa7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 3,  // Word's default: 2160
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Wingdings",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 3,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xb7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 4,  // Word's default: 2880
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Symbol",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 4,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0x6f }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 5,  // Word's default: 3600
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Courier New",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 5,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xa7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 6,  // Word's default: 4320
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Wingdings",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 6,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xb7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 7,  // Word's default: 5040
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Symbol",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 7,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0x6f }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 8,  // Word's default: 5760
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Courier New",
            },
            new NumberingLevelDefaultValue()
            {
                LevelIndex = 8,
                StartNumbering = 1,
                NumberingFormat = NumberFormatValues.Bullet,
                LevelText = Encoding.UTF8.GetString(new byte[] { 0xef, 0x82, 0xa7 }),
                LevelJustification = LevelJustificationValues.Left,
                IndentationLeft = IndentationLeftBase * 9,  // Word's default: 6480
                IndentationHanging = 360,                   // Word's default: 360
                FontName = "Wingdings",
            },
        };
    }
}
