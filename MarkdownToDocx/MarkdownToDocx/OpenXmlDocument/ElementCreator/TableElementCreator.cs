using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class TableElementCreator
    {
        public static Table CreateTableElement(Style style, int leftIndentation)
        {
            if (style == null) throw new ArgumentNullException(nameof(style), "Cannot use null for this parameter.");
            if (style.Type != StyleValues.Table) throw new InvalidStyleTypeException(style);
            if (leftIndentation < 0) throw new ArgumentOutOfRangeException(nameof(leftIndentation), leftIndentation, "The left indentation is must be greater than equal 0.");

            var tableProperties = new TableProperties(
                new TableStyle()
                {
                    Val = style.StyleId.Value,
                },
                new TableWidth()
                {
                    Width = "0",
                    Type = TableWidthUnitValues.Auto,
                },
                new TableLook()
                {
                    FirstRow = true,           // "Header Row" in MS Word UI.
                    LastRow = false,           // "Total Row" in MS Word UI.
                    FirstColumn = false,       // "First Column" in MS Word UI.
                    LastColumn = false,        // "Last Column" in MS Word UI.
                    NoHorizontalBand = false,  // "Banded Rows" in MS Word UI.
                    NoVerticalBand = true,     // "Banded Columns" in MS Word UI.
                }
            );

            if (leftIndentation > 0)
            {
                tableProperties.Append(new TableIndentation()
                {
                    Width = leftIndentation,
                    Type = TableWidthUnitValues.Dxa,
                });
            }

            return new Table(tableProperties);
        }

        public static TableRow CreateTableRowElement()
        {
            return new TableRow(
                new TableRowProperties()
            );
        }

        public static TableCell CreateTableCellElement()
        {
            return new TableCell(
                new TableCellProperties()
            );
        }
    }
}
