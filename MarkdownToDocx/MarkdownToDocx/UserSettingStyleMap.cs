using System;
using System.Collections.Generic;

namespace MarkdownToDocx
{
    internal sealed class UserSettingStyleMap
    {
        public enum StyleMapKeyType
        {
            Heading,
            Paragraph,
            Hyperlink,
            ListItem,
            Quote,
            Code,
            Table
        }

        public abstract class StyleMapArgs
        { }

        public sealed class HeadingStyleMapArgs : StyleMapArgs
        {
            public int Level { get; set; }
        }

        public IReadOnlyDictionary<string, string> StyleMap { get; private set; }

        public UserSettingStyleMap(IReadOnlyDictionary<string, string> styleMap)
        {
            StyleMap = styleMap ?? throw new ArgumentNullException(nameof(styleMap), "Cannot be set null for this parameter.");
        }

        public string GetStyleId(StyleMapKeyType styleKeyType, StyleMapArgs args)
        {
            var key = GetStyleMapKeyName(styleKeyType, args);
            return StyleMap.ContainsKey(key) ? StyleMap[key] : "";
        }

        private static string GetStyleMapKeyName(StyleMapKeyType styleKeyType, StyleMapArgs args)
        {
            switch (styleKeyType)
            {
                case StyleMapKeyType.Heading:
                    if (args == null) throw new ArgumentNullException(nameof(args), "Cannot be set null for this parameter.");

                    var headingArgs = (HeadingStyleMapArgs)args;
                    if (headingArgs.Level < 1) throw new ArgumentOutOfRangeException(nameof(headingArgs.Level), headingArgs.Level, "The acceptable value is greater than equal to 1.");

                    return "Heading" + headingArgs.Level.ToString();

                case StyleMapKeyType.Paragraph:
                    return "Paragraph";

                case StyleMapKeyType.Hyperlink:
                    return "Hyperlink";

                case StyleMapKeyType.ListItem:
                    return "List";

                case StyleMapKeyType.Quote:
                    return "Quote";

                case StyleMapKeyType.Code:
                    return "Code";

                case StyleMapKeyType.Table:
                    return "Table";

                default:
                    throw new NotImplementedException(string.Format("Unknown style mapping key type: {0}", styleKeyType.ToString()));
            }
        }
    }
}
