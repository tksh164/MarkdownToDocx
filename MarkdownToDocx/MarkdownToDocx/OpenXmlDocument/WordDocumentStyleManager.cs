using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument
{
    internal struct WordDocumentStyleSammary
    {
        public string StyleId { get; set; }
        public string StyleName { get; set; }
        public StyleValues StyleType { get; set; }
    }

    internal class WordDocumentStyleManager
    {
        private StyleDefinitionsPart StyleDefinitionsPart { get; set; }

        public WordDocumentStyleManager(StyleDefinitionsPart styleDefinitionsPart)
        {
            StyleDefinitionsPart = styleDefinitionsPart;
        }

        public Style FindStyleById(string styleId)
        {
            return StyleDefinitionsPart.Styles.Elements<Style>().SingleOrDefault((style) =>
            {
                return string.Compare(style.StyleId.Value, styleId, StringComparison.OrdinalIgnoreCase) == 0;
            });
        }

        public bool ContainsStyleId(string styleId)
        {
            return FindStyleById(styleId) != null ? true : false;
        }

        public WordDocumentStyleSammary[] GetStyleSummaries()
        {
            var styleSammaries = new List<WordDocumentStyleSammary>();

            var styles = StyleDefinitionsPart.Styles.Elements<Style>();
            foreach (var style in styles)
            {
                var styleName = style.GetFirstChild<StyleName>();
                styleSammaries.Add(new WordDocumentStyleSammary
                {
                    StyleId = style.StyleId.Value,
                    StyleName = styleName.Val.Value,
                    StyleType = style.Type.Value,
                });
            }

            return styleSammaries.ToArray();
        }
    }
}
