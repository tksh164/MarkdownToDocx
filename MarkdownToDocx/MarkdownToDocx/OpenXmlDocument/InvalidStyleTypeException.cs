using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument
{
    public sealed class InvalidStyleTypeException : Exception
    {
        public InvalidStyleTypeException()
            : base()
        {
        }

        public InvalidStyleTypeException(Style style)
            : base()
        {
            SetStyleData(style);
        }

        public InvalidStyleTypeException(Style style, string message)
            : base(message)
        {
            SetStyleData(style);
        }

        public InvalidStyleTypeException(Style style, string message, Exception innerException)
            : base(message, innerException)
        {
            SetStyleData(style);
        }

        private void SetStyleData(Style style)
        {
            Data.Add("StyleType", style.Type.Value.ToString());
            Data.Add("StyleId", style.StyleId.Value);

            var styleName = style.GetFirstChild<StyleName>();
            Data.Add("StyleName", styleName.Val.Value);
        }
    }
}
