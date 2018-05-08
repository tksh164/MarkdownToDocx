using System;

namespace MarkdownToDocx.OpenXmlDocument
{
    public sealed class ImageTypeNotSupportedException : Exception
    {
        public ImageTypeNotSupportedException()
        {
        }

        public ImageTypeNotSupportedException(string message)
            : base(message)
        {
        }

        public ImageTypeNotSupportedException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
