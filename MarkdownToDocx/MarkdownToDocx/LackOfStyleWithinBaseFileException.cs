using System;

namespace MarkdownToDocx
{
    public sealed class LackOfStyleWithinBaseFileException : Exception
    {
        public LackOfStyleWithinBaseFileException()
            : base()
        {
        }

        public LackOfStyleWithinBaseFileException(string message)
            : base(message)
        {
        }

        public LackOfStyleWithinBaseFileException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
