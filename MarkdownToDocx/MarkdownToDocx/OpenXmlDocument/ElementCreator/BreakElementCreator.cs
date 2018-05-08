using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.OpenXmlDocument.ElementCreator
{
    internal static class BreakElementCreator
    {
        public static Break CreateBreakElement()
        {
            return new Break();
        }
    }
}
