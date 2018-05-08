using System.Collections.Generic;

namespace MarkdownToDocx
{
    internal sealed class AppSettings
    {
        public string BaseFilePath { get; set; }
        public Dictionary<string, string> BaseFileUserStyleMap { get; set; }
    }
}
