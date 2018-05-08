using System;
using System.IO;
using System.Collections.Generic;
using Markdig;
using Markdig.Syntax;
using MarkdownToDocx.OpenXmlDocument;

namespace MarkdownToDocx
{
    internal sealed class MarkdownToDocxConverter
    {
        public string BaseFilePath { get; private set; }
        public UserSettingStyleMap UserSettingStyleMap { get; private set; }

        public MarkdownToDocxConverter(string baseFilePath, IReadOnlyDictionary<string, string> userSettingStyleMap)
        {
            BaseFilePath = baseFilePath ?? throw new ArgumentNullException(nameof(baseFilePath), "Cannot be set null for this parameter.");
            UserSettingStyleMap = new UserSettingStyleMap(userSettingStyleMap);
        }

        public void ConvertDocument(string sourceMarkdownFilePath, string outputDocxFilePath)
        {
            // Check the style existence within the docx document.
            var lackedStyleIds = GetLackedStyleIdsWithinBaseFile(BaseFilePath, UserSettingStyleMap);
            if (lackedStyleIds.Length != 0)
            {
                var ex = new LackOfStyleWithinBaseFileException(string.Format("The base file does not contain some style IDs: {0}", string.Join(", ", lackedStyleIds)));
                ex.Data.Add("LackedStyleIds", lackedStyleIds);
                throw ex;
            }

            var mdDoc = ParseMarkdownDocument(sourceMarkdownFilePath);

            // Build the Word document based-on the base docx file.
            using (var manipulator = new WordDocumentManipulator(BaseFilePath, outputDocxFilePath, true))
            {
                var baseFolderPathForRelativePath = Path.GetDirectoryName(sourceMarkdownFilePath);
                var blockConverter = new BlockConverter(manipulator, UserSettingStyleMap, baseFolderPathForRelativePath);

                foreach (var block in mdDoc)
                {
                    var elements = blockConverter.Convert(block);
                    manipulator.AppendElementsToDocumentBody(elements);
                }

                // Save the Word document.
                manipulator.Save();
            }
        }

        private static MarkdownDocument ParseMarkdownDocument(string markdownFilePath)
        {
            var markdownText = File.ReadAllText(markdownFilePath);
            var pipeline = new MarkdownPipelineBuilder().UsePipeTables().Build();
            return Markdown.Parse(markdownText, pipeline);
        }

        private static string[] GetLackedStyleIdsWithinBaseFile(string baseFilePath, UserSettingStyleMap userSettingStyleMap)
        {
            var tempFilePath = Path.GetTempFileName();

            try
            {
                using (var manipulator = new WordDocumentManipulator(baseFilePath, tempFilePath, true))
                {
                    // Gathering the lacked style IDs.
                    var lackedStyleIds = new List<string>();
                    foreach (var styleId in userSettingStyleMap.StyleMap.Values)
                    {
                        if (!manipulator.StyleManager.ContainsStyleId(styleId))
                        {
                            lackedStyleIds.Add(styleId);
                        }
                    }

                    return lackedStyleIds.ToArray();
                }
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }
    }
}
