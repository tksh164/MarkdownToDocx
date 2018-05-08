using System;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json;
using MarkdownToDocx.OpenXmlDocument;

namespace MarkdownToDocx
{
    class Program
    {
        static int Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(ExceptionTrapper.UnhandledExceptionTrapper);

            if (args.Length == 1)
            {
                var docxFilePath = args[0];
                PrintStyleInformation(docxFilePath);
            }
            else if (args.Length == 3)
            {
                string markdownFilePath = args[0];
                string docxFilePath = args[1];
                string settingsFilePath = args[2];
                ConvertMarkdownToDocx(markdownFilePath, docxFilePath, settingsFilePath);
            }
            else
            {
                PrintUsage();
                return -1;
            }

            return 0;
        }

        private static void PrintUsage()
        {
            var commandName = Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule.FileName);

            Console.WriteLine(@"

 Usage: {0} <InputMarkdownFilePath> <OutputDocxFilePath> <SettingsFilePath>

   Convert the specified Markdown file to docx file.

 Usage: {0} <BaseDocxFilePath>

   Display the style information that used within the specified based docx file.

", commandName);
        }

        private static void ConvertMarkdownToDocx(string inputMarkdownFilePath, string outputDocxFilePath, string settingsFilePath)
        {
            var settings = JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(settingsFilePath));

            var converter = new MarkdownToDocxConverter(settings.BaseFilePath, settings.BaseFileUserStyleMap);
            converter.ConvertDocument(inputMarkdownFilePath, outputDocxFilePath);
        }

        private static void PrintStyleInformation(string docxFilePath)
        {
            var tempFilePath = Path.GetTempFileName();

            try
            {
                using (var manipulator = new WordDocumentManipulator(docxFilePath, tempFilePath, true))
                {
                    var styleSummaries = manipulator.StyleManager.GetStyleSummaries();

                    Console.WriteLine(@"The {0} styles contained within ""{1}"" file.", styleSummaries.Length, docxFilePath);

                    foreach (var styleSammary in styleSummaries)
                    {
                        Console.WriteLine(@"  ID:{0}, Type:{1}, Name:""{2}""", styleSammary.StyleId, styleSammary.StyleType.ToString(), styleSammary.StyleName);
                    }
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
