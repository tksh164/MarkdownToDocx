using System;
using System.Text;

namespace MarkdownToDocx
{
    internal static class ExceptionTrapper
    {
        public static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = (Exception)e.ExceptionObject;
            var exceptionInfoText = BuildExceptionInformationText(ex);
            Console.Error.WriteLine("**** EXCEPTION ****");
            Console.Error.WriteLine(exceptionInfoText);
        }

        public static string BuildExceptionInformationText(Exception exception)
        {
            var builder = new StringBuilder();

            var ex = exception;
            while (true)
            {
                builder.AppendFormat("{0}: {1}", ex.GetType().FullName, ex.Message);
                builder.AppendLine();

                if (ex.Data.Count != 0)
                {
                    builder.AppendLine("Data:");
                    foreach (string key in ex.Data.Keys)
                    {
                        builder.AppendLine(string.Format("   {0}: {1}", key, ex.Data[key]));
                    }
                }

                builder.AppendLine("Stack Trace:");
                builder.AppendLine(ex.StackTrace);

                if (ex.InnerException == null) break;

                ex = ex.InnerException;
                builder.AppendLine(@"--- Inner exception is below ---");
            }

            return builder.ToString();
        }
    }
}
