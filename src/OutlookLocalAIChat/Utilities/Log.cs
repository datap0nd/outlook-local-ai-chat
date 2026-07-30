using System;
using System.IO;

namespace OutlookLocalAIChat.Utilities
{
    internal static class Log
    {
        public static void Error(string operation, Exception exception)
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookLocalAIChat",
                    "logs");
                Directory.CreateDirectory(directory);
                var path = Path.Combine(directory, "addin.log");
                var line =
                    DateTime.UtcNow.ToString("O") + " " +
                    operation + " " +
                    exception.GetType().Name + " " +
                    DiagnosticDetails.CodeForLog(exception) +
                    Environment.NewLine;
                File.AppendAllText(path, line);
            }
            catch
            {
            }
        }
    }
}
