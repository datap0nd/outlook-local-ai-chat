using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookLocalAIChat.Utilities
{
    // User-initiated self update. This is never exposed to the model:
    // only the Settings window calls it after an explicit confirmation.
    public static class SelfUpdater
    {
        public const string InstallerUrl =
            "https://github.com/datap0nd/outlook-local-ai-chat/releases/latest/download/OutlookLocalAIChatSetup.exe";

        public const int MaxInstallerBytes = 100 * 1024 * 1024;
        public const int MinInstallerBytes = 200 * 1024;

        public static async Task<string> DownloadInstallerAsync(
            CancellationToken cancellationToken)
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12;
            var path = Path.Combine(
                Path.GetTempPath(),
                "MetoAI-Update-" +
                Guid.NewGuid().ToString("N") +
                ".exe");
            using (var http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromMinutes(5);
                using (var response = await http.GetAsync(
                    InstallerUrl,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "The update download failed: " +
                            (int)response.StatusCode +
                            " " +
                            response.ReasonPhrase +
                            ".");
                    }

                    using (var source = await response.Content
                        .ReadAsStreamAsync()
                        .ConfigureAwait(true))
                    using (var target = File.Create(path))
                    {
                        var buffer = new byte[81920];
                        var total = 0;
                        while (true)
                        {
                            var read = await source.ReadAsync(
                                buffer,
                                0,
                                buffer.Length,
                                cancellationToken).ConfigureAwait(true);
                            if (read == 0)
                            {
                                break;
                            }

                            total += read;
                            if (total > MaxInstallerBytes)
                            {
                                throw new InvalidOperationException(
                                    "The downloaded installer is larger than expected.");
                            }

                            target.Write(buffer, 0, read);
                        }

                        if (total < MinInstallerBytes)
                        {
                            throw new InvalidOperationException(
                                "The downloaded installer looks incomplete.");
                        }
                    }
                }
            }

            using (var check = File.OpenRead(path))
            {
                var header = new byte[2];
                if (check.Read(header, 0, 2) != 2 ||
                    header[0] != 'M' ||
                    header[1] != 'Z')
                {
                    throw new InvalidOperationException(
                        "The downloaded file is not a Windows installer.");
                }
            }

            return path;
        }

        // The script receives the installer path as %1 so the file itself
        // stays pure ASCII regardless of the user's profile path.
        public static string BuildUpdateScript()
        {
            var builder = new StringBuilder();
            builder.AppendLine("@echo off");
            builder.AppendLine("set \"installer=%~1\"");
            builder.AppendLine("set tries=0");
            builder.AppendLine(":wait");
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq OUTLOOK.EXE\" | " +
                "find /I \"OUTLOOK.EXE\" >nul");
            builder.AppendLine("if errorlevel 1 goto install");
            builder.AppendLine("set /a tries+=1");
            builder.AppendLine("if %tries% GEQ 150 exit /b 1");
            builder.AppendLine("timeout /T 2 /NOBREAK >nul");
            builder.AppendLine("goto wait");
            builder.AppendLine(":install");
            builder.AppendLine(
                "\"%installer%\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART");
            builder.AppendLine("start \"\" outlook.exe");
            builder.AppendLine("del \"%installer%\"");
            return builder.ToString();
        }

        public static void LaunchUpdateAndQuitOutlook(
            object outlookApplication,
            string installerPath)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(
                    nameof(outlookApplication));
            }

            var scriptPath = Path.Combine(
                Path.GetTempPath(),
                "MetoAI-Update-" +
                Guid.NewGuid().ToString("N") +
                ".cmd");
            File.WriteAllText(
                scriptPath,
                BuildUpdateScript(),
                Encoding.ASCII);

            // The script runs through cmd.exe invoked by full path,
            // never through the .cmd shell association: locked-down
            // machines can remap or block script associations, which
            // surfaces as "not a valid application for this OS
            // platform". /d also skips any cmd AutoRun commands.
            var comSpec = Environment.GetEnvironmentVariable(
                "ComSpec");
            if (string.IsNullOrEmpty(comSpec))
            {
                comSpec = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.System),
                    "cmd.exe");
            }

            var start = new System.Diagnostics.ProcessStartInfo
            {
                FileName = comSpec,
                Arguments = "/d /c \"\"" + scriptPath + "\" \"" +
                    installerPath + "\"\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle =
                    System.Diagnostics.ProcessWindowStyle.Hidden
            };
            try
            {
                System.Diagnostics.Process.Start(start);
            }
            catch (Exception exception)
            {
                Log.Error("UpdateScriptLaunch", exception);
                // Last resort on machines that block running
                // anything from the temp folder: open the
                // downloaded installer itself so the user can click
                // through it after closing Outlook. Outlook is not
                // quit on this path.
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = installerPath,
                        UseShellExecute = true
                    });
                return;
            }

            dynamic application = outlookApplication;
            application.Quit();
        }

        public static string InstalledVersion()
        {
            try
            {
                return System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(
                        typeof(SelfUpdater).Assembly.Location)
                    .FileVersion ?? "unknown";
            }
            catch
            {
                return "unknown";
            }
        }
    }
}
