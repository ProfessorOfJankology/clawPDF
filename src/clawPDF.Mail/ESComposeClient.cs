using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace clawSoft.clawPDF.Mail
{
    public sealed class ESComposeClient : IEmailClient
    {
        private const string EsComposeExePath = @"C:\Apps\ESCompose\ESCompose.exe";

        public bool IsClientInstalled
        {
            get { return File.Exists(EsComposeExePath); }
        }

        public bool ShowEmailClient(Email email)
        {
            if (!IsClientInstalled)
                return false;

            if (email == null || email.Attachments == null || email.Attachments.Count == 0)
                return false;

            var files = email.Attachments
                .Select(a => a == null ? null : a.Filename)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToArray();

            if (files.Length == 0)
                return false;

            try
            {
                var args = string.Join(" ", files.Select(Quote));

                var psi = new ProcessStartInfo
                {
                    FileName = EsComposeExePath,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(EsComposeExePath)
                        ?? Environment.CurrentDirectory
                };

                Process.Start(psi);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string Quote(string s)
        {
            return "\"" + s.Replace("\"", "\\\"") + "\"";
        }
    }
}
