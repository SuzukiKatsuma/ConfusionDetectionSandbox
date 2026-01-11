using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ConfusionDetectionSandbox.Services
{
    public sealed class CsvLogWriter : IDisposable
    {
        private readonly StreamWriter _writer;
        private readonly object _gate = new();

        public string FilePath { get; }

        public CsvLogWriter(string directoryPath, string fileNamePrefix)
        {
            Directory.CreateDirectory(directoryPath);
            FilePath = Path.Combine(directoryPath, $"{fileNamePrefix}_{DateTime.UtcNow:yyyyMMdd_HHmmss}_utc.csv");

            var fileStream = new FileStream(
                FilePath,
                FileMode.Append,
                FileAccess.Write,
                FileShare.ReadWrite);

            _writer = new StreamWriter(fileStream, new UTF8Encoding(true));
            _writer.AutoFlush = false;

            if (new FileInfo(FilePath).Length == 0)
            {
                _writer.WriteLine("timestamp,target_pid,target_app_name,operation,x,y,delta,virtual_key");
            }
        }

        public void Write(InputHookService.InputLogEntry e, string targetAppName)
        {
            string ts = e.TimestampUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture);

            string pid = e.TargetPid.ToString(CultureInfo.InvariantCulture);
            string app = EscapeCsv(targetAppName);
            string op = EscapeCsv(e.Operation);

            string x = e.X?.ToString(CultureInfo.InvariantCulture) ?? "";
            string y = e.Y?.ToString(CultureInfo.InvariantCulture) ?? "";
            string delta = e.Delta?.ToString(CultureInfo.InvariantCulture) ?? "";
            string vk = e.VirtualKey?.ToString(CultureInfo.InvariantCulture) ?? "";

            lock (_gate)
            {
                _writer.WriteLine($"{ts},{pid},{app},{op},{x},{y},{delta},{vk}");
            }
        }

        public void Flush()
        {
            lock (_gate)
            {
                try
                {
                    _writer.Flush();
                }
                catch
                {
                    // ignore
                }
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                _writer?.Dispose();
            }
        }

        private static string EscapeCsv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) < 0) return s;
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
    }
}
