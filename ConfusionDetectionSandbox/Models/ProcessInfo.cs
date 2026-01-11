using Microsoft.UI.Xaml.Media.Imaging;

namespace ConfusionDetectionSandbox.Models
{
    /// <summary>
    /// プロセス情報を保持するクラス
    /// </summary>
    internal class ProcessInfo
    {
        public string ProcessName { get; set; } = string.Empty;
        public string MainWindowTitle { get; set; } = string.Empty;
        public BitmapImage? Icon { get; set; }
        public int ProcessId { get; set; }

        public string DisplayName => $"{MainWindowTitle} ({ProcessName})";
    }
}
