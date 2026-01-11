namespace ConfusionDetectionSandbox.Models
{
    internal class InputLogViewItem
    {
        public string TimestampUtc { get; init; } = string.Empty;
        public int TargetPid { get; init; }
        public string TargetAppName { get; init; } = string.Empty;
        public string Operation { get; init; } = string.Empty;
        public string X { get; init; } = string.Empty;
        public string Y { get; init; } = string.Empty;
        public string Delta { get; init; } = string.Empty;
        public string VirtualKey { get; init; } = string.Empty;
    }
}
