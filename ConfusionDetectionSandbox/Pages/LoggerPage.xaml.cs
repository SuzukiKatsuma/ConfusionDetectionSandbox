using ConfusionDetectionSandbox.Models;
using ConfusionDetectionSandbox.Services;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using Windows.Storage.Streams;

namespace ConfusionDetectionSandbox.Pages
{
    /// <summary>
    /// ログ記録についての設定ページ
    /// </summary>
    public sealed partial class LoggerPage : Page
    {
        private bool isLogging = false;

        private ProcessInfo? selectedProcess = null;
        private int? loggingTargetPid;
        private string? loggingTargetAppName;

        private ObservableCollection<ProcessInfo> runningApps = [];
        private DispatcherQueueTimer? runningAppsTimer;

        private InputHookService? hookService;
        private CsvLogWriter? csvWriter;

        private readonly ObservableCollection<InputLogViewItem> recentLogs = new();
        private const int RecentLogLimit = 200;

        private Task? _loggingTask;
        private Channel<InputHookService.InputLogEntry>? _logChannel;
        private CancellationTokenSource? _loggingCts;

        public LoggerPage()
        {
            InitializeComponent();
            RunningAppsListView.ItemsSource = runningApps;
            RecentLogsListView.ItemsSource = recentLogs;

            Unloaded += LoggerPage_Unloaded;
            StartRunningAppsMonitor();
        }

        private void StartRunningAppsMonitor()
        {
            LoadRunningApplications();

            runningAppsTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
            runningAppsTimer.Interval = TimeSpan.FromSeconds(2);
            runningAppsTimer.Tick += RunningAppsTimer_Tick;
            runningAppsTimer.Start();
        }

        private void RunningAppsTimer_Tick(DispatcherQueueTimer sender, object args)
        {
            LoadRunningApplications();
        }

        private void LoggerPage_Unloaded(object sender, RoutedEventArgs e)
        {
            if (runningAppsTimer is not null)
            {
                runningAppsTimer.Tick -= RunningAppsTimer_Tick;
                runningAppsTimer.Stop();
                runningAppsTimer = null;
            }

            _ = StopLoggingIfNeededAsync();
        }

        private void RunningAppsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RunningAppsListView.SelectedItem is ProcessInfo processInfo)
            {
                selectedProcess = processInfo;
                TargetApplicationText.Text = $"Target App: {processInfo.DisplayName}";
            }
            else
            {
                selectedProcess = null;
                TargetApplicationText.Text = "Target App: Not selected";
            }
        }

        private async Task<BitmapImage?> ExtractIconAsync(string filePath)
        {
            try
            {
                Icon? icon = Icon.ExtractAssociatedIcon(filePath);
                if (icon == null) return null;

                using var bitmap = icon.ToBitmap();
                using var memory = new MemoryStream();
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                var randomAccessStream = new InMemoryRandomAccessStream();
                await randomAccessStream.WriteAsync(memory.ToArray().AsBuffer());
                randomAccessStream.Seek(0);
                await bitmapImage.SetSourceAsync(randomAccessStream);

                return bitmapImage;
            }
            catch
            {
                return null;
            }
        }

        private async void LoadRunningApplications()
        {
            List<ProcessInfo> currentApps = [];

            try
            {
                var processes = Process.GetProcesses()
                    .Where(p => !string.IsNullOrWhiteSpace(p.MainWindowTitle))
                    .OrderBy(p => p.ProcessName);

                foreach (var process in processes)
                {
                    try
                    {
                        var processInfo = new ProcessInfo
                        {
                            ProcessName = process.ProcessName,
                            MainWindowTitle = process.MainWindowTitle,
                            ProcessId = process.Id
                        };

                        // アイコンを取得
                        try
                        {
                            string? filePath = process.MainModule?.FileName;
                            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                            {
                                processInfo.Icon = await ExtractIconAsync(filePath);
                            }
                        }
                        catch
                        {
                            // アイコン取得失敗時は null のまま
                        }

                        currentApps.Add(processInfo);
                    }
                    catch
                    {
                        // アクセスできないプロセスはスキップ
                    }
                }

                if (currentApps.Count == 0)
                {
                    currentApps.Add(new ProcessInfo
                    {
                        ProcessName = "起動中のアプリケーションが見つかりません"
                    });
                }
            }
            catch (Exception ex)
            {
                currentApps.Clear();
                currentApps.Add(new ProcessInfo
                {
                    ProcessName = $"エラー: {ex.Message}"
                });
            }

            // 削除されたアプリを除去
            for (int i = runningApps.Count - 1; i >= 0; i--)
            {
                if (!currentApps.Any(app => app.ProcessId == runningApps[i].ProcessId))
                {
                    runningApps.RemoveAt(i);
                }
            }

            // 新しいアプリを追加
            foreach (var app in currentApps)
            {
                if (!runningApps.Any(existing => existing.ProcessId == app.ProcessId))
                {
                    // 正しい位置に挿入（ソート順を維持）
                    int insertIndex = 0;
                    for (int i = 0; i < runningApps.Count; i++)
                    {
                        if (string.Compare(app.ProcessName, runningApps[i].ProcessName, StringComparison.Ordinal) > 0)
                        {
                            insertIndex = i + 1;
                        }
                        else
                        {
                            break;
                        }
                    }
                    runningApps.Insert(insertIndex, app);
                }
            }
        }

        private async void ToggleLoggingButton_Click(object sender, RoutedEventArgs e)
        {
            isLogging = !isLogging;

            if (isLogging)
            {
                if (selectedProcess is null)
                {
                    isLogging = false;
                    LoggingStatusText.Text = "Logging Status: ⏸Stopped (Target not selected)";
                    ToggleLoggingButton.Content = "Start";
                    ToggleLoggingButton.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
                    return;
                }

                // ログ記録開始
                StartLogging(selectedProcess);
                LoggingStatusText.Text = "Logging Status: 👀Recording";
                ToggleLoggingButton.Content = "Stop";
                ToggleLoggingButton.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
            }
            else
            {
                // ログ記録停止
                await StopLoggingIfNeededAsync();
                LoggingStatusText.Text = "Logging Status: ⏸Stopped";
                ToggleLoggingButton.Content = "Start";
                ToggleLoggingButton.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
            }
        }

        private void StartLogging(ProcessInfo target)
        {
            loggingTargetPid = target.ProcessId;
            loggingTargetAppName = target.ProcessName;

            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ConfusionDetectionLogs");
            csvWriter = new CsvLogWriter(dir, "operation_log");

            // 非同期キュー（Channel）の初期化
            _loggingCts = new CancellationTokenSource();
            _logChannel = Channel.CreateUnbounded<InputHookService.InputLogEntry>();

            // 先にバックグラウンドタスクを回す
            var writer = csvWriter;
            var reader = _logChannel.Reader;
            var token = _loggingCts.Token;
            _loggingTask = Task.Run(() => ProcessLogQueueAsync(writer, reader, token));

            // 最後にフックを開始する
            hookService = new InputHookService { TargetPid = loggingTargetPid };
            hookService.OnLog += HookService_OnLog;
            hookService.Start();
        }

        private async Task ProcessLogQueueAsync(CsvLogWriter writer, ChannelReader<InputHookService.InputLogEntry> reader, CancellationToken ct)
        {
            var lastFlush = DateTime.UtcNow;
            try
            {
                // キューからデータが来るのを待機（バックグラウンドスレッドで実行）
                await foreach (var entry in reader.ReadAllAsync(ct))
                {
                    writer.Write(entry, loggingTargetAppName ?? "");

                    // 10秒ごとに一括書き出し
                    if ((DateTime.UtcNow - lastFlush).TotalSeconds >= 10)
                    {
                        writer.Flush();
                        lastFlush = DateTime.UtcNow;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                /* 停止時の正常な中断 */
            }
            finally
            {
                writer.Flush();
            }
        }

        private async Task StopLoggingIfNeededAsync()
        {
            if (hookService != null)
            {
                hookService.OnLog -= HookService_OnLog;
                hookService.Stop();
                hookService = null;
            }

            _loggingCts?.Cancel();
            _logChannel?.Writer.TryComplete();

            // バックグラウンドタスクが残りのデータを書き終えるのを待機
            if (_loggingTask != null)
            {
                try
                {
                    await _loggingTask;
                }
                catch
                {
                    /* ignore */
                }
                _loggingTask = null;
            }

            csvWriter?.Dispose();
            csvWriter = null;

            _logChannel = null;
            _loggingCts?.Dispose();
            _loggingCts = null;

            hookService = null;
            loggingTargetPid = null;
            loggingTargetAppName = null;
        }

        private void HookService_OnLog(InputHookService.InputLogEntry e)
        {
            if (loggingTargetAppName is null) return;

            var channel = _logChannel;
            if (channel == null) return;

            // UIへの表示更新 (DispatcherQueueでUIスレッドへ委譲)`
            DispatcherQueue.TryEnqueue(() =>
            {
                recentLogs.Add(new InputLogViewItem
                {
                    TimestampUtc = e.TimestampUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff"),
                    TargetPid = e.TargetPid,
                    TargetAppName = loggingTargetAppName,
                    Operation = e.Operation,
                    X = e.X?.ToString() ?? "",
                    Y = e.Y?.ToString() ?? "",
                    Delta = e.Delta?.ToString() ?? "",
                    VirtualKey = e.VirtualKey?.ToString() ?? "",
                });

                while (recentLogs.Count > RecentLogLimit)
                {
                    recentLogs.RemoveAt(0);
                }

                if (recentLogs.Count > 0)
                {
                    RecentLogsListView.ScrollIntoView(recentLogs[^1]);
                }
            });

            // CSV書き込み用キュー（Channel）へ投入
            // TryWrite はメモリへの保存のみ、I/O待ちによるフックの遅延が発生しなくなる。
            _logChannel?.Writer.TryWrite(e);
        }

    }
}

