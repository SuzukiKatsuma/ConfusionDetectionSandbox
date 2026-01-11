using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ConfusionDetectionSandbox.Services
{
    public sealed partial class InputHookService : IDisposable
    {
        // ===== Public API =====
        public int? TargetPid { get; set; }

        /// <summary>
        /// ログは「押下(down)のみ」記録する。
        /// キーボード: WM_KEYDOWN / WM_SYSKEYDOWN
        /// マウス: WM_*BUTTONDOWN / WM_MOUSEWHEEL / WM_MOUSEHWHEEL
        /// </summary>
        public event Action<InputLogEntry>? OnLog;

        public void Start()
        {
            if (_running) return;
            _running = true;

            _kbProc = KeyboardHookCallback; // GC対策: フィールド保持
            _msProc = MouseHookCallback;

            // hMod は null でも可（LLフックは別プロセスDLL注入しない）
            _kbHook = SetWindowsHookEx(WH_KEYBOARD_LL, _kbProc, GetModuleHandle(null), 0);
            _msHook = SetWindowsHookEx(WH_MOUSE_LL, _msProc, GetModuleHandle(null), 0);

            if (_kbHook == IntPtr.Zero || _msHook == IntPtr.Zero)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

            _cts = new CancellationTokenSource();
            _ = Task.Run(() => DrainLoop(_cts.Token));
        }

        public void Stop()
        {
            if (!_running) return;
            _running = false;

            try { _cts?.Cancel(); } catch { }

            if (_kbHook != IntPtr.Zero) UnhookWindowsHookEx(_kbHook);
            if (_msHook != IntPtr.Zero) UnhookWindowsHookEx(_msHook);

            _kbHook = IntPtr.Zero;
            _msHook = IntPtr.Zero;

            _kbProc = null;
            _msProc = null;

            _cts = null;
        }

        public void Dispose() => Stop();

        // ===== internal queue (I/O はコールバックでしない) =====
        private readonly ConcurrentQueue<InputLogEntry> _queue = new();
        private CancellationTokenSource? _cts;

        private async Task DrainLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                while (_queue.TryDequeue(out var e))
                {
                    try { OnLog?.Invoke(e); } catch { /* swallow */ }
                }
                await Task.Delay(10, token).ConfigureAwait(false);
            }
        }

        // ===== hooks =====
        private const int WH_KEYBOARD_LL = 13;
        private const int WH_MOUSE_LL = 14;

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_MOUSEHWHEEL = 0x020E;

        private bool _running;
        private IntPtr _kbHook = IntPtr.Zero;
        private IntPtr _msHook = IntPtr.Zero;

        private HookProc? _kbProc;
        private HookProc? _msProc;

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                if (TryGetTargetForeground(out var ctx))
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    _queue.Enqueue(InputLogEntry.KeyboardDown(ctx, vkCode));
                }
            }

            return CallNextHookEx(_kbHook, nCode, wParam, lParam);
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                if (msg is WM_LBUTTONDOWN or WM_RBUTTONDOWN or WM_MBUTTONDOWN or WM_MOUSEWHEEL or WM_MOUSEHWHEEL)
                {
                    if (TryGetTargetForeground(out var ctx))
                    {
                        // MouseHookCallback 内の座標取得部分を修正
                        var info = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                        // info.pt はスクリーン座標（画面全体）
                        POINT pt = info.pt;

                        // ターゲットウィンドウの座標系に変換
                        // もし ScreenToClient が失敗しても、生の座標（pt.x, pt.y）を保持するようにする
                        bool ok = ScreenToClient(ctx.Hwnd, ref pt);

                        // 失敗時はスクリーン座標を代入してみる
                        int? cx = ok ? pt.x : info.pt.x;
                        int? cy = ok ? pt.y : info.pt.y;

                        int? delta = null;
                        if (msg == WM_MOUSEWHEEL || msg == WM_MOUSEHWHEEL)
                        {
                            // mouseData 上位16bit（signed）
                            delta = (short)((info.mouseData >> 16) & 0xffff);
                        }

                        _queue.Enqueue(InputLogEntry.MouseDown(
                            ctx,
                            operation: MapMouseOperation(msg),
                            x: cx,
                            y: cy,
                            delta: delta));
                    }
                }
            }

            return CallNextHookEx(_msHook, nCode, wParam, lParam);
        }

        private static string MapMouseOperation(int msg) => msg switch
        {
            WM_LBUTTONDOWN => "mouse_l",
            WM_RBUTTONDOWN => "mouse_r",
            WM_MBUTTONDOWN => "mouse_m",
            WM_MOUSEWHEEL => "wheel_v",
            WM_MOUSEHWHEEL => "wheel_h",
            _ => "mouse"
        };

        // ===== foreground filtering =====
        private bool TryGetTargetForeground(out ForegroundContext ctx)
        {
            ctx = default;
            if (TargetPid is null) return false;

            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return false;

            GetWindowThreadProcessId(hwnd, out uint pid);

            // 自分自身のPIDを取得して比較する
            uint myPid = (uint)Environment.ProcessId;
            if (pid == myPid) return false;
            if ((int)pid != TargetPid.Value) return false;

            ctx = new ForegroundContext(hwnd, (int)pid);
            return true;
        }

        public readonly record struct ForegroundContext(IntPtr Hwnd, int Pid);

        public readonly record struct InputLogEntry(
            DateTime TimestampUtc,
            int TargetPid,
            string Operation,
            int? X,
            int? Y,
            int? Delta,
            int? VirtualKey)
        {
            public static InputLogEntry KeyboardDown(ForegroundContext ctx, int vk)
                => new(DateTime.UtcNow, ctx.Pid, "keyboard", null, null, null, vk);

            public static InputLogEntry MouseDown(ForegroundContext ctx, string operation, int? x, int? y, int? delta)
                => new(DateTime.UtcNow, ctx.Pid, operation, x, y, delta, null);
        }

        // ===== Win32 =====
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x, y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string? lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);
    }
}
