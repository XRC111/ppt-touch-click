// PptTouchClick - 让 PowerPoint(Win10) 放映时"触控单击 = 翻页"的放映代理
//
// 编译（无需装 SDK，用 .NET Framework 4.8 自带编译器；GitHub Actions 上也是这么编）：
//   %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:winexe /optimize+ ^
//       /win32manifest:app.manifest /out:PptTouchClick.exe *.cs
//
// app.manifest 声明 PerMonitorV2 DPI 感知：否则在 125%/150% 缩放的屏幕上，
// 钩子拿到的是物理像素、窗口矩形是虚拟化坐标，两者对不上，空白/内容判定会整体错位。
//
// 参数：
//   --mouse-only         不做触摸注入，只把"有效单击"换成 PageDown（仍过滤滑动/长按）
//   --install-autostart  写入 HKCU\...\Run 开机启动（安装脚本用）
//   --remove-autostart   移除开机启动（卸载脚本用）

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PptTouchClick
{
    internal static class Program
    {
        public static bool MouseOnly;
        public static bool Verbose;
        public static string LogPath;

        [STAThread]
        private static int Main(string[] args)
        {
            // ---- 安装/卸载用的一次性命令（不启动主程序）----
            if (args.Any(a => a.Equals("--install-autostart", StringComparison.OrdinalIgnoreCase)))
                return Tray.SetAutostart(true, Application.ExecutablePath);
            if (args.Any(a => a.Equals("--remove-autostart", StringComparison.OrdinalIgnoreCase)))
                return Tray.SetAutostart(false, Application.ExecutablePath);

            if (args.Any(a => a.Equals("--mouse-only", StringComparison.OrdinalIgnoreCase))) MouseOnly = true;
            if (args.Any(a => a.Equals("--verbose", StringComparison.OrdinalIgnoreCase))) Verbose = true;

            try { LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PptTouchClick.log"); } catch { LogPath = null; }
            try
            {
                if (LogPath != null && File.Exists(LogPath) && new FileInfo(LogPath).Length > 2 * 1024 * 1024)
                    File.Delete(LogPath);
            }
            catch { }

            bool createdNew;
            using (var mutex = new Mutex(true, @"Local\PptTouchClick.SingleInstance", out createdNew))
            {
                if (!createdNew) return 0;   // 已有实例在跑

                Native.EnableDpiAwareness();   // 兜底：manifest 缺失时也要对齐物理像素
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                AppDomain.CurrentDomain.UnhandledException += (s, e) => Log("fatal: " + e.ExceptionObject);
                Application.ThreadException += (s, e) => Log("ui-exception: " + e.Exception);

                try
                {
                    using (var daemon = new DaemonForm())
                        Application.Run(daemon);
                }
                catch (Exception ex) { Log("fatal: " + ex); }
                finally { Log("stopped"); }
            }
            return 0;
        }

        /// <summary>后台写日志；钩子回调里永远不会碰文件。</summary>
        public static void Log(string msg)
        {
            string path = LogPath;
            if (string.IsNullOrEmpty(path)) return;
            string line = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg + Environment.NewLine;
            ThreadPool.QueueUserWorkItem(_ => { try { File.AppendAllText(path, line); } catch { } });
        }
    }

    internal sealed class RawEvent
    {
        public enum Kind { Click }
        public Kind EventKind;
        public DateTime Time;
        public Point Position;
        public uint Button;      // 1=左 2=右 4=中
        public string Source;    // mouse / touch
    }

    internal struct DebounceKey
    {
        public IntPtr Hwnd;
        public int Slide;      // -1 = 未知
        public uint Button;
        public DebounceKey(IntPtr h, int slide, uint button) { Hwnd = h; Slide = slide; Button = button; }
    }

    /// <summary>一次翻页后留下的"这一小块别再管"区域。</summary>
    internal struct Mark
    {
        public Rectangle Rect;
        public int Budget;       // 还能"吃掉"几次点击（一次触摸常伴随 mouse + pointer 两路事件）
        public DateTime Until;
    }

    /// <summary>无界面的消息泵窗口：仅用于接收 WH_MOUSE_LL / WH_POINTER / WH_KEYBOARD_LL 回调。</summary>
    internal sealed class DaemonForm : Form
    {
        private readonly object _qLock = new object();
        private readonly Queue<RawEvent> _queue = new Queue<RawEvent>();
        private readonly AutoResetEvent _signal = new AutoResetEvent(false);
        private readonly System.Collections.Concurrent.ConcurrentDictionary<int, TouchState> _touches =
            new System.Collections.Concurrent.ConcurrentDictionary<int, TouchState>();
        private readonly List<Mark> _marks = new List<Mark>();
        private readonly Dictionary<DebounceKey, DateTime> _debounce = new Dictionary<DebounceKey, DateTime>();

        private CancellationTokenSource _cts;
        private Thread _worker;
        private IntPtr _mouseHook, _touchHook, _kbHook;
        private HookProc _mouseProc, _touchProc, _kbProc;   // 必须持有委托，否则会被 GC
        private Tray _tray;

        private SlideshowState _state = new SlideshowState();
        private IntPtr _lastHwnd = IntPtr.Zero;
        private DateTime _lastRefresh = DateTime.MinValue;
        private DateTime _keyboardUntil = DateTime.MinValue;
        private DateTime _lastTrim = DateTime.MinValue;
        private bool _clickRunning;
        private Point _mouseDownPos;
        private bool _mouseDownSeen;

        /// <summary>托盘菜单里的"暂停翻页代理"。</summary>
        internal bool Enabled_ = true;

        // 判定参数（HKCU\Software\PptTouchClick 下可覆盖）
        internal int MoveTol = 26;          // 触摸抬起时允许的最大位移(px)，超过 = 滑动
        internal int TapMs = 1500;          // 按住超过这么久不算单击
        internal int DebounceMs = 800;      // 同窗口/同页/同键位的最短翻页间隔
        internal int MarkMs = 1200;         // 翻页后屏蔽该点周围的时长
        internal int AdvanceDelayMs = 160;  // 让 PPT 自己翻页后，等多久确认

        internal int SwallowCount, AdvancedCount;

        public DaemonForm()
        {
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-4000, -4000);
            Size = new Size(1, 1);
            Opacity = 0;
            Text = "PptTouchClick";
            LoadSettings();
        }

        internal void LoadSettings()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\PptTouchClick"))
                {
                    if (k == null) return;
                    MoveTol = GetInt(k, "MoveTol", MoveTol);
                    TapMs = GetInt(k, "TapMs", TapMs);
                    DebounceMs = GetInt(k, "DebounceMs", DebounceMs);
                    MarkMs = GetInt(k, "MarkMs", MarkMs);
                    AdvanceDelayMs = GetInt(k, "AdvanceDelayMs", AdvanceDelayMs);
                }
            }
            catch { }
        }

        private static int GetInt(Microsoft.Win32.RegistryKey k, string name, int d)
        {
            try { var o = k.GetValue(name); return o == null ? d : Convert.ToInt32(o); } catch { return d; }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Native.SetWindowPos(Handle, IntPtr.Zero, -4000, -4000, 1, 1, Native.SWP_NOZORDER | Native.SWP_NOACTIVATE);
            Program.Log("start pid=" + Process.GetCurrentProcess().Id + " mouseOnly=" + Program.MouseOnly);

            _tray = new Tray(this);
            _cts = new CancellationTokenSource();
            _worker = new Thread(WorkerLoop) { IsBackground = true, Name = "click-engine" };
            _worker.Start();

            _mouseProc = MouseHookProc;
            _mouseHook = Native.SetWindowsHookEx(Native.WH_MOUSE_LL, _mouseProc, Native.GetModuleHandle(null), 0);
            _kbProc = KeyboardHookProc;
            _kbHook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, _kbProc, Native.GetModuleHandle(null), 0);

            bool penTouch = SafePenTouch();
            if (penTouch)
            {
                _touchProc = TouchHookProc;
                _touchHook = Native.SetWindowsHookEx(Native.WH_POINTER, _touchProc, Native.GetModuleHandle(null), 0);
            }
            Program.Log("hooks: mouse=" + (_mouseHook != IntPtr.Zero)
                        + " keyboard=" + (_kbHook != IntPtr.Zero)
                        + " touch=" + (_touchHook != IntPtr.Zero)
                        + " penTouchApi=" + penTouch
                        + " touchInject=" + TouchInjector.Available);

            SlideSource.Start();
            RefreshState(true);
            if (_tray != null) _tray.Refresh(true);
        }

        private static bool SafePenTouch()
        {
            try { return Native.IsPenTouchAvailable(); } catch { return false; }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); return; }
            try { if (_cts != null) _cts.Cancel(); } catch { }
            Unhook();
            base.OnFormClosing(e);
        }

        private void Unhook()
        {
            if (_mouseHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
            if (_touchHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(_touchHook); _touchHook = IntPtr.Zero; }
            if (_kbHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(_kbHook); _kbHook = IntPtr.Zero; }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _tray != null) { _tray.Dispose(); _tray = null; }
            base.Dispose(disposing);
        }

        internal void ExitApp()
        {
            Program.Log("exit requested");
            if (_tray != null) _tray.AllowClose();
            Unhook();
            try { Application.Exit(); } catch { }
        }

        // ------------------------------------------------------------------ 鼠标钩子
        private IntPtr MouseHookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0) return Native.CallNextHookEx(_mouseHook, code, wParam, lParam);
            try
            {
                var ms = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                if (ms.dwExtraInfo == Native.INJECT_TAG) return Native.CallNextHookEx(_mouseHook, code, wParam, lParam);

                int msg = wParam.ToInt32();
                if (msg == Native.WM_LBUTTONDOWN)
                {
                    _mouseDownPos = new Point(ms.ptX, ms.ptY);
                    _mouseDownSeen = true;
                }
                else if (msg == Native.WM_LBUTTONUP)
                {
                    // 只认左键：右键/中键（例如放映时右键开菜单）不该触发翻页
                    if (_mouseDownSeen) EnqueueClick(_mouseDownPos, 1, "mouse");
                    _mouseDownSeen = false;
                }
            }
            catch (Exception ex) { Program.Log("mouse hook: " + ex.Message); }
            return Native.CallNextHookEx(_mouseHook, code, wParam, lParam);
        }

        // ------------------------------------------------------------------ 触摸钩子
        /// <summary>只把"单击"交给判定；滑动和长按直接吞掉（避免被当成换页手势，也避免误缩放）。</summary>
        private IntPtr TouchHookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0) return Native.CallNextHookEx(_touchHook, code, wParam, lParam);
            try
            {
                var info = (POINTER_INFO)Marshal.PtrToStructure(lParam, typeof(POINTER_INFO));
                if (info.pointerType != Native.PT_TOUCH)
                    return Native.CallNextHookEx(_touchHook, code, wParam, lParam);

                // 只在放映时接管触摸：否则会把全系统的滑动/缩放/长按一起吞掉
                var cur = _state;
                if (!Enabled_ || cur == null || !cur.Fullscreen)
                    return Native.CallNextHookEx(_touchHook, code, wParam, lParam);

                int id = (int)info.pointerId;
                uint f = info.pointerFlags;
                bool down = (f & Native.POINTER_FLAG_DOWN) != 0;
                bool up = (f & Native.POINTER_FLAG_UP) != 0;
                bool cancel = (f & Native.POINTER_FLAG_CANCELED) != 0;
                var pt = new Point(info.ptPixelLocation.X, info.ptPixelLocation.Y);
                IntPtr swallow = new IntPtr(1);

                if (down)
                {
                    if (_touches.Count > 24) SweepTouches();
                    _touches[id] = new TouchState { Tick = Environment.TickCount, Origin = pt };
                }
                else if (cancel)
                {
                    TouchState dropped;
                    _touches.TryRemove(id, out dropped);
                    return swallow;
                }
                else
                {
                    TouchState s;
                    if (_touches.TryGetValue(id, out s))
                    {
                        int dx = pt.X - s.Origin.X, dy = pt.Y - s.Origin.Y;
                        if (Math.Sqrt((double)dx * dx + (double)dy * dy) > MoveTol)
                        {
                            s.Swallow = true;   // 确定是滑动：这个触点之后的所有事件都吃掉
                            _touches[id] = s;
                            return swallow;
                        }
                        if (s.Swallow)
                        {
                            // 滑动的抬起也要吃掉，并把触点从表里移除，避免状态越积越多
                            if (up)
                            {
                                TouchState dropped0;
                                _touches.TryRemove(id, out dropped0);
                            }
                            return swallow;
                        }
                        if (up)
                        {
                            long held = unchecked(Environment.TickCount - s.Tick);
                            TouchState dropped;
                            _touches.TryRemove(id, out dropped);
                            if (held < 0 || held < TapMs) EnqueueClick(pt, 1, "touch");
                            else { Interlocked.Increment(ref SwallowCount); return swallow; }
                        }
                    }
                }
            }
            catch (Exception ex) { Program.Log("touch hook: " + ex.Message); }
            return Native.CallNextHookEx(_touchHook, code, wParam, lParam);
        }

        private void SweepTouches()
        {
            int now = Environment.TickCount;
            foreach (var kv in _touches.ToArray())
            {
                if (unchecked(now - kv.Value.Tick) > 150000) { TouchState dropped; _touches.TryRemove(kv.Key, out dropped); }
            }
        }

        private struct TouchState
        {
            public int Tick;
            public Point Origin;
            public bool Swallow;
        }

        // ------------------------------------------------------------------ 键盘钩子
        private IntPtr KeyboardHookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0) return Native.CallNextHookEx(_kbHook, code, wParam, lParam);
            try
            {
                int msg = wParam.ToInt32();
                if (msg == Native.WM_KEYDOWN || msg == Native.WM_SYSKEYDOWN)
                {
                    var kb = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                    int vk = (int)kb.vkCode;
                    if (kb.dwExtraInfo != Native.INJECT_TAG &&
                        (vk == Native.VK_NEXT || vk == Native.VK_PRIOR || vk == Native.VK_SPACE || vk == Native.VK_RETURN ||
                         vk == Native.VK_LEFT || vk == Native.VK_RIGHT || vk == Native.VK_UP || vk == Native.VK_DOWN ||
                         vk == Native.VK_BACK))
                        _keyboardUntil = DateTime.Now.AddMilliseconds(1200);   // 键盘/翻页笔正在翻页，1.2 秒内不抢
                }
            }
            catch { }
            return Native.CallNextHookEx(_kbHook, code, wParam, lParam);
        }

        // ------------------------------------------------------------------ 队列 / 引擎
        private void EnqueueClick(Point p, uint button, string source)
        {
            lock (_qLock)
                _queue.Enqueue(new RawEvent { EventKind = RawEvent.Kind.Click, Time = DateTime.Now, Position = p, Button = button, Source = source });
            _signal.Set();
        }

        private void WorkerLoop()
        {
            var batch = new List<RawEvent>(16);
            while (_cts != null && !_cts.IsCancellationRequested)
            {
                _signal.WaitOne(80);
                lock (_qLock) { while (_queue.Count > 0) batch.Add(_queue.Dequeue()); }
                if (batch.Count > 0)
                {
                    foreach (var ev in batch)
                    {
                        try { ProcessClick(ev); }
                        catch (Exception ex) { Program.Log("process click: " + ex.Message); }
                    }
                    batch.Clear();
                }
                if ((DateTime.Now - _lastRefresh).TotalMilliseconds > 300)
                {
                    _lastRefresh = DateTime.Now;
                    try { RefreshState(false); } catch (Exception ex) { Program.Log("refresh: " + ex.Message); }
                }
            }
        }

        internal SlideshowState State { get { return _state; } }
        internal DateTime KeyboardGraceUntil { get { return _keyboardUntil; } }

        /// <summary>找出正在放映的窗口、读出当前页；换页才重扫这一页的形状。</summary>
        internal void RefreshState(bool force)
        {
            IntPtr h = Slideshow.FindSlideshowWindow();
            if (h != IntPtr.Zero)
            {
                // 必须复用同一个状态对象：否则每次刷新都是"空状态"，
                // 页码缓存永远命中不了，会变成每 300ms 全屏重扫一次。
                var st = _state;
                bool switched = _lastHwnd != h;
                if (switched)
                {
                    st.Handle = h;
                    st.Fullscreen = true;
                    st.Enabled = true;
                    st.ContentRect = Rectangle.Empty;
                    st.PageShapes = null;
                    st.PageShapesTime = DateTime.MinValue;
                    st.Key = SlideKey.None;
                }
                st.Handle = h;
                st.Fullscreen = true;
                st.Enabled = true;

                bool changed = SlideModel.Update(h, st);
                if (switched || changed || force)
                    Program.Log("presenting hwnd=0x" + h.ToString("X")
                                + " slide=" + st.Key
                                + " content=" + st.ContentRect
                                + " shapes=" + (st.PageShapes == null ? "n/a" : st.PageShapes.Count.ToString()));
                if (switched || changed) ClearTransient();
            }
            else if (_lastHwnd != IntPtr.Zero || force)
            {
                _state = new SlideshowState();
                Program.Log("not presenting");
                ClearTransient();
            }
            _lastHwnd = h;
            MaybeTrimLog();
            if (_tray != null) _tray.Refresh(false);
        }

        /// <summary>长时间放映时日志会持续增长：每 10 分钟检查一次，超过 2MB 就截断重开。</summary>
        private void MaybeTrimLog()
        {
            if ((DateTime.Now - _lastTrim).TotalMinutes < 10) return;
            _lastTrim = DateTime.Now;
            try
            {
                string path = Program.LogPath;
                if (path != null && File.Exists(path) && new FileInfo(path).Length > 2 * 1024 * 1024)
                    File.WriteAllText(path, string.Empty);
            }
            catch { }
        }

        private void ClearTransient()
        {
            lock (_qLock) { _marks.Clear(); _debounce.Clear(); }
        }

        private void ProcessClick(RawEvent ev)
        {
            var st = _state;
            if (ev.Button != 1) return;                                            // 只处理左键/单指点击
            if (!Enabled_ || !st.Fullscreen || !st.Enabled) return;
            if (ev.Time < _keyboardUntil) return;
            if (_clickRunning) return;

            Point p = ev.Position;
            if (!st.ContentRect.IsEmpty && !st.ContentRect.Contains(p)) return;   // 黑边 / 放映控制条
            if (IsMarked(p)) return;                                               // 刚处理过的位置
            if (st.PageShapes != null && SlideModel.PointInsideShape(st.PageShapes, p)) return; // 文字/图片/按钮上

            // 只有知道当前页码时才做"同页去抖"：页码未知时按页去抖会把快速连点吃掉
            int slide = st.Key.Current ?? -1;
            if (slide >= 0)
            {
                var key = new DebounceKey(st.Handle, slide, ev.Button);
                var tnow = DateTime.Now;
                DateTime last;
                if (_debounce.TryGetValue(key, out last) && (tnow - last).TotalMilliseconds < DebounceMs) return;
                _debounce[key] = tnow;
                if (_debounce.Count > 64) _debounce.Clear();
            }

            _clickRunning = true;
            Mark(p, 2);
            try
            {
                var before = st.Key;
                if (!Program.MouseOnly && TouchInjector.Available)
                {
                    // 先让 PowerPoint 自己按"触摸点击"处理（触摸模式里它就是前进）
                    if (TouchInjector.Tap(st.Handle, p))
                    {
                        if (AdvanceDelayMs > 0) Thread.Sleep(AdvanceDelayMs);
                        RefreshState(false);
                        if (_state.Key.Current != before.Current || _state.Key.Count != before.Count)
                        {
                            Interlocked.Increment(ref AdvancedCount);
                            Program.Log("advance(touch) " + before + " -> " + _state.Key + " at " + p);
                            return;
                        }
                    }
                }
                // 兜底：直接发 PageDown（放映中的"下一页"）
                Native.SendInputKey(Native.VK_NEXT);
                Interlocked.Increment(ref AdvancedCount);
                Program.Log("advance(pagedown) " + before + " at " + p + " src=" + ev.Source);
            }
            finally { _clickRunning = false; }
        }

        internal void Mark(Point p, int budget)
        {
            var area = new Rectangle(p.X - 16, p.Y - 16, 32, 32);
            lock (_qLock) _marks.Add(new Mark { Rect = area, Budget = budget, Until = DateTime.Now.AddMilliseconds(MarkMs) });
        }

        private bool IsMarked(Point p)
        {
            lock (_qLock)
            {
                if (_marks.Count == 0) return false;
                var now = DateTime.Now;
                for (int i = _marks.Count - 1; i >= 0; i--)
                {
                    var m = _marks[i];
                    if (m.Until <= now || m.Budget <= 0) { _marks.RemoveAt(i); continue; }
                    if (m.Rect.Contains(p))
                    {
                        m.Budget--;
                        _marks[i] = m;
                        return true;
                    }
                }
                return false;
            }
        }

        internal int Advanced { get { return AdvancedCount; } }
        internal int Swallowed { get { return SwallowCount; } }
        internal bool TouchInjectOk { get { return !Program.MouseOnly && TouchInjector.Available; } }
        internal bool SlideSourceOk { get { return SlideSource.StartedOk; } }
        internal bool SlideSourceFailed { get { return SlideSource.Failed; } }
    }
}
