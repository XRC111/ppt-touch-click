using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace PptTouchClick
{
    // --------------------------------------------------------------------------
    // Win32 互操作
    // --------------------------------------------------------------------------
    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MOUSEINPUT
    {
        public int dx, dy;
        public uint mouseData, dwFlags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct KEYBDINPUT
    {
        public ushort wVk, wScan;
        public uint dwFlags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct INPUT
    {
        [FieldOffset(0)] public uint type;
        [FieldOffset(8)] public MOUSEINPUT mi;
        [FieldOffset(8)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSLLHOOKSTRUCT
    {
        public int ptX, ptY;
        public uint mouseData, flags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct KBDLLHOOKSTRUCT
    {
        public uint vkCode, scanCode, flags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINTER_INFO
    {
        public uint pointerType, pointerId, frameId, pointerFlags;
        public IntPtr sourceDevice, hwndTarget;
        public POINT ptPixelLocation;
        public POINT ptHimetricLocation;
        public POINT ptPixelLocationRaw;
        public POINT ptHimetricLocationRaw;
        public uint dwTime, historyCount, inputState, dwKeyStates;
        public ulong PerformanceCount;
        public int ButtonChangeType;
    }

    internal delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    /// <summary>页码来源：PptSlideProc.ocx（随插件安装的 COM 组件，读 PowerPoint 的"最近访问幻灯片"记录）。</summary>
    internal static class SlideSource
    {
        private static readonly object _lock = new object();
        private static bool _started;
        private static volatile bool _available;
        private static Thread _thread;

        internal static void Start()
        {
            lock (_lock)
            {
                if (_started) return;
                _started = true;
                try
                {
                    _thread = new Thread(Run) { IsBackground = true, Name = "slide-source" };
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                catch (Exception ex) { Program.Log("slide-source start failed: " + ex.Message); }
            }
        }

        internal static bool StartedOk { get { lock (_lock) return _available; } }
        /// <summary>组件不可用时：换页不再提前知道页码，只能按页变化"事后"重扫（精度略降，不影响点击判定）。</summary>
        internal static bool Failed { get { lock (_lock) return _started && !_available; } }

        private static void Run()
        {
            object watcher = null;
            try
            {
                var type = Type.GetTypeFromProgID("PptTouchClick.SlideWatcher", true);
                if (type == null) throw new InvalidOperationException("PptSlideProc 未注册");
                watcher = Activator.CreateInstance(type);
                object ok = type.InvokeMember("Start", System.Reflection.BindingFlags.InvokeMethod, null, watcher, null);
                if (!(ok is bool) || !(bool)ok) throw new InvalidOperationException("Start() 失败");
                _available = true;
                object status = type.InvokeMember("GetStatus", System.Reflection.BindingFlags.InvokeMethod, null, watcher, null);
                Program.Log("slide-source OK: " + status);

                // 关键：WH_POINTER 是低层钩子，回调由"安装它的线程"的消息循环派发。
                // 之前这里只 Sleep，钩子虽然装上了却永远收不到事件（超时还会被系统摘掉）。
                // 用 ApplicationContext 起一个最简消息泵，保住这个 COM 对象和钩子所在的线程。
                Application.Run(new ApplicationContext());
            }
            catch (Exception ex)
            {
                _available = false;
                Program.Log("slide-source unavailable: " + ex.Message);
            }
            finally
            {
                if (watcher != null)
                {
                    try { Marshal.FinalReleaseComObject(watcher); } catch { }
                }
            }
        }
    }

    internal static class Native
    {
        internal const uint MOUSE_MARKER = 0x5A;
        internal static readonly IntPtr INJECT_TAG = new IntPtr(MOUSE_MARKER);   // 注入事件的 dwExtraInfo 标记

        internal const int WH_MOUSE_LL = 14;
        internal const int WH_KEYBOARD_LL = 13;
        internal const int WH_POINTER = 24;

        internal const int WM_LBUTTONDOWN = 0x0201;
        internal const int WM_LBUTTONUP = 0x0202;
        internal const int WM_RBUTTONDOWN = 0x0204;
        internal const int WM_RBUTTONUP = 0x0205;
        internal const int WM_MBUTTONDOWN = 0x0207;
        internal const int WM_MBUTTONUP = 0x0208;
        internal const int WM_KEYDOWN = 0x0100;
        internal const int WM_KEYUP = 0x0101;
        internal const int WM_SYSKEYDOWN = 0x0104;
        internal const int WM_SYSKEYUP = 0x0105;

        internal const uint PT_TOUCH = 2;
        internal const uint POINTER_FLAG_NEW = 0x00000001;
        internal const uint POINTER_FLAG_INRANGE = 0x00000002;
        internal const uint POINTER_FLAG_INCONTACT = 0x00000004;
        internal const uint POINTER_FLAG_DOWN = 0x00010000;
        internal const uint POINTER_FLAG_UPDATE = 0x00020000;
        internal const uint POINTER_FLAG_UP = 0x00040000;
        internal const uint POINTER_FLAG_CANCELED = 0x80000000;

        internal const int INPUT_MOUSE = 0;
        internal const int INPUT_KEYBOARD = 1;
        internal const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        internal const uint MOUSEEVENTF_LEFTUP = 0x0004;
        internal const uint KEYEVENTF_KEYUP = 0x0002;

        internal const int VK_NEXT = 0x22;
        internal const int VK_PRIOR = 0x21;
        internal const int VK_SPACE = 0x20;
        internal const int VK_RETURN = 0x0D;
        internal const int VK_BACK = 0x08;
        internal const int VK_LEFT = 0x25;
        internal const int VK_RIGHT = 0x27;
        internal const int VK_UP = 0x26;
        internal const int VK_DOWN = 0x28;
        internal const int VK_ESCAPE = 0x1B;

        internal const uint SWP_NOZORDER = 0x0004;
        internal const uint SWP_NOACTIVATE = 0x0010;

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        internal static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")]
        internal static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint SendInput(uint n, INPUT[] inputs, int size);

        [DllImport("user32.dll")]
        internal static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        internal static extern bool IsPenTouchAvailable();

        [DllImport("shcore.dll")]
        internal static extern int SetProcessDpiAwareness(int value);   // 2 = PerMonitor

        [DllImport("user32.dll")]
        internal static extern bool SetProcessDPIAware();

        /// <summary>
        /// 声明 DPI 感知：不声明时钩子的物理像素坐标和窗口矩形（虚拟化坐标）对不上，
        /// 高缩放下"空白/内容"判定会整体错位。manifest 已声明时这里会失败，无害。
        /// </summary>
        internal static void EnableDpiAwareness()
        {
            try { if (SetProcessDpiAwareness(2) == 0) return; } catch { }
            try { SetProcessDPIAware(); } catch { }
        }

        internal static void SendInputKey(int vk)
        {
            // INPUT 的联合体在 32 位下偏移是 4、64 位下是 8；这里按 64 位布局声明，
            // 32 位进程（32 位 Windows）改用 keybd_event，否则 SendInput 会因 cbSize 不对而失败。
            if (IntPtr.Size != 8)
            {
                keybd_event((byte)vk, 0, 0, new UIntPtr((ulong)MOUSE_MARKER));
                keybd_event((byte)vk, 0, KEYEVENTF_KEYUP, new UIntPtr((ulong)MOUSE_MARKER));
                return;
            }
            var down = new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = (ushort)vk, dwExtraInfo = INJECT_TAG } };
            var up = new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = (ushort)vk, dwFlags = KEYEVENTF_KEYUP, dwExtraInfo = INJECT_TAG } };
            var arr = new[] { down, up };
            SendInput(2, arr, Marshal.SizeOf(typeof(INPUT)));
        }

        internal static void SendInputClick()
        {
            var down = new INPUT { type = INPUT_MOUSE, mi = new MOUSEINPUT { dwFlags = MOUSEEVENTF_LEFTDOWN, dwExtraInfo = INJECT_TAG } };
            var up = new INPUT { type = INPUT_MOUSE, mi = new MOUSEINPUT { dwFlags = MOUSEEVENTF_LEFTUP, dwExtraInfo = INJECT_TAG } };
            var arr = new[] { down, up };
            SendInput(2, arr, Marshal.SizeOf(typeof(INPUT)));
        }
    }
}
