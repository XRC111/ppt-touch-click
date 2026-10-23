using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace PptTouchClick
{
    /// <summary>
    /// 在放映窗口上注入一次真实触摸点击（Windows 8+ 的 InitializeTouchInjection / InjectTouchInput）。
    /// 好处：PPT 在触摸模式下会把"触摸点击"当作前进，和它自己的换页逻辑一致，也不依赖窗口是否在前台。
    /// 不可用时自动退回 PageDown。
    /// </summary>
    internal static class TouchInjector
    {
        private const uint TOUCH_FEEDBACK_DEFAULT = 0x1;
        private const uint TOUCH_FLAG_NONE = 0x00000000;
        private const uint PT_TOUCH = 0x00000002;
        private const uint POINTER_FLAG_NEW = 0x00000001;
        private const uint POINTER_FLAG_INRANGE = 0x00000002;
        private const uint POINTER_FLAG_INCONTACT = 0x00000004;
        private const uint POINTER_FLAG_DOWN = 0x00010000;
        private const uint POINTER_FLAG_UP = 0x00040000;
        private const uint TOUCH_MASK_CONTACTAREA = 0x1;
        private const uint TOUCH_MASK_PRESSURE = 0x2;
        private const uint MAX_TOUCHES = 10;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINTER_TOUCH_INFO
        {
            public POINTER_INFO PointerInfo;
            public uint TouchFlags, TouchMask;
            public RECT ContactArea;
            public RECT ContactAreaRaw;
            public uint Orientation;
            public int Pressure;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool InitializeTouchInjection(uint maxCount, uint feedbackMode);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool InjectTouchInput(uint count, POINTER_TOUCH_INFO[] contacts);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private static bool _initialized;
        private static bool _disabled;
        private static int _failures;
        // 注入的 pointerId 必须在 [0, MAX_TOUCHES) 范围内（InitializeTouchInjection 的约定），
        // 之前从 100 开始，InjectTouchInput 会以 ERROR_INVALID_PARAMETER 失败。
        private static uint _nextId;

        /// <summary>
        /// 本机是否支持触摸注入。
        ///
        /// 旧实现调用 user32!IsPenTouchAvailable —— 这个导出不存在，
        /// DllImport 抛 EntryPointNotFoundException 后被 catch 成 false，
        /// 于是 Available 恒为 false、触摸注入从未生效过（issue #3 的 touchInject=False）。
        /// 现在按真实条件判断：有触摸数字化仪，且 InitializeTouchInjection 成功。
        /// </summary>
        internal static bool Available
        {
            get
            {
                if (_disabled) return false;
                if (_initialized) return true;
                if (!Native.HasTouchDigitizer()) return false;
                return EnsureInit();
            }
        }

        private static bool EnsureInit()
        {
            if (_initialized) return true;
            try
            {
                if (!InitializeTouchInjection(MAX_TOUCHES, TOUCH_FEEDBACK_DEFAULT)) return false;
                _initialized = true;
                Program.Log("touch injection initialized");
                return true;
            }
            catch (Exception ex)
            {
                Program.Log("touch init exception: " + ex.Message);
                return false;
            }
        }

        /// <summary>在屏幕点处注入一次"点按"（按下 → 60ms → 抬起）。成功返回 true。</summary>
        internal static bool Tap(IntPtr hwnd, Point screenPoint)
        {
            if (!Available || hwnd == IntPtr.Zero) return false;
            if (!EnsureInit()) { _disabled = true; return false; }

            try
            {
                if (GetForegroundWindow() != hwnd) SetForegroundWindow(hwnd);

                uint id = _nextId;
                _nextId = (_nextId + 1) % MAX_TOUCHES;
                int x = screenPoint.X, y = screenPoint.Y;
                var area = new RECT { Left = x - 5, Top = y - 5, Right = x + 5, Bottom = y + 5 };

                var down = new POINTER_TOUCH_INFO
                {
                    PointerInfo = new POINTER_INFO
                    {
                        pointerType = PT_TOUCH,
                        pointerId = id,
                        pointerFlags = POINTER_FLAG_NEW | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT | POINTER_FLAG_DOWN,
                        ptPixelLocation = new POINT { X = x, Y = y },
                        ptPixelLocationRaw = new POINT { X = x, Y = y }
                    },
                    TouchFlags = TOUCH_FLAG_NONE,
                    TouchMask = TOUCH_MASK_CONTACTAREA | TOUCH_MASK_PRESSURE,
                    ContactArea = area,
                    Pressure = 32000
                };
                var up = down;
                up.PointerInfo.pointerFlags = POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT | POINTER_FLAG_UP;

                if (!InjectTouchInput(1, new[] { down })) { Fail("down"); return false; }
                System.Threading.Thread.Sleep(60);
                if (!InjectTouchInput(1, new[] { up })) { Fail("up"); return false; }
                _failures = 0;
                return true;
            }
            catch (Exception ex)
            {
                Program.Log("touch inject exception: " + ex.Message);
                return false;
            }
        }

        private static void Fail(string phase)
        {
            _failures++;
            int err = Marshal.GetLastWin32Error();
            Program.Log("touch inject failed (" + phase + "): win32=" + err);
            if (_failures >= 3)
            {
                _disabled = true;
                Program.Log("touch inject disabled; fallback to PageDown");
            }
        }
    }
}
