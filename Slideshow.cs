using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PptTouchClick
{
    internal sealed class SlideshowState
    {
        public IntPtr Handle;
        public bool Fullscreen;
        public bool Enabled;
        public Rectangle ContentRect = Rectangle.Empty;   // 幻灯片显示区域（屏幕坐标）
        public SlideKey Key = SlideKey.None;              // 当前页（缓存键）
        public List<Rectangle> PageShapes;                // 当前页的形状包围盒（屏幕坐标）；null = 未知
        public DateTime PageShapesTime = DateTime.MinValue;
    }

    /// <summary>识别"哪一个窗口是 PowerPoint 的全屏放映窗口"。</summary>
    internal static class Slideshow
    {
        private const int PW_RENDERFULLCONTENT = 2;
        private const int MONITOR_DEFAULTTONEAREST = 2;

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr h, ref POINT p);
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr h, IntPtr hdc, int flags);
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr h, int flags);
        [DllImport("user32.dll")] private static extern bool GetMonitorInfo(IntPtr mon, ref MONITORINFO mi);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();

        private delegate bool EnumProc(IntPtr h, IntPtr lParam);
        private static IntPtr _cached = IntPtr.Zero;

        public static IntPtr ForegroundWindow
        {
            get { try { return GetForegroundWindow(); } catch { return IntPtr.Zero; } }
        }

        /// <summary>找到正在放映的黑色全屏窗口（类名 screenClass、无标题、铺满一块显示器）。</summary>
        public static IntPtr FindSlideshowWindow()
        {
            if (_cached != IntPtr.Zero && StillValid(_cached)) return _cached;

            IntPtr fg = ForegroundWindow;
            if (IsBlackWindow(fg) && FillsMonitor(fg)) { _cached = fg; return fg; }

            IntPtr found = IntPtr.Zero, fallback = IntPtr.Zero;
            EnumWindows((h, lp) =>
            {
                try
                {
                    if (IsBlackWindow(h) && FillsMonitor(h))
                    {
                        if (IsPowerPointWindow(h)) { found = h; return false; }
                        if (fallback == IntPtr.Zero) fallback = h;
                    }
                }
                catch { }
                return true;
            }, IntPtr.Zero);

            if (found == IntPtr.Zero) found = fallback;
            _cached = found;
            return found;
        }

        private static bool StillValid(IntPtr h)
        {
            return h != IntPtr.Zero && IsWindowVisible(h) && !IsIconic(h) && IsBlackWindow(h) && FillsMonitor(h);
        }

        private static bool IsBlackWindow(IntPtr h)
        {
            if (h == IntPtr.Zero || !IsWindowVisible(h) || IsIconic(h)) return false;
            var sb = new StringBuilder(64);
            GetClassName(h, sb, sb.Capacity);
            if (!sb.ToString().Equals("screenClass", StringComparison.OrdinalIgnoreCase)) return false;
            var sbT = new StringBuilder(8);
            Native.GetWindowText(h, sbT, sbT.Capacity);
            return sbT.Length == 0;            // 主放映窗口无标题；演示者视图有标题
        }

        private static bool FillsMonitor(IntPtr h)
        {
            RECT r;
            if (!GetWindowRect(h, out r)) return false;
            var mon = MonitorFromWindow(h, MONITOR_DEFAULTTONEAREST);
            var mi = new MONITORINFO(); mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
            if (!GetMonitorInfo(mon, ref mi)) return false;
            int w = r.Right - r.Left, hgt = r.Bottom - r.Top;
            int mw = mi.rcMonitor.Right - mi.rcMonitor.Left, mh = mi.rcMonitor.Bottom - mi.rcMonitor.Top;
            return w >= mw - 2 && hgt >= mh - 2;
        }

        public static bool IsPowerPointWindow(IntPtr h)
        {
            try
            {
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                using (var p = System.Diagnostics.Process.GetProcessById((int)pid))
                    return p.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase);
            }
            catch { return true; }
        }

        public static Rectangle ClientScreenRect(IntPtr h)
        {
            RECT c;
            if (!GetClientRect(h, out c)) return Screen.FromHandle(h).Bounds;
            var p = new POINT { X = c.Left, Y = c.Top };
            ClientToScreen(h, ref p);
            return new Rectangle(p.X, p.Y, c.Right - c.Left, c.Bottom - c.Top);
        }

        public static Rectangle MonitorBounds(IntPtr h)
        {
            try
            {
                var mon = MonitorFromWindow(h, MONITOR_DEFAULTTONEAREST);
                var mi = new MONITORINFO(); mi.cbSize = Marshal.SizeOf(typeof(MONITORINFO));
                if (GetMonitorInfo(mon, ref mi))
                    return Rectangle.FromLTRB(mi.rcMonitor.Left, mi.rcMonitor.Top, mi.rcMonitor.Right, mi.rcMonitor.Bottom);
            }
            catch { }
            try { return Screen.FromHandle(h).Bounds; } catch { return Rectangle.Empty; }
        }

        /// <summary>截取放映窗口当前画面（用 PrintWindow 离屏渲染，不需要窗口在前台）。</summary>
        public static Bitmap Capture(IntPtr h)
        {
            RECT r;
            if (!GetWindowRect(h, out r)) return null;
            int w = r.Right - r.Left, hgt = r.Bottom - r.Top;
            if (w <= 0 || hgt <= 0) return null;
            var bmp = new Bitmap(w, hgt, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                IntPtr hdc = g.GetHdc();
                bool ok = PrintWindow(h, hdc, PW_RENDERFULLCONTENT);
                g.ReleaseHdc(hdc);
                if (!ok) { bmp.Dispose(); return null; }
            }
            return bmp;
        }
    }

    /// <summary>当前页（缓存键）：来自 PptSlideProc 读出的"最近访问幻灯片"记录。</summary>
    internal struct SlideKey
    {
        public int? Current;   // 从 0 开始
        public int Count;      // 总页数；0 = 未知

        public static readonly SlideKey None = new SlideKey { Count = 0 };

        public bool Valid { get { return Count > 0; } }

        public override string ToString()
        {
            if (Count <= 0) return "?/?";
            int cur = Current.HasValue ? Current.Value + 1 : 0;
            return cur + "/" + Count;
        }

        public static SlideKey Read()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\PptTouchClick\Slide"))
                {
                    if (k == null) return None;
                    int n = Convert.ToInt32(k.GetValue("Count", 0));
                    int c = Convert.ToInt32(k.GetValue("Current", -1));
                    int age = Convert.ToInt32(k.GetValue("AgeMs", 0));
                    if (n <= 0 || age > 15000) return None;
                    return new SlideKey { Count = n, Current = (c >= 0 && c < n) ? (int?)c : null };
                }
            }
            catch { return None; }
        }
    }

    /// <summary>从画面里找出"幻灯片内容区"和当前页上的形状（文字 / 图片 / 图表 / 按钮）。</summary>
    internal static class SlideModel
    {
        /// <summary>更新状态里的内容区与当前页；换页/换窗口时重扫形状。返回是否有变化。</summary>
        internal static bool Update(IntPtr hwnd, SlideshowState st)
        {
            var oldKey = st.Key;
            var oldRect = st.ContentRect;
            st.Key = SlideKey.Read();

            bool pageChanged = oldKey.Current != st.Key.Current || oldKey.Count != st.Key.Count;
            if (pageChanged) st.PageShapes = null;

            if (st.ContentRect.IsEmpty || st.PageShapes == null || pageChanged)
                LoadPageShapes(hwnd, st);

            return pageChanged || (st.ContentRect != oldRect);
        }

        /// <summary>扫描一页：确定内容区，得到形状包围盒（屏幕坐标）。一页只做一次。</summary>
        internal static void LoadPageShapes(IntPtr hwnd, SlideshowState st)
        {
            try
            {
                using (var bmp = Slideshow.Capture(hwnd))
                {
                    if (bmp == null) return;
                    RECT wr;
                    Native.GetWindowRect(hwnd, out wr);
                    var offset = new Point(wr.Left, wr.Top);

                    var content = ContentRect(bmp);
                    if (content.Width < 40 || content.Height < 40) return;
                    var contentScreen = new Rectangle(content.X + offset.X, content.Y + offset.Y, content.Width, content.Height);

                    var shapes = ScanShapes(bmp, content);
                    var screenShapes = new List<Rectangle>(shapes.Count);
                    foreach (var s in shapes)
                        screenShapes.Add(new Rectangle(s.X + offset.X, s.Y + offset.Y, s.Width, s.Height));

                    st.ContentRect = contentScreen;
                    st.PageShapes = screenShapes;
                    st.PageShapesTime = DateTime.Now;
                    Program.Log("page scanned " + st.Key + ": content=" + contentScreen + " shapes=" + screenShapes.Count);
                }
            }
            catch (Exception ex)
            {
                Program.Log("page scan failed: " + ex.Message);
            }
        }

        public static bool PointInsideShape(List<Rectangle> shapes, Point p)
        {
            foreach (var r in shapes) if (r.Contains(p)) return true;
            return false;
        }

        // ------------------------------------------------------------------ 内容区
        private static bool NotBlack(int c)
        {
            int r = (c >> 16) & 0xFF, g = (c >> 8) & 0xFF, b = c & 0xFF;
            return r + g + b > 72;
        }

        /// <summary>内容区 = 画面里非黑区域的外接矩形；左右两侧的"放映控制条"自动排除。</summary>
        internal static Rectangle ContentRect(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int strideI = data.Stride / 4;
                var buf = new int[strideI * h];
                Marshal.Copy(data.Scan0, buf, 0, buf.Length);

                int minX = int.MaxValue, minY = int.MaxValue, maxX = -1, maxY = -1;
                for (int y = 0; y < h; y++)
                {
                    int row = y * strideI;
                    for (int x = 0; x < w; x++)
                    {
                        if (NotBlack(buf[row + x]))
                        {
                            if (x < minX) minX = x;
                            if (x > maxX) maxX = x;
                            if (y < minY) minY = y;
                            if (y > maxY) maxY = y;
                        }
                    }
                }
                if (maxX < 0) return Rectangle.Empty;

                int left = minX, top = minY, right = maxX, bottom = maxY;

                if (left > 0 && left < w / 10)
                {
                    int inner = FindContentColumn(buf, strideI, left, right, top, bottom, +1);
                    if (inner > left) left = inner;
                }
                if (right < w - 1 && w - 1 - right < w / 10)
                {
                    int inner = FindContentColumn(buf, strideI, right, left, top, bottom, -1);
                    if (inner >= 0 && inner < right) right = inner;
                }
                return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
            }
            finally { bmp.UnlockBits(data); }
        }

        /// <summary>从 edge 往 dir 找第一条"有内容"的列（≥60% 像素非黑）。</summary>
        private static int FindContentColumn(int[] buf, int strideI, int edge, int limit, int top, int bottom, int dir)
        {
            int x = edge;
            int stop = Math.Max(edge, limit);
            while (x >= 0 && x <= stop)
            {
                int bright = 0, total = 0;
                for (int y = top; y <= bottom; y += 2)
                {
                    if (NotBlack(buf[y * strideI + x])) bright++;
                    total++;
                }
                if (total > 0 && bright * 100 / total >= 60) return x;
                x += dir;
            }
            return -1;
        }

        // ------------------------------------------------------------------ 形状
        private sealed class Comp
        {
            public int X0, Y0, X1, Y1, Count;
        }

        private struct RectI
        {
            public int L, T, R, B;
            public RectI(int l, int t, int r, int b) { L = l; T = t; R = r; B = b; }
            public int W { get { return R - L; } }
            public int H { get { return B - T; } }
            public Rectangle ToRectangle() { return Rectangle.FromLTRB(L, T, R, B); }
            public bool Intersects(RectI o, int pad) { return L - pad < o.R && o.L - pad < R && T - pad < o.B && o.T - pad < B; }
            public RectI Union(RectI o) { return new RectI(Math.Min(L, o.L), Math.Min(T, o.T), Math.Max(R, o.R), Math.Max(B, o.B)); }
        }

        /// <summary>
        /// 局部对比度 &gt; 85 的像素视为"内容"：白底黑字、深底浅字、照片、图表都能框住；
        /// 再按小间距合并，得到每块文字/图形的粗略包围盒（画面像素坐标）。
        /// </summary>
        internal static List<Rectangle> ScanShapes(Bitmap bmp, Rectangle content)
        {
            var result = new List<Rectangle>();
            int step = 2;
            int w = bmp.Width, h = bmp.Height;
            var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int strideI = data.Stride / 4;
                var buf = new int[strideI * h];
                Marshal.Copy(data.Scan0, buf, 0, buf.Length);

                int gw = content.Width / step, gh = content.Height / step;
                if (gw < 8 || gh < 8) return result;
                var lum = new int[gw * gh];
                for (int gy = 0; gy < gh; gy++)
                {
                    int y = content.Y + gy * step;
                    for (int gx = 0; gx < gw; gx++)
                    {
                        int c = buf[y * strideI + content.X + gx * step];
                        int r = (c >> 16) & 0xFF, g = (c >> 8) & 0xFF, b = c & 0xFF;
                        lum[gy * gw + gx] = (r * 77 + g * 151 + b * 28) >> 8;
                    }
                }

                const int R = 2;
                var mask = new bool[lum.Length];
                for (int gy = 0; gy < gh; gy++)
                {
                    int y0 = Math.Max(0, gy - R), y1 = Math.Min(gh - 1, gy + R);
                    for (int gx = 0; gx < gw; gx++)
                    {
                        int x0 = Math.Max(0, gx - R), x1 = Math.Min(gw - 1, gx + R);
                        int mn = 255, mx = 0;
                        for (int yy = y0; yy <= y1; yy++)
                        {
                            int row = yy * gw;
                            for (int xx = x0; xx <= x1; xx++)
                            {
                                int v = lum[row + xx];
                                if (v < mn) mn = v;
                                if (v > mx) mx = v;
                            }
                        }
                        mask[gy * gw + gx] = (mx - mn) > 85;
                    }
                }

                var visited = new bool[mask.Length];
                var stack = new Stack<int>();
                var kept = new List<RectI>();
                for (int i = 0; i < mask.Length; i++)
                {
                    if (!mask[i] || visited[i]) continue;
                    var comp = new Comp { X0 = int.MaxValue, Y0 = int.MaxValue, X1 = -1, Y1 = -1 };
                    stack.Push(i);
                    visited[i] = true;
                    while (stack.Count > 0)
                    {
                        int idx = stack.Pop();
                        int cx = idx % gw, cy = idx / gw;
                        comp.Count++;
                        if (cx < comp.X0) comp.X0 = cx;
                        if (cy < comp.Y0) comp.Y0 = cy;
                        if (cx > comp.X1) comp.X1 = cx;
                        if (cy > comp.Y1) comp.Y1 = cy;
                        for (int dy = -1; dy <= 1; dy++)
                        {
                            int ny = cy + dy;
                            if (ny < 0 || ny >= gh) continue;
                            for (int dx = -1; dx <= 1; dx++)
                            {
                                int nx = cx + dx;
                                if (nx < 0 || nx >= gw) continue;
                                int ni = ny * gw + nx;
                                if (mask[ni] && !visited[ni]) { visited[ni] = true; stack.Push(ni); }
                            }
                        }
                    }
                    if (comp.Count >= 6)
                        kept.Add(new RectI(content.X + comp.X0 * step, content.Y + comp.Y0 * step,
                                           content.X + (comp.X1 + 1) * step, content.Y + (comp.Y1 + 1) * step));
                }

                foreach (var r in Merge(kept, 12))
                {
                    if (r.W < 6 || r.H < 6) continue;
                    result.Add(r.ToRectangle());
                }
            }
            finally { bmp.UnlockBits(data); }
            return result;
        }

        private static List<RectI> Merge(List<RectI> rects, int pad)
        {
            bool changed = true;
            var cur = rects;
            while (changed)
            {
                changed = false;
                var next = new List<RectI>();
                var used = new bool[cur.Count];
                for (int i = 0; i < cur.Count; i++)
                {
                    if (used[i]) continue;
                    var r = cur[i];
                    bool merged = true;
                    while (merged)
                    {
                        merged = false;
                        for (int j = i + 1; j < cur.Count; j++)
                        {
                            if (used[j]) continue;
                            if (r.Intersects(cur[j], pad))
                            {
                                r = r.Union(cur[j]);
                                used[j] = true;
                                merged = true;
                                changed = true;
                            }
                        }
                    }
                    next.Add(r);
                }
                cur = next;
            }
            return cur;
        }
    }
}
