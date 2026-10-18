using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PptTouchClick
{
    /// <summary>托盘图标：显示状态、开关、打开日志、开机自启、退出。</summary>
    internal sealed class Tray : IDisposable
    {
        private readonly DaemonForm _daemon;
        private readonly NotifyIcon _icon;
        private readonly Icon _iconImage;
        private readonly ToolStripMenuItem _status;
        private readonly ToolStripMenuItem _toggle;
        private readonly ToolStripMenuItem _autostart;
        private bool _enabled = true;
        private string _lastText = "";
        private DateTime _lastUpdate = DateTime.MinValue;

        private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValue = "PptTouchClick";

        internal Tray(DaemonForm daemon)
        {
            _daemon = daemon;

            var menu = new ContextMenuStrip();
            _status = new ToolStripMenuItem("未在放映") { Enabled = false };
            menu.Items.Add(_status);
            menu.Items.Add(new ToolStripSeparator());

            _toggle = new ToolStripMenuItem("暂停翻页代理", null, (s, e) => { _enabled = !_enabled; _daemon.Enabled_ = _enabled; UpdateStatus(true); });
            menu.Items.Add(_toggle);
            menu.Items.Add(new ToolStripMenuItem("状态详情…", null, (s, e) => ShowDetails()));
            menu.Items.Add(new ToolStripMenuItem("打开日志", null, (s, e) => OpenLog()));
            menu.Items.Add(new ToolStripMenuItem("重新载入设置", null, (s, e) => { _daemon.LoadSettings(); ShowBalloon("已重新载入设置"); }));
            menu.Items.Add(new ToolStripSeparator());

            _autostart = new ToolStripMenuItem("开机自动启动", null, (s, e) => ToggleAutostart())
            {
                CheckOnClick = false,
                Checked = IsAutostartOn()
            };
            menu.Items.Add(_autostart);
            menu.Items.Add(new ToolStripMenuItem("退出", null, (s, e) => _daemon.ExitApp()));

            _iconImage = MakeIcon();   // 只创建一次：每次状态变化重建会泄漏 HICON
            _icon = new NotifyIcon
            {
                Icon = _iconImage,
                Text = "PPT 触控单击翻页",
                Visible = true,
                ContextMenuStrip = menu
            };
            _icon.DoubleClick += (s, e) => ShowDetails();
        }

        internal void AllowClose() { }

        internal void UpdateStatus(bool force)
        {
            if ((DateTime.Now - _lastUpdate).TotalMilliseconds < 250 && !force) return;
            _lastUpdate = DateTime.Now;
            var st = _daemon.State;
            string text;
            if (!_enabled) text = "已暂停";
            else if (st.Fullscreen) text = "放映中 " + st.Key;
            else text = "等待放映";
            if (text == _lastText && !force) return;
            _lastText = text;
            _status.Text = text;
            _toggle.Text = _enabled ? "暂停翻页代理" : "恢复翻页代理";
            _autostart.Checked = IsAutostartOn();
            try
            {
                _icon.Text = "PPT 触控单击翻页 - " + text;
            }
            catch { }
        }

        /// <summary>工作线程调用：切回 UI 线程再改控件，避免跨线程访问。</summary>
        internal void Refresh(bool force)
        {
            try
            {
                if (_daemon.IsHandleCreated && _daemon.InvokeRequired)
                    _daemon.BeginInvoke((Action)(() => UpdateStatus(force)));
                else
                    UpdateStatus(force);
            }
            catch { }
        }

        private void ShowDetails()
        {
            var st = _daemon.State;
            string lines =
                "状态：" + (_enabled ? "运行中" : "已暂停") + Environment.NewLine +
                "放映窗口：" + (st.Fullscreen ? "有 (0x" + st.Handle.ToString("X") + ")" : "无") + Environment.NewLine +
                "当前页：" + st.Key + Environment.NewLine +
                "内容区：" + (st.ContentRect.IsEmpty ? "未知" : st.ContentRect.ToString()) + Environment.NewLine +
                "本页形状：" + (st.PageShapes == null ? "未扫描" : st.PageShapes.Count + " 处") + Environment.NewLine +
                "触摸点击注入：" + (_daemon.TouchInjectOk ? "可用" : (Program.MouseOnly ? "已关闭(--mouse-only)" : "不可用→用 PageDown")) + Environment.NewLine +
                "页码来源(PptSlideProc)：" + (_daemon.SlideSourceOk ? "正常" : (_daemon.SlideSourceFailed ? "不可用（换页后重扫）" : "启动中")) + Environment.NewLine +
                "累计翻页：" + _daemon.Advanced + "，吞掉的滑动/长按：" + _daemon.Swallowed + Environment.NewLine +
                "日志：" + Program.LogPath;
            MessageBox.Show(lines, "PPT 触控单击翻页", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowBalloon(string s)
        {
            try { _icon.ShowBalloonTip(1500, "PPT 触控单击翻页", s, ToolTipIcon.Info); } catch { }
        }

        private static void OpenLog()
        {
            try
            {
                if (!File.Exists(Program.LogPath)) return;
                Process.Start("notepad.exe", "\"" + Program.LogPath + "\"");
            }
            catch { }
        }

        internal static bool IsAutostartOn()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey))
                    return k != null && k.GetValue(RunValue) != null;
            }
            catch { return false; }
        }

        /// <summary>由安装脚本使用：--install-autostart / --remove-autostart。</summary>
        internal static int SetAutostart(bool on, string exePath)
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKey))
                {
                    if (k == null) return 2;
                    if (on) k.SetValue(RunValue, "\"" + exePath + "\"");
                    else k.DeleteValue(RunValue, false);
                }
                return 0;
            }
            catch { return 1; }
        }

        private void ToggleAutostart()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKey))
                {
                    if (k == null) return;
                    if (IsAutostartOn()) k.DeleteValue(RunValue, false);
                    else k.SetValue(RunValue, "\"" + Application.ExecutablePath + "\"");
                }
                _autostart.Checked = IsAutostartOn();
                ShowBalloon(_autostart.Checked ? "已设置为开机自动启动" : "已取消开机启动");
            }
            catch (Exception ex) { ShowBalloon("设置失败：" + ex.Message); }
        }

        /// <summary>运行时画一个图标（不依赖外部 .ico 文件）。</summary>
        private static Icon MakeIcon()
        {
            using (var bmp = new Bitmap(32, 32))
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var b = new SolidBrush(Color.FromArgb(255, 194, 59)))
                    g.FillRectangle(b, 1, 1, 30, 30);
                using (var pen = new Pen(Color.White, 3))
                    g.DrawPolygon(pen, new[] { new PointF(13, 8), new PointF(24, 16), new PointF(13, 24) });
                IntPtr h = bmp.GetHicon();
                return Icon.FromHandle(h);   // 进程生命周期内不释放，避免托盘图标失效
            }
        }

        public void Dispose()
        {
            try { _icon.Visible = false; _icon.Dispose(); } catch { }
            try { _iconImage.Dispose(); } catch { }
        }
    }
}
