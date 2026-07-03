using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelCommonTools.Core;

namespace ExcelCommonTools.Services
{
    /// <summary>
    /// 聚光灯服务：通过半透明覆盖层高亮选中单元格所在的行和列。
    /// 不修改工作簿任何内容，纯视觉效果。
    /// </summary>
    public class SpotlightService
    {
        private readonly Excel.Application _app;
        private bool _isEnabled;
        private SpotlightOverlay _overlay;
        private Timer _refreshTimer;
        private Bitmap _currentBmp;
        private double _dpiScale;
        private double _zoom;

        private static readonly Color HighlightColor = Color.FromArgb(60, 0xA1, 0xD6, 0x68);

        public SpotlightService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public bool IsEnabled => _isEnabled;

        public void Enable()
        {
            _isEnabled = true;
            if (_overlay == null)
                _overlay = new SpotlightOverlay();
            _overlay.Show();
            _app.SheetSelectionChange += OnSelectionChange;
            _refreshTimer = new Timer { Interval = 100 };
            _refreshTimer.Tick += OnRefreshTick;
            _refreshTimer.Start();
            UpdateOverlay();
        }

        public void Disable()
        {
            _isEnabled = false;
            _app.SheetSelectionChange -= OnSelectionChange;
            if (_refreshTimer != null)
            {
                _refreshTimer.Stop();
                _refreshTimer.Dispose();
                _refreshTimer = null;
            }
            DisposeBitmap();
            _overlay?.Clear();
            _overlay?.Hide();
        }

        private void OnSelectionChange(object sh, Excel.Range target)
        {
            if (_isEnabled) UpdateOverlay();
        }

        private void OnRefreshTick(object sender, EventArgs e)
        {
            if (_isEnabled) UpdateOverlay();
        }

        private void DisposeBitmap()
        {
            if (_currentBmp != null)
            {
                _currentBmp.Dispose();
                _currentBmp = null;
            }
        }

        private void UpdateOverlay()
        {
            try
            {
                if (_overlay == null || !_isEnabled) return;

                // Excel 非前台时隐藏
                IntPtr excelHwnd = (IntPtr)_app.Hwnd;
                IntPtr foreground = NativeMethods.GetForegroundWindow();
                if (foreground != excelHwnd && NativeMethods.GetAncestor(foreground, 2) != excelHwnd)
                {
                    DisposeBitmap();
                    _overlay.Clear();
                    return;
                }

                Excel.Window window = _app.ActiveWindow;
                if (window == null) return;

                Excel.Range selection = _app.Selection as Excel.Range;
                if (selection == null || IsEntireRowOrColumn(selection))
                {
                    DisposeBitmap();
                    _overlay.Clear();
                    return;
                }

                Excel.Range visRange = window.VisibleRange;
                if (visRange == null) return;

                Excel.Range visFirst = (Excel.Range)visRange.Cells[1, 1];
                double visFirstLeft = Convert.ToDouble(visFirst.Left);
                double visFirstTop = Convert.ToDouble(visFirst.Top);

                int pts2px_0_x = window.PointsToScreenPixelsX(0);
                int pts2px_0_y = window.PointsToScreenPixelsY(0);

                IntPtr gridHwnd = NativeMethods.FindExcelGridWindow(excelHwnd);
                NativeMethods.RECT gridRect = new NativeMethods.RECT();
                if (gridHwnd != IntPtr.Zero)
                    NativeMethods.GetWindowRect(gridHwnd, out gridRect);

                _zoom = Convert.ToDouble(window.Zoom) / 100.0;
                float dpiX, dpiY;
                using (var gfx = Graphics.FromHwnd(excelHwnd))
                {
                    dpiX = gfx.DpiX;
                    dpiY = gfx.DpiY;
                }
                _dpiScale = dpiY / 72.0;
                double scaleX = _zoom * dpiX / 72.0;
                double scaleY = _zoom * dpiY / 72.0;

                Excel.Worksheet activeSheet = _app.ActiveSheet as Excel.Worksheet;

                // Overlay 定位
                int overlayX = pts2px_0_x + (int)Math.Round(visFirstLeft * scaleX);
                int overlayY;
                if (activeSheet != null && visFirst.Row > 1)
                {
                    int scrollPixels = AccumulatePixelY(activeSheet, 1, visFirst.Row, scaleY);
                    overlayY = pts2px_0_y + scrollPixels;
                }
                else
                {
                    overlayY = pts2px_0_y + (int)Math.Round(visFirstTop * scaleY);
                }

                // Overlay 尺寸
                Excel.Range visLast = (Excel.Range)visRange.Cells[visRange.Rows.Count, visRange.Columns.Count];
                double visLastRight = Convert.ToDouble(visLast.Left) + Convert.ToDouble(visLast.Width);
                double visLastBottom = Convert.ToDouble(visLast.Top) + Convert.ToDouble(visLast.Height);
                double totalPtsW = visLastRight - visFirstLeft;
                double totalPtsH = visLastBottom - visFirstTop;

                int overlayW = (int)Math.Round(totalPtsW * scaleX);
                int overlayH = gridHwnd != IntPtr.Zero ? (gridRect.Bottom - overlayY) : AccumulatePixelY(activeSheet, visFirst.Row, visFirst.Row + visRange.Rows.Count, scaleY);

                // 边界裁剪
                if (gridHwnd != IntPtr.Zero)
                {
                    int maxW = gridRect.Right - overlayX;
                    int maxH = gridRect.Bottom - overlayY;
                    if (maxW < overlayW && maxW > 0) overlayW = maxW;
                    if (maxH < overlayH && maxH > 0) overlayH = maxH;
                }

                if (overlayW <= 0 || overlayH <= 0) return;

                DisposeBitmap();
                _currentBmp = new Bitmap(overlayW, overlayH, PixelFormat.Format32bppArgb);

                using (var g = Graphics.FromImage(_currentBmp))
                {
                    g.Clear(Color.Transparent);
                    using (var brush = new SolidBrush(HighlightColor))
                    {
                        foreach (Excel.Range area in selection.Areas)
                        {
                            double aLeft = Convert.ToDouble(area.Left);
                            double aRight = aLeft + Convert.ToDouble(area.Width);

                            int ax = (int)Math.Round((aLeft - visFirstLeft) * scaleX);
                            int ax2 = (int)Math.Round((aRight - visFirstLeft) * scaleX);
                            int aw = ax2 - ax;

                            int ay = AccumulatePixelY(activeSheet, visFirst.Row, area.Row, scaleY);
                            int ay2 = AccumulatePixelY(activeSheet, visFirst.Row, area.Row + area.Rows.Count, scaleY);
                            int ah = ay2 - ay;

                            // 行带（全宽，排除选中区域）
                            if (ax > 0)
                                g.FillRectangle(brush, 0, ay, ax, ah);
                            if (ax2 < overlayW)
                                g.FillRectangle(brush, ax2, ay, overlayW - ax2, ah);

                            // 列带（排除行带重叠和选中区域）
                            if (ay > 0)
                                g.FillRectangle(brush, ax, 0, aw, ay);
                            if (ay2 < overlayH)
                                g.FillRectangle(brush, ax, ay2, aw, overlayH - ay2);
                        }
                    }
                }

                _overlay.UpdateLayered(_currentBmp, overlayX, overlayY);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SpotlightService] UpdateOverlay: {ex.Message}");
            }
        }

        /// <summary>
        /// 逐行累加像素高度，匹配 Excel 的像素渲染逻辑：
        /// 1. baseH = Round(rowH_pts * dpi/72)
        /// 2. zoomedH = Round(baseH * zoom)
        /// </summary>
        private int AccumulatePixelY(Excel.Worksheet sheet, int fromRow, int toRow, double scaleY)
        {
            if (sheet == null || fromRow >= toRow) return 0;
            int total = 0;
            for (int r = fromRow; r < toRow; r++)
            {
                double rowH = Convert.ToDouble(((Excel.Range)sheet.Rows[r]).Height);
                int basePixel = (int)Math.Round(rowH * _dpiScale);
                int zoomedPixel = (int)Math.Round(basePixel * _zoom);
                total += zoomedPixel;
            }
            return total;
        }

        private bool IsEntireRowOrColumn(Excel.Range selection)
        {
            foreach (Excel.Range area in selection.Areas)
            {
                if (area.Columns.Count >= 16384) return true;
                if (area.Rows.Count >= 1048576) return true;
            }
            return false;
        }
    }

    internal class SpotlightOverlay : Form
    {
        public SpotlightOverlay()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                // WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00080000 | 0x00000020 | 0x00000080 | 0x08000000;
                return cp;
            }
        }

        protected override bool ShowWithoutActivation => true;

        public void Clear()
        {
            using (var bmp = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
                UpdateLayered(bmp, Left, Top);
        }

        public void UpdateLayered(Bitmap bmp, int screenX, int screenY)
        {
            IntPtr screenDc = NativeMethods.GetDC(IntPtr.Zero);
            IntPtr memDc = NativeMethods.CreateCompatibleDC(screenDc);
            IntPtr hBitmap = bmp.GetHbitmap(Color.FromArgb(0));
            IntPtr oldBitmap = NativeMethods.SelectObject(memDc, hBitmap);
            try
            {
                var size = new NativeMethods.SIZE { cx = bmp.Width, cy = bmp.Height };
                var srcPoint = new NativeMethods.POINT { x = 0, y = 0 };
                var dstPoint = new NativeMethods.POINT { x = screenX, y = screenY };
                var blend = new NativeMethods.BLENDFUNCTION
                {
                    BlendOp = 0, BlendFlags = 0,
                    SourceConstantAlpha = 255, AlphaFormat = 1
                };
                NativeMethods.UpdateLayeredWindow(
                    this.Handle, screenDc, ref dstPoint, ref size,
                    memDc, ref srcPoint, 0, ref blend, 2);
            }
            finally
            {
                NativeMethods.SelectObject(memDc, oldBitmap);
                NativeMethods.DeleteObject(hBitmap);
                NativeMethods.DeleteDC(memDc);
                NativeMethods.ReleaseDC(IntPtr.Zero, screenDc);
            }
        }
    }

    internal static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int x, y; }

        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE { public int cx, cy; }

        [StructLayout(LayoutKind.Sequential)]
        public struct BLENDFUNCTION
        {
            public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter,
            string lpszClass, string lpszWindow);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteObject(IntPtr hObject);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UpdateLayeredWindow(
            IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize,
            IntPtr hdcSrc, ref POINT pptSrc, int crKey,
            ref BLENDFUNCTION pblend, int dwFlags);

        public static IntPtr FindExcelGridWindow(IntPtr excelHwnd)
        {
            IntPtr xlDesk = FindWindowEx(excelHwnd, IntPtr.Zero, "XLDESK", null);
            if (xlDesk == IntPtr.Zero) return IntPtr.Zero;
            return FindWindowEx(xlDesk, IntPtr.Zero, "EXCEL7", null);
        }
    }
}
