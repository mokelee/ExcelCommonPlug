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
    /// 支持普通模式、冻结窗格模式和拆分窗格模式。
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
        private string _lastLogKey = "";
        private bool _shouldLog;

        public SpotlightService(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public bool IsEnabled => _isEnabled;

        public void Enable()
        {
            _isEnabled = true;
            Logger.Debug("Spotlight", "Enable called");
            if (_overlay == null)
                _overlay = new SpotlightOverlay();
            _overlay.Show();
            _app.SheetSelectionChange += OnSelectionChange;
            _refreshTimer = new Timer { Interval = 100 };
            _refreshTimer.Tick += OnRefreshTick;
            _refreshTimer.Start();
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
            Logger.Debug("Spotlight", " Disabled");
        }

        private void OnSelectionChange(object sh, Excel.Range target)
        {
            // 选区变化由定时器统一处理
        }

        private void OnRefreshTick(object sender, EventArgs e)
        {
            if (!_isEnabled) return;
            UpdateOverlay();
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
                if (activeSheet == null) return;

                // 检测拆分/冻结状态
                bool hasFreezePane = false;
                bool hasSplit = false;
                int splitRow = 0;
                int splitCol = 0;
                try
                {
                    hasFreezePane = window.FreezePanes;
                    splitRow = window.SplitRow;
                    splitCol = window.SplitColumn;
                    hasSplit = !hasFreezePane && (splitRow > 0 || splitCol > 0);
                }
                catch { }

                // 生成日志 key，避免重复输出
                string selAddr = "";
                try { selAddr = selection.Address; } catch { }
                string logKey = $"{selAddr}|{hasFreezePane}|{hasSplit}|{splitRow}|{splitCol}|{gridRect.Left},{gridRect.Top}";
                bool shouldLog = (logKey != _lastLogKey);
                if (shouldLog)
                {
                    _lastLogKey = logKey;
                    _shouldLog = true;
                    Logger.Debug("Spotlight", $" --- UpdateOverlay ---");
                    Logger.Debug("Spotlight", $" Selection={selAddr}, hasFreezePane={hasFreezePane}, hasSplit={hasSplit}, splitRow={splitRow}, splitCol={splitCol}");
                }
                else
                {
                    _shouldLog = false;
                }

                // 确定 overlay 覆盖区域
                int overlayX, overlayY, overlayW, overlayH;

                if ((hasFreezePane || hasSplit) && gridHwnd != IntPtr.Zero)
                {
                    // overlay 覆盖整个网格窗口
                    overlayX = gridRect.Left;
                    overlayY = gridRect.Top;
                    overlayW = gridRect.Right - gridRect.Left;
                    overlayH = gridRect.Bottom - gridRect.Top;

                    if (_shouldLog)
                    {
                        Logger.Debug("Spotlight", $"Mode=FreezeOrSplit, gridRect=({gridRect.Left},{gridRect.Top},{gridRect.Right},{gridRect.Bottom})");
                        Logger.Debug("Spotlight", $"Overlay: pos=({overlayX},{overlayY}), size=({overlayW},{overlayH})");
                    }
                }
                else if (gridHwnd != IntPtr.Zero)
                {
                    Excel.Range visRange = window.VisibleRange;
                    if (visRange == null) return;
                    Excel.Range visFirst = (Excel.Range)visRange.Cells[1, 1];
                    double visFirstLeft = Convert.ToDouble(visFirst.Left);
                    double visFirstTop = Convert.ToDouble(visFirst.Top);

                    int pts2px_0_x = window.PointsToScreenPixelsX(0);
                    int pts2px_0_y = window.PointsToScreenPixelsY(0);

                    overlayX = pts2px_0_x + (int)Math.Round(visFirstLeft * scaleX);
                    overlayY = pts2px_0_y + (visFirst.Row > 1
                        ? AccumulatePixelY(activeSheet, 1, visFirst.Row, scaleY)
                        : (int)Math.Round(visFirstTop * scaleY));

                    overlayW = gridRect.Right - overlayX;
                    overlayH = gridRect.Bottom - overlayY;
                    if (_shouldLog)
                    {
                        Logger.Debug("Spotlight", $"Mode=Simple, pts2px(0)=({pts2px_0_x},{pts2px_0_y}), visFirst=R{visFirst.Row}C{visFirst.Column}, visFirstLeft={visFirstLeft:F2}, visFirstTop={visFirstTop:F2}");
                        Logger.Debug("Spotlight", $"Overlay: pos=({overlayX},{overlayY}), size=({overlayW},{overlayH}), gridRect=({gridRect.Left},{gridRect.Top},{gridRect.Right},{gridRect.Bottom})");
                        Logger.Debug("Spotlight", $"scaleX={scaleX:F4}, scaleY={scaleY:F4}, zoom={_zoom}, dpiScale={_dpiScale:F4}");
                    }
                }
                else
                {
                    return;
                }

                if (overlayW <= 0 || overlayH <= 0) return;

                DisposeBitmap();
                _currentBmp = new Bitmap(overlayW, overlayH, PixelFormat.Format32bppArgb);

                using (var g = Graphics.FromImage(_currentBmp))
                {
                    g.Clear(Color.Transparent);
                    using (var brush = new SolidBrush(HighlightColor))
                    {
                        if (hasFreezePane || hasSplit)
                        {
                            DrawWithPanes(g, brush, window, activeSheet, selection,
                                          splitRow, splitCol, hasFreezePane,
                                          overlayX, overlayY, overlayW, overlayH);
                        }
                        else
                        {
                            DrawSimple(g, brush, window, activeSheet, selection,
                                       overlayW, overlayH, scaleX, scaleY);
                        }
                    }
                }

                _overlay.UpdateLayered(_currentBmp, overlayX, overlayY);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SpotlightService] UpdateOverlay: {ex.Message}");
                Logger.Error("Spotlight", $"UpdateOverlay: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 无拆分/冻结时的简单绘制逻辑
        /// </summary>
        private void DrawSimple(Graphics g, SolidBrush brush,
            Excel.Window window, Excel.Worksheet sheet, Excel.Range selection,
            int overlayW, int overlayH, double scaleX, double scaleY)
        {
            Excel.Range visRange = window.VisibleRange;
            Excel.Range visFirst = (Excel.Range)visRange.Cells[1, 1];
            double visFirstLeft = Convert.ToDouble(visFirst.Left);

            if (_shouldLog) Logger.Debug("Spotlight", $"DrawSimple: visFirst=R{visFirst.Row}C{visFirst.Column}, visFirstLeft={visFirstLeft:F2}, overlayW={overlayW}, overlayH={overlayH}");

            foreach (Excel.Range area in selection.Areas)
            {
                double aLeft = Convert.ToDouble(area.Left);
                double aRight = aLeft + Convert.ToDouble(area.Width);

                int ax = (int)Math.Round((aLeft - visFirstLeft) * scaleX);
                int ax2 = (int)Math.Round((aRight - visFirstLeft) * scaleX);
                int aw = ax2 - ax;

                int ay = AccumulatePixelY(sheet, visFirst.Row, area.Row, scaleY);
                int ay2 = AccumulatePixelY(sheet, visFirst.Row, area.Row + area.Rows.Count, scaleY);
                int ah = ay2 - ay;

                if (_shouldLog) Logger.Debug("Spotlight", $"DrawSimple: sel=R{area.Row}C{area.Column}, aLeft={aLeft:F2}, ax={ax}, ay={ay}, aw={aw}, ah={ah}");

                if (ax > 0)
                    g.FillRectangle(brush, 0, ay, ax, ah);
                if (ax2 < overlayW)
                    g.FillRectangle(brush, ax2, ay, overlayW - ax2, ah);
                if (ay > 0)
                    g.FillRectangle(brush, ax, 0, aw, ay);
                if (ay2 < overlayH)
                    g.FillRectangle(brush, ax, ay2, aw, overlayH - ay2);
            }
        }

        /// <summary>
        /// 冻结/拆分模式下的分窗格绘制逻辑。
        /// 使用 Pane.PointsToScreenPixelsX/Y 精确获取每个单元格在屏幕上的位置。
        /// </summary>
        private void DrawWithPanes(Graphics g, SolidBrush brush,
            Excel.Window window, Excel.Worksheet sheet, Excel.Range selection,
            int splitRow, int splitCol, bool isFrozen,
            int overlayX, int overlayY, int overlayW, int overlayH)
        {
            Excel.Panes panes = window.Panes;
            int paneCount = panes.Count;

            if (_shouldLog) Logger.Debug("Spotlight", $"DrawWithPanes: isFrozen={isFrozen}, splitRow={splitRow}, splitCol={splitCol}, paneCount={paneCount}, overlayW={overlayW}, overlayH={overlayH}");

            foreach (Excel.Range area in selection.Areas)
            {
                int selRow = area.Row;
                int selCol = area.Column;
                int selRowEnd = selRow + area.Rows.Count;
                int selColEnd = selCol + area.Columns.Count;

                double selLeft = Convert.ToDouble(((Excel.Range)sheet.Cells[1, selCol]).Left);
                double selRight = Convert.ToDouble(((Excel.Range)sheet.Cells[1, selColEnd]).Left);
                double selTop = Convert.ToDouble(((Excel.Range)sheet.Cells[selRow, 1]).Top);
                double selBottom = Convert.ToDouble(((Excel.Range)sheet.Cells[selRowEnd, 1]).Top);

                for (int i = 0; i < paneCount; i++)
                {
                    Excel.Pane pane = panes[i + 1];

                    // 用 Pane.PointsToScreenPixelsX/Y 精确计算屏幕坐标
                    int ax = pane.PointsToScreenPixelsX((int)selLeft) - overlayX;
                    int ax2 = pane.PointsToScreenPixelsX((int)selRight) - overlayX;
                    int aw = ax2 - ax;

                    int ay = pane.PointsToScreenPixelsY((int)selTop) - overlayY;
                    int ay2 = pane.PointsToScreenPixelsY((int)selBottom) - overlayY;
                    int ah = ay2 - ay;

                    // 获取 Pane 的可见区域边界
                    Excel.Range paneVis = pane.VisibleRange;
                    Excel.Range paneFirst = (Excel.Range)paneVis.Cells[1, 1];
                    Excel.Range paneLast = (Excel.Range)paneVis.Cells[paneVis.Rows.Count, paneVis.Columns.Count];

                    int paneLeft = pane.PointsToScreenPixelsX((int)Convert.ToDouble(paneFirst.Left)) - overlayX;
                    int paneTop = pane.PointsToScreenPixelsY((int)Convert.ToDouble(paneFirst.Top)) - overlayY;
                    int paneRight = pane.PointsToScreenPixelsX((int)(Convert.ToDouble(paneLast.Left) + Convert.ToDouble(paneLast.Width))) - overlayX;
                    int paneBottom = pane.PointsToScreenPixelsY((int)(Convert.ToDouble(paneLast.Top) + Convert.ToDouble(paneLast.Height))) - overlayY;

                    int pw = paneRight - paneLeft;
                    int ph = paneBottom - paneTop;
                    if (pw <= 0 || ph <= 0) continue;

                    if (_shouldLog && i == paneCount - 1)
                        Logger.Debug("Spotlight", $"  Pane[{i}]: ax={ax}, ay={ay}, aw={aw}, ah={ah}, paneRect=({paneLeft},{paneTop},{pw},{ph})");

                    if (isFrozen)
                    {
                        // 行带：在该行可见的 Pane 中画
                        bool rowVisible = (ay >= paneTop && ay < paneBottom);
                        if (rowVisible)
                        {
                            g.SetClip(new Rectangle(paneLeft, paneTop, pw, ph));
                            if (ax > paneLeft)
                                g.FillRectangle(brush, paneLeft, ay, ax - paneLeft, ah);
                            if (ax2 < paneRight)
                                g.FillRectangle(brush, ax2, ay, paneRight - ax2, ah);
                        }

                        // 列带：在该列可见的 Pane 中画
                        bool colVisible = (ax >= paneLeft && ax < paneRight);
                        if (colVisible)
                        {
                            g.SetClip(new Rectangle(paneLeft, paneTop, pw, ph));
                            if (ay > paneTop)
                                g.FillRectangle(brush, ax, paneTop, aw, ay - paneTop);
                            if (ay2 < paneBottom)
                                g.FillRectangle(brush, ax, ay2, aw, paneBottom - ay2);
                        }
                    }
                    else
                    {
                        // 拆分模式：仅在选中单元格所在的 Pane 内绘制
                        if (ax >= paneRight || ax2 <= paneLeft) continue;
                        if (ay >= paneBottom || ay2 <= paneTop) continue;

                        g.SetClip(new Rectangle(paneLeft, paneTop, pw, ph));
                        if (ax > paneLeft)
                            g.FillRectangle(brush, paneLeft, ay, ax - paneLeft, ah);
                        if (ax2 < paneRight)
                            g.FillRectangle(brush, ax2, ay, paneRight - ax2, ah);
                        if (ay > paneTop)
                            g.FillRectangle(brush, ax, paneTop, aw, ay - paneTop);
                        if (ay2 < paneBottom)
                            g.FillRectangle(brush, ax, ay2, aw, paneBottom - ay2);
                    }
                }
            }

            g.ResetClip();
        }

        /// <summary>
        /// 逐行累加像素高度
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
