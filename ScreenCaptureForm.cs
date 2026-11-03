using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic
{
    public partial class ScreenCaptureForm : Form
    {
        #region Win32 API

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT pt);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetShellWindow();

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint GA_ROOTOWNER = 3;

        #endregion

        private Bitmap _screenBitmap;
        private Rectangle _virtualScreen;
        private Rectangle _highlightRect = Rectangle.Empty;
        private Rectangle _selectRect = Rectangle.Empty;
        private Rectangle _lastSelectRect = Rectangle.Empty;
        private Point _startPoint;
        private bool _isDragging = false;
        private bool _isSelected = false;
        private Font _tipFont;
        private SolidBrush _maskBrush;
        private Pen _highlightPen;
        private Pen _selectPen;
        private SolidBrush _tipBrush;
        private SolidBrush _tipBgBrush;

        public Bitmap CapturedImage { get; private set; }

        public ScreenCaptureForm()
        {
            InitializeComponent();
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);

            _tipFont = new Font("微软雅黑", 11F, FontStyle.Regular);
            _maskBrush = new SolidBrush(Color.FromArgb(100, 0, 0, 0));
            _highlightPen = new Pen(Color.FromArgb(0, 120, 215), 2);
            _selectPen = new Pen(Color.FromArgb(0, 120, 215), 2);
            _tipBrush = new SolidBrush(Color.White);
            _tipBgBrush = new SolidBrush(Color.FromArgb(180, 0, 0, 0));
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            _virtualScreen = SystemInformation.VirtualScreen;
            this.Bounds = _virtualScreen;
            this.Cursor = Cursors.Cross;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            CaptureScreen();
            this.Activate();
            this.Invalidate();
        }

        private void CaptureScreen()
        {
            _screenBitmap = new Bitmap(_virtualScreen.Width, _virtualScreen.Height);
            using (Graphics g = Graphics.FromImage(_screenBitmap))
            {
                g.CopyFromScreen(_virtualScreen.X, _virtualScreen.Y, 0, 0, _screenBitmap.Size);
            }
        }

        private Rectangle GetWindowRectAtPoint(Point pt)
        {
            IntPtr desktop = GetDesktopWindow();
            IntPtr shell = GetShellWindow();
            IntPtr self = this.Handle;

            IntPtr hWnd = WindowFromPoint(new POINT { X = pt.X, Y = pt.Y });
            if (hWnd != IntPtr.Zero && hWnd != desktop && hWnd != shell)
            {
                IntPtr root = GetAncestor(hWnd, GA_ROOTOWNER);
                if (root != IntPtr.Zero && root != desktop && root != shell && root != self && IsWindowVisible(root))
                {
                    Rectangle r = RectFromHandle(root);
                    if (r != Rectangle.Empty && r.Contains(pt))
                        return r;
                }
            }

            return FindTopLevelWindowAtPoint(pt, desktop, shell, self);
        }

        private Rectangle FindTopLevelWindowAtPoint(Point pt, IntPtr desktop, IntPtr shell, IntPtr self)
        {
            Rectangle found = Rectangle.Empty;

            EnumWindows((hWnd, lParam) =>
            {
                if (hWnd == desktop || hWnd == shell || hWnd == self)
                    return true;
                if (!IsWindowVisible(hWnd))
                    return true;

                RECT r;
                if (!GetWindowRect(hWnd, out r))
                    return true;

                int w = r.Right - r.Left;
                int h = r.Bottom - r.Top;
                if (w <= 0 || h <= 0)
                    return true;

                if (pt.X < r.Left || pt.X > r.Right || pt.Y < r.Top || pt.Y > r.Bottom)
                    return true;

                int len = GetWindowTextLength(hWnd);
                if (len <= 0)
                    return true;

                found = new Rectangle(r.Left, r.Top, w, h);
                return false;
            }, IntPtr.Zero);

            return found;
        }

        private Rectangle RectFromHandle(IntPtr hWnd)
        {
            RECT r;
            if (!GetWindowRect(hWnd, out r))
                return Rectangle.Empty;
            int w = r.Right - r.Left;
            int h = r.Bottom - r.Top;
            if (w <= 0 || h <= 0)
                return Rectangle.Empty;
            return new Rectangle(r.Left, r.Top, w, h);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);

            if (e.Button == MouseButtons.Left)
            {
                if (_isSelected && _selectRect.Width > 0 && _selectRect.Height > 0)
                {
                    _lastSelectRect = _selectRect;
                }
                _isDragging = true;
                _isSelected = false;
                _startPoint = e.Location;
                _selectRect = new Rectangle(e.Location, Size.Empty);

                this.Invalidate();
            }
            else if (e.Button == MouseButtons.Right)
            {
                CancelCapture();
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (_isDragging)
            {
                if (Math.Abs(e.X - _startPoint.X) > 3 || Math.Abs(e.Y - _startPoint.Y) > 3)
                {
                    _lastSelectRect = Rectangle.Empty;
                }
                int x = Math.Min(_startPoint.X, e.X);
                int y = Math.Min(_startPoint.Y, e.Y);
                int w = Math.Abs(e.X - _startPoint.X);
                int h = Math.Abs(e.Y - _startPoint.Y);
                _selectRect = new Rectangle(x, y, w, h);
                this.Invalidate();
            }
            else if (!_isSelected)
            {
                Point screenPt = new Point(e.X + _virtualScreen.X, e.Y + _virtualScreen.Y);
                Rectangle winRect = GetWindowRectAtPoint(screenPt);
                if (winRect != Rectangle.Empty)
                {
                    _highlightRect = new Rectangle(
                        winRect.X - _virtualScreen.X,
                        winRect.Y - _virtualScreen.Y,
                        winRect.Width,
                        winRect.Height);
                }
                else
                {
                    _highlightRect = Rectangle.Empty;
                }
                this.Invalidate();
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);

            if (e.Button == MouseButtons.Left && _isDragging)
            {
                _isDragging = false;
                if (_selectRect.Width > 5 && _selectRect.Height > 5)
                {
                    _isSelected = true;
                    _highlightRect = Rectangle.Empty;
                    this.Invalidate();
                }
                else
                {
                    _selectRect = Rectangle.Empty;
                    this.Invalidate();
                }
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);

            _isDragging = false;

            if (e.Button == MouseButtons.Left)
            {
                if (_lastSelectRect.Width > 5 && _lastSelectRect.Height > 5)
                {
                    _selectRect = _lastSelectRect;
                    _isSelected = true;
                    _lastSelectRect = Rectangle.Empty;
                    ConfirmCapture();
                    return;
                }

                if (_isSelected && _selectRect.Width > 0 && _selectRect.Height > 0)
                {
                    ConfirmCapture();
                    return;
                }

                Point screenPt = new Point(e.X + _virtualScreen.X, e.Y + _virtualScreen.Y);
                Rectangle winRect = GetWindowRectAtPoint(screenPt);
                if (winRect != Rectangle.Empty)
                {
                    _selectRect = new Rectangle(
                        winRect.X - _virtualScreen.X,
                        winRect.Y - _virtualScreen.Y,
                        winRect.Width,
                        winRect.Height);
                    _isSelected = true;
                    _highlightRect = Rectangle.Empty;
                    ConfirmCapture();
                }
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                CancelCapture();
                return true;
            }
            if (keyData == Keys.Enter)
            {
                if (_isSelected)
                {
                    ConfirmCapture();
                    return true;
                }
                if (_highlightRect != Rectangle.Empty && !_isDragging)
                {
                    _selectRect = _highlightRect;
                    _isSelected = true;
                    _highlightRect = Rectangle.Empty;
                    ConfirmCapture();
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ConfirmCapture()
        {
            if (_selectRect.Width <= 0 || _selectRect.Height <= 0)
            {
                CancelCapture();
                return;
            }

            Rectangle captureRect = new Rectangle(
                _selectRect.X + _virtualScreen.X,
                _selectRect.Y + _virtualScreen.Y,
                _selectRect.Width,
                _selectRect.Height);

            CapturedImage = new Bitmap(_selectRect.Width, _selectRect.Height);
            using (Graphics g = Graphics.FromImage(CapturedImage))
            {
                g.CopyFromScreen(captureRect.X, captureRect.Y, 0, 0, CapturedImage.Size);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelCapture()
        {
            CapturedImage = null;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (_screenBitmap != null)
            {
                g.DrawImage(_screenBitmap, 0, 0);
            }
            else
            {
                g.Clear(Color.Black);
            }

            bool isDragOrSelected = (_isDragging || _isSelected) && _selectRect.Width > 0 && _selectRect.Height > 0;

            if (isDragOrSelected)
            {
                using (Region region = new Region(this.ClientRectangle))
                {
                    region.Exclude(_selectRect);
                    g.FillRegion(_maskBrush, region);
                }
                g.DrawRectangle(_selectPen, _selectRect);
                DrawSizeLabel(g, _selectRect);
            }
            else if (_highlightRect != Rectangle.Empty)
            {
                using (Region region = new Region(this.ClientRectangle))
                {
                    region.Exclude(_highlightRect);
                    g.FillRegion(_maskBrush, region);
                }
                g.DrawRectangle(_highlightPen, _highlightRect);
                DrawSizeLabel(g, _highlightRect);
            }
            else
            {
                g.FillRectangle(_maskBrush, this.ClientRectangle);
            }

            DrawTipText(g);
        }

        private void DrawSizeLabel(Graphics g, Rectangle rect)
        {
            string sizeText = string.Format("{0} × {1}", rect.Width, rect.Height);
            using (Font f = new Font("微软雅黑", 9F))
            {
                SizeF textSize = g.MeasureString(sizeText, f);
                int labelX = rect.X;
                int labelY = rect.Y - (int)textSize.Height - 4;
                if (labelY < 0) labelY = rect.Y + 4;

                using (SolidBrush bg = new SolidBrush(Color.FromArgb(200, 0, 120, 215)))
                {
                    g.FillRectangle(bg, labelX, labelY, textSize.Width + 8, textSize.Height + 2);
                }
                g.DrawString(sizeText, f, _tipBrush, labelX + 4, labelY + 1);
            }
        }

        private void DrawTipText(Graphics g)
        {
            string tip;
            if (_isSelected)
            {
                tip = "Enter 确认截图并插入Excel  |  Esc 取消  |  右键取消";
            }
            else if (_isDragging)
            {
                tip = "拖动鼠标选择区域  |  Esc 取消";
            }
            else
            {
                tip = "移动鼠标自动识别窗口  |  拖动鼠标画框选择  |  双击截取当前窗口  |  Esc 取消";
            }

            using (Font f = new Font("微软雅黑", 10F))
            {
                SizeF textSize = g.MeasureString(tip, f);
                float x = (this.ClientRectangle.Width - textSize.Width) / 2;
                float y = this.ClientRectangle.Height - textSize.Height - 20;

                g.FillRectangle(_tipBgBrush, x - 8, y - 4, textSize.Width + 16, textSize.Height + 8);
                g.DrawString(tip, f, _tipBrush, x, y);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                _tipFont?.Dispose();
                _maskBrush?.Dispose();
                _highlightPen?.Dispose();
                _selectPen?.Dispose();
                _tipBrush?.Dispose();
                _tipBgBrush?.Dispose();
                _screenBitmap?.Dispose();
            }
            base.Dispose(disposing);
        }

        public static void InsertImageIntoExcel(Bitmap image)
        {
            if (image == null) return;

            string tempPath = System.IO.Path.GetTempFileName() + ".png";
            image.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);

            try
            {
                Excel.Application app = ThisAddIn.app;
                Excel.Worksheet sheet = (Excel.Worksheet)app.ActiveSheet;
                Excel.Range sel = (Excel.Range)app.Selection;

                Excel.Range endCell = sheet.Cells[sel.Row + 9, sel.Column + 19];
                Excel.Range targetRange = sheet.Range[sel, endCell];
                double targetWidth = (double)targetRange.Width;
                double targetHeight = (double)targetRange.Height;

                double aspectRatio = (double)image.Width / image.Height;
                double finalWidth, finalHeight;
                if (targetWidth / targetHeight > aspectRatio)
                {
                    finalHeight = targetHeight;
                    finalWidth = targetHeight * aspectRatio;
                }
                else
                {
                    finalWidth = targetWidth;
                    finalHeight = targetWidth / aspectRatio;
                }

                Excel.Shape shape = sheet.Shapes.AddPicture(
                    tempPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    sel.Left, sel.Top, (float)finalWidth, (float)finalHeight);
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("插入图片到Excel失败: " + ex.Message);
            }
            finally
            {
                try { System.IO.File.Delete(tempPath); } catch { }
            }
        }
    }
}