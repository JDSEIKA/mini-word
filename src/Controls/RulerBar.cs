using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace WordPadClone.Controls
{
    /// <summary>
    /// Fixed ruler strip.  It lives outside the scrollable editor so scrolling
    /// never has to repaint a floating overlay and cannot leave stale copies.
    /// </summary>
    internal sealed class RulerBar : Control
    {
        private readonly NativeDocumentEditor _editor;

        public RulerBar(NativeDocumentEditor editor)
        {
            _editor = editor;
            AccessibleName = "Regla horizontal";
            AccessibleDescription = "Regla de página en centímetros";
            AccessibleRole = AccessibleRole.Indicator;
            Height = 22;
            Dock = DockStyle.Top;
            BackColor = Color.FromArgb(241, 244, 248);
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            if (_editor != null)
            {
                _editor.Scroll += delegate { Invalidate(); };
                _editor.ViewChanged += delegate { Invalidate(); };
                _editor.SizeChanged += delegate { Invalidate(); };
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.Clear(BackColor);
            if (_editor == null) return;

            float pageLeft, pageWidth, leftMargin, rightMargin;
            _editor.GetRulerMetrics(out pageLeft, out pageWidth, out leftMargin, out rightMargin);
            float pxPerCm = 96f / 2.54f * _editor.Zoom;
            if (pxPerCm <= 0f || pageWidth <= 1f) return;

            float rulerHeight = Math.Max(18f, Height);
            using (Brush margin = new SolidBrush(Color.FromArgb(226, 222, 210)))
            using (Pen edge = new Pen(Color.FromArgb(160, 155, 140)))
            using (Pen tick = new Pen(Color.FromArgb(95, 90, 80), 1f))
            using (Brush text = new SolidBrush(Color.FromArgb(70, 65, 55)))
            using (Font font = new Font("Segoe UI", 6.5f))
            {
                float left = pageLeft;
                float right = pageLeft + pageWidth;
                g.FillRectangle(margin, left, 0f, Math.Max(0f, leftMargin), rulerHeight);
                g.FillRectangle(margin, right - Math.Max(0f, rightMargin), 0f,
                    Math.Max(0f, rightMargin), rulerHeight);
                g.DrawRectangle(edge, left, 0f, pageWidth, rulerHeight);

                string widthLabel = (pageWidth / pxPerCm).ToString("0.0", CultureInfo.InvariantCulture) + " cm";
                SizeF measured = g.MeasureString(widthLabel, font);
                float labelLeft = right + 4f;
                bool inside = labelLeft + measured.Width > ClientSize.Width - 2f;
                if (inside) labelLeft = right - measured.Width - 5f;

                int centimeters = (int)Math.Floor(pageWidth / pxPerCm);
                for (int i = 0; i <= centimeters; i++)
                {
                    float x = left + i * pxPerCm;
                    bool major = i % 5 == 0;
                    float length = major ? 9f : 5f;
                    g.DrawLine(tick, x, rulerHeight - length, x, rulerHeight);
                    if (i < centimeters && (!inside || x + 12f < labelLeft))
                        g.DrawString(i.ToString(), font, text, x + 1f, 1f);
                }

                if (inside)
                {
                    g.FillRectangle(Brushes.White, labelLeft, 0f,
                        right - labelLeft, rulerHeight);
                    g.DrawRectangle(edge, labelLeft, 0f,
                        right - labelLeft, rulerHeight);
                }
                g.DrawString(widthLabel, font, text, labelLeft + 2f, 1f);
                g.DrawLine(edge, 0f, rulerHeight - 1f, ClientSize.Width, rulerHeight - 1f);
            }
        }
    }
}
