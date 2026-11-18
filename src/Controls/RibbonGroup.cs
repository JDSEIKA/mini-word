using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Controls
{
    /// <summary>A Word-style ribbon group: a strip of commands with a caption.</summary>
    internal sealed class RibbonGroup : TableLayoutPanel
    {
        public readonly ToolStrip Strip;
        private readonly Label _caption;

        public RibbonGroup(string caption)
        {
            BackColor = Color.FromArgb(249, 250, 252);
            Margin = new Padding(2, 2, 2, 2);
            Padding = new Padding(4, 2, 4, 0);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ColumnCount = 1;
            RowCount = 2;
            RowStyles.Add(new RowStyle(SizeType.Absolute, 64f));
            RowStyles.Add(new RowStyle(SizeType.Absolute, 17f));

            Strip = new ToolStrip();
            Strip.GripStyle = ToolStripGripStyle.Hidden;
            Strip.Dock = DockStyle.Fill;
            Strip.AutoSize = true;
            Strip.CanOverflow = false;
            Strip.Padding = new Padding(0);
            Strip.Margin = new Padding(0);
            Strip.ImageScalingSize = new Size(32, 32);
            Strip.BackColor = Color.Transparent;
            Strip.Renderer = new RibbonRenderer();
            Strip.LayoutStyle = ToolStripLayoutStyle.Flow;
            Controls.Add(Strip, 0, 0);

            _caption = new Label();
            _caption.Text = caption ?? string.Empty;
            _caption.Dock = DockStyle.Fill;
            _caption.TextAlign = ContentAlignment.MiddleCenter;
            _caption.Font = new Font("Segoe UI", 7.5f);
            _caption.ForeColor = Color.FromArgb(92, 104, 118);
            _caption.BackColor = Color.Transparent;
            _caption.AutoEllipsis = true;
            Controls.Add(_caption, 0, 1);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (Pen pen = new Pen(Color.FromArgb(220, 225, 232)))
            {
                e.Graphics.DrawLine(pen, Width - 1, 4, Width - 1, Height - 6);
            }
        }
    }

    internal sealed class RibbonRenderer : ToolStripProfessionalRenderer
    {
        public RibbonRenderer()
            : base(new RibbonColorTable())
        {
            RoundedEdges = false;
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            // Flat ribbon: no chrome around the command strip.
        }
    }

    internal sealed class RibbonColorTable : ProfessionalColorTable
    {
        public override Color ToolStripGradientBegin { get { return Color.FromArgb(249, 250, 252); } }
        public override Color ToolStripGradientMiddle { get { return Color.FromArgb(249, 250, 252); } }
        public override Color ToolStripGradientEnd { get { return Color.FromArgb(249, 250, 252); } }
        public override Color ButtonSelectedHighlight { get { return Color.FromArgb(220, 232, 247); } }
        public override Color ButtonSelectedBorder { get { return Color.FromArgb(160, 190, 225); } }
        public override Color ButtonPressedHighlight { get { return Color.FromArgb(201, 220, 242); } }
        public override Color ButtonPressedBorder { get { return Color.FromArgb(120, 160, 205); } }
        public override Color ButtonCheckedGradientBegin { get { return Color.FromArgb(215, 229, 246); } }
        public override Color ButtonCheckedGradientMiddle { get { return Color.FromArgb(215, 229, 246); } }
        public override Color ButtonCheckedGradientEnd { get { return Color.FromArgb(215, 229, 246); } }
        public override Color ButtonCheckedHighlight { get { return Color.FromArgb(215, 229, 246); } }
        public override Color ButtonCheckedHighlightBorder { get { return Color.FromArgb(140, 175, 215); } }
        public override Color MenuItemSelected { get { return Color.FromArgb(220, 232, 247); } }
        public override Color MenuBorder { get { return Color.FromArgb(200, 207, 216); } }
        public override Color MenuItemBorder { get { return Color.FromArgb(160, 190, 225); } }
        public override Color StatusStripGradientBegin { get { return Color.FromArgb(238, 242, 247); } }
        public override Color StatusStripGradientEnd { get { return Color.FromArgb(238, 242, 247); } }
    }
}
