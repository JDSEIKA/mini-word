using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using WordPadClone.Controls;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class ShapeFormatDialog : Form
    {
        private readonly FloatingShapeKind _kind;
        private readonly Button _fillButton;
        private readonly Button _borderButton;
        private readonly NumericUpDown _borderWidth;
        private readonly Panel _preview;

        public Color FillColor { get; private set; }
        public Color BorderColor { get; private set; }
        public int BorderSize { get { return (int)_borderWidth.Value; } }

        public ShapeFormatDialog(FloatingShapeKind kind, Color fill, Color border, int borderSize)
        {
            _kind = kind;
            FillColor = fill;
            BorderColor = border;
            Text = "Formato de forma";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(380, 260);

            Label fillLabel = new Label();
            fillLabel.Text = "Color de relleno:";
            fillLabel.Location = new Point(16, 20);
            fillLabel.AutoSize = true;
            Controls.Add(fillLabel);

            _fillButton = new Button();
            _fillButton.Location = new Point(150, 15);
            _fillButton.Size = new Size(80, 28);
            _fillButton.FlatStyle = FlatStyle.Flat;
            _fillButton.BackColor = fill;
            _fillButton.Click += delegate { ChooseColor(true); };
            Controls.Add(_fillButton);

            Label borderLabel = new Label();
            borderLabel.Text = "Color del contorno:";
            borderLabel.Location = new Point(16, 60);
            borderLabel.AutoSize = true;
            Controls.Add(borderLabel);

            _borderButton = new Button();
            _borderButton.Location = new Point(150, 55);
            _borderButton.Size = new Size(80, 28);
            _borderButton.FlatStyle = FlatStyle.Flat;
            _borderButton.BackColor = border;
            _borderButton.Click += delegate { ChooseColor(false); };
            Controls.Add(_borderButton);

            Label widthLabel = new Label();
            widthLabel.Text = "Grosor del contorno:";
            widthLabel.Location = new Point(16, 100);
            widthLabel.AutoSize = true;
            Controls.Add(widthLabel);

            _borderWidth = new NumericUpDown();
            _borderWidth.Location = new Point(150, 95);
            _borderWidth.Width = 100;
            _borderWidth.Minimum = 1;
            _borderWidth.Maximum = 16;
            _borderWidth.Value = Math.Max(1, Math.Min(16, borderSize));
            _borderWidth.TextAlign = HorizontalAlignment.Left;
            _borderWidth.ForeColor = Color.Black;
            _borderWidth.BackColor = Color.White;
            _borderWidth.ValueChanged += delegate { _preview.Invalidate(); };
            Controls.Add(_borderWidth);

            Label widthUnit = new Label();
            widthUnit.Text = "px";
            widthUnit.Location = new Point(150, 121);
            widthUnit.AutoSize = true;
            Controls.Add(widthUnit);

            _preview = new Panel();
            _preview.Location = new Point(250, 15);
            _preview.Size = new Size(110, 110);
            _preview.BackColor = Color.White;
            _preview.BorderStyle = BorderStyle.FixedSingle;
            _preview.Paint += delegate(object sender, PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                RectangleF rect = new RectangleF(15, 15, 80, 80);
                using (GraphicsPath path = ShapeGeometry.Create(_kind, rect))
                using (Brush previewFill = new SolidBrush(FillColor))
                using (Pen previewBorder = new Pen(BorderColor, Math.Max(1, BorderSize)))
                {
                    e.Graphics.FillPath(previewFill, path);
                    e.Graphics.DrawPath(previewBorder, path);
                }
            };
            Controls.Add(_preview);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(196, 215);
            ok.Width = 80;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(282, 215);
            cancel.Width = 80;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void ChooseColor(bool fill)
        {
            using (ColorDialog dialog = new ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = fill ? FillColor : BorderColor;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                if (fill)
                {
                    FillColor = dialog.Color;
                    _fillButton.BackColor = FillColor;
                }
                else
                {
                    BorderColor = dialog.Color;
                    _borderButton.BackColor = BorderColor;
                }
                _preview.Invalidate();
            }
        }
    }
}
