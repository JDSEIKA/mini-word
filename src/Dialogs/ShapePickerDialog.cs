using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using WordPadClone.Controls;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class ShapePickerDialog : Form
    {
        private readonly ToolTip _toolTip = new ToolTip();

        public FloatingShapeKind SelectedKind { get; private set; }

        public ShapePickerDialog()
        {
            Text = "Insertar forma";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 228);

            Label hint = new Label();
            hint.Text = "Seleccione una forma:";
            hint.Location = new Point(14, 12);
            hint.AutoSize = true;
            Controls.Add(hint);

            string[] names = new string[]
            {
                "Rectángulo", "Círculo", "Triángulo", "Estrella",
                "Flecha derecha", "Flecha izquierda", "Rombo", "Corazón"
            };
            FloatingShapeKind[] kinds = new FloatingShapeKind[]
            {
                FloatingShapeKind.Rectangle, FloatingShapeKind.Ellipse, FloatingShapeKind.Triangle,
                FloatingShapeKind.Star, FloatingShapeKind.Arrow, FloatingShapeKind.ArrowLeft,
                FloatingShapeKind.Diamond, FloatingShapeKind.Heart
            };
            for (int i = 0; i < names.Length; i++)
            {
                Button button = new Button();
                button.Text = string.Empty;
                button.Tag = kinds[i];
                button.AccessibleName = names[i];
                button.Location = new Point(14 + (i % 4) * 86, 40 + (i / 4) * 68);
                button.Size = new Size(72, 58);
                button.FlatStyle = FlatStyle.Flat;
                button.BackColor = Color.White;
                button.UseVisualStyleBackColor = false;
                button.FlatAppearance.BorderColor = Color.FromArgb(180, 195, 210);
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 242, 252);
                button.Cursor = Cursors.Hand;
                FloatingShapeKind kind = kinds[i];
                button.Paint += delegate(object sender, PaintEventArgs e)
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    RectangleF shapeRect = new RectangleF(13, 9, 46, 40);
                    using (GraphicsPath path = ShapeGeometry.Create(kind, shapeRect))
                    using (Brush fill = new SolidBrush(Color.FromArgb(221, 235, 247)))
                    using (Pen border = new Pen(Color.FromArgb(68, 114, 196), 2f))
                    {
                        e.Graphics.FillPath(fill, path);
                        e.Graphics.DrawPath(border, path);
                    }
                };
                button.Click += delegate(object sender, EventArgs e)
                {
                    SelectedKind = (FloatingShapeKind)((Button)sender).Tag;
                    DialogResult = DialogResult.OK;
                    Close();
                };
                _toolTip.SetToolTip(button, names[i]);
                Controls.Add(button);
            }

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(296, 190);
            cancel.Width = 80;
            Controls.Add(cancel);
            CancelButton = cancel;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _toolTip.Dispose();
            base.Dispose(disposing);
        }
    }
}
