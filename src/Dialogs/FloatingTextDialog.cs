using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class FloatingTextDialog : Form
    {
        private readonly TextBox _text;
        private readonly ComboBox _font;
        private readonly NumericUpDown _size;
        private readonly CheckBox _bold;
        private readonly CheckBox _italic;
        private readonly CheckBox _underline;
        private readonly Button _color;
        private readonly Button _background;
        private readonly CheckBox _transparent;
        private Color? _foreColor;
        private Color? _fillColor;
        private readonly RunFormat _initial;

        public string Value { get { return _text.Text; } }

        public RunFormat Format
        {
            get
            {
                RunFormat result = _initial.Clone();
                result.FontFamily = string.IsNullOrEmpty(_font.Text) ? result.FontFamily : _font.Text;
                result.SizeHalfPoints = Math.Max(2, (int)Math.Round(_size.Value * 2m));
                result.Bold = _bold.Checked;
                result.Italic = _italic.Checked;
                result.Underline = _underline.Checked;
                result.ForeColor = _foreColor;
                return result;
            }
        }

        public FloatingTextDialog(string value)
            : this(value, new RunFormat("Arial", 20))
        {
        }

        public FloatingTextDialog(string value, RunFormat format, Color? fill = null)
        {
            _initial = format == null ? new RunFormat("Arial", 20) : format.Clone();
            _foreColor = _initial.ForeColor;
            _fillColor = fill ?? Color.FromArgb(250, 250, 250);
            Text = "Texto del cuadro";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            ClientSize = new Size(500, 320);
            MinimumSize = new Size(400, 240);

            Panel top = new Panel();
            top.Dock = DockStyle.Top;
            top.Height = 70;
            Controls.Add(top);

            Label fontLabel = new Label();
            fontLabel.Text = "Fuente:";
            fontLabel.Location = new Point(8, 12);
            fontLabel.AutoSize = true;
            top.Controls.Add(fontLabel);

            _font = new ComboBox();
            _font.Location = new Point(58, 8);
            _font.Width = 140;
            _font.DropDownStyle = ComboBoxStyle.DropDown;
            _font.Items.AddRange(new string[] { "Arial", "Calibri", "Segoe UI", "Times New Roman", "Courier New", "Georgia", "Verdana" });
            _font.Text = _initial.FontFamily;
            top.Controls.Add(_font);

            Label sizeLabel = new Label();
            sizeLabel.Text = "Tamaño:";
            sizeLabel.Location = new Point(208, 12);
            sizeLabel.AutoSize = true;
            top.Controls.Add(sizeLabel);

            _size = new NumericUpDown();
            _size.Location = new Point(262, 8);
            _size.Width = 58;
            _size.Minimum = 6;
            _size.Maximum = 96;
            _size.DecimalPlaces = 1;
            _size.Increment = 0.5m;
            _size.Value = Math.Max(6m, Math.Min(96m, _initial.SizeHalfPoints / 2m));
            top.Controls.Add(_size);

            _bold = FormatToggle("N", 332, top, _initial.Bold, FontStyle.Bold);
            _italic = FormatToggle("I", 368, top, _initial.Italic, FontStyle.Italic);
            _underline = FormatToggle("S", 404, top, _initial.Underline, FontStyle.Underline);

            _color = new Button();
            _color.Text = "Color";
            _color.Location = new Point(440, 8);
            _color.Width = 50;
            _color.Height = 26;
            _color.BackColor = _foreColor ?? Color.Black;
            _color.ForeColor = Contrast(_color.BackColor);
            _color.Click += delegate { ChooseColor(); };
            top.Controls.Add(_color);

            Label fillLabel = new Label();
            fillLabel.Text = "Fondo:";
            fillLabel.Location = new Point(8, 44);
            fillLabel.AutoSize = true;
            top.Controls.Add(fillLabel);

            _background = new Button();
            _background.Text = "Color";
            _background.Location = new Point(58, 40);
            _background.Width = 60;
            _background.Height = 26;
            _background.BackColor = _fillColor.Value;
            _background.ForeColor = Contrast(_background.BackColor);
            _background.Click += delegate { ChooseFillColor(); };
            top.Controls.Add(_background);

            _transparent = new CheckBox();
            _transparent.Text = "Sin fondo";
            _transparent.Location = new Point(126, 44);
            _transparent.AutoSize = true;
            _transparent.CheckedChanged += delegate { UpdatePreview(); };
            top.Controls.Add(_transparent);

            _text = new TextBox();
            _text.Multiline = true;
            _text.ScrollBars = ScrollBars.Vertical;
            _text.Dock = DockStyle.Fill;
            _text.Text = value ?? string.Empty;
            _text.Font = new Font(_font.Text, (float)_size.Value);
            _font.SelectedIndexChanged += delegate { UpdatePreview(); };
            _font.TextChanged += delegate { UpdatePreview(); };
            _size.ValueChanged += delegate { UpdatePreview(); };
            Controls.Add(_text);
            _text.BringToFront();

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 42;
            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(326, 8);
            ok.Width = 74;
            bottom.Controls.Add(ok);
            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(406, 8);
            cancel.Width = 74;
            bottom.Controls.Add(cancel);
            Controls.Add(bottom);
            AcceptButton = ok;
            CancelButton = cancel;
            Shown += delegate { _text.Focus(); };
        }

        private CheckBox FormatToggle(string text, int x, Control parent, bool value, FontStyle style)
        {
            CheckBox box = new CheckBox();
            box.Text = text;
            box.Location = new Point(x, 8);
            box.Size = new Size(30, 28);
            box.Appearance = Appearance.Button;
            box.TextAlign = ContentAlignment.MiddleCenter;
            box.Checked = value;
            box.Font = new Font("Segoe UI", 9f, style);
            parent.Controls.Add(box);
            box.CheckedChanged += delegate { UpdatePreview(); };
            return box;
        }

        private void UpdatePreview()
        {
            try
            {
                _text.Font = new Font(_font.Text, (float)_size.Value,
                    (_bold.Checked ? FontStyle.Bold : FontStyle.Regular)
                    | (_italic.Checked ? FontStyle.Italic : FontStyle.Regular)
                    | (_underline.Checked ? FontStyle.Underline : FontStyle.Regular));
                _text.ForeColor = _foreColor ?? Color.Black;
                _text.BackColor = _transparent.Checked ? SystemColors.Window : _fillColor.Value;
            }
            catch
            {
            }
        }

        public Color? FillColor { get { return _transparent.Checked ? (Color?)null : _fillColor; } }

        private void ChooseFillColor()
        {
            using (ColorDialog dialog = new ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = _fillColor.Value;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _fillColor = dialog.Color;
                _background.BackColor = dialog.Color;
                _background.ForeColor = Contrast(dialog.Color);
                UpdatePreview();
            }
        }

        private void ChooseColor()
        {
            using (ColorDialog dialog = new ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = _foreColor ?? Color.Black;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _foreColor = dialog.Color;
                _color.BackColor = dialog.Color;
                _color.ForeColor = Contrast(dialog.Color);
                UpdatePreview();
            }
        }

        private static Color Contrast(Color color)
        {
            int brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
            return brightness > 140 ? Color.Black : Color.White;
        }
    }
}
