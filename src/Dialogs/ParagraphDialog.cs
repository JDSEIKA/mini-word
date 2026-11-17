using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class ParagraphDialog : Form
    {
        private readonly ComboBox _alignment;
        private readonly ComboBox _outline;
        private readonly NumericUpDown _left;
        private readonly NumericUpDown _right;
        private readonly ComboBox _special;
        private readonly NumericUpDown _by;
        private readonly NumericUpDown _before;
        private readonly NumericUpDown _after;
        private readonly ComboBox _lineSpacing;
        private readonly CheckBox _noSpace;
        private readonly CheckBox _pageBreak;
        private readonly CheckBox _keepLines;
        private readonly CheckBox _keepNext;
        private readonly CheckBox _widow;

        public ParagraphAlignment Alignment
        {
            get { return (ParagraphAlignment)_alignment.SelectedIndex; }
        }
        public int LeftIndentTwips { get { return CmToTwips(_left.Value); } }
        public int RightIndentTwips { get { return CmToTwips(_right.Value); } }
        public int FirstLineIndentTwips
        {
            get
            {
                if (_special.SelectedIndex == 0) return 0;
                int amount = CmToTwips(_by.Value);
                return _special.SelectedIndex == 1 ? amount : -amount;
            }
        }
        public int SpaceBeforeTwips { get { return (int)Math.Round(_before.Value * 20m); } }
        public int SpaceAfterTwips { get { return (int)Math.Round(_after.Value * 20m); } }
        public int LineSpacingValue
        {
            get
            {
                if (_lineSpacing.SelectedIndex <= 1) return 0;
                if (_lineSpacing.SelectedIndex == 2) return 1;
                return 2;
            }
        }
        public bool PageBreakBefore { get { return _pageBreak.Checked; } }
        public bool NoSpaceSameStyle { get { return _noSpace.Checked; } }
        public int OutlineLevel { get { return _outline.SelectedIndex; } }
        public bool KeepLinesTogether { get { return _keepLines.Checked; } }
        public bool KeepWithNext { get { return _keepNext.Checked; } }
        public bool WidowControl { get { return _widow.Checked; } }

        public ParagraphDialog(Paragraph paragraph)
        {
            Text = "Párrafo";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(570, 470);
            MinimumSize = new Size(520, 430);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            TabControl tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            Controls.Add(tabs);

            TabPage indent = new TabPage("Sangría y espacio");
            tabs.TabPages.Add(indent);

            GroupBox general = new GroupBox();
            general.Text = "General";
            general.Location = new Point(10, 10);
            general.Size = new Size(520, 90);
            indent.Controls.Add(general);
            AddLabel(general, "Alineación:", 12, 25);
            _alignment = new ComboBox();
            _alignment.DropDownStyle = ComboBoxStyle.DropDownList;
            _alignment.Items.AddRange(new string[] { "Izquierda", "Centrada", "Derecha", "Justificada" });
            _alignment.Location = new Point(130, 22);
            _alignment.Width = 180;
            _alignment.SelectedIndex = (int)paragraph.Alignment;
            general.Controls.Add(_alignment);

            AddLabel(general, "Nivel de esquema:", 12, 58);
            _outline = new ComboBox();
            _outline.DropDownStyle = ComboBoxStyle.DropDownList;
            _outline.Items.AddRange(new string[] { "Texto independiente", "Nivel 1", "Nivel 2", "Nivel 3" });
            _outline.Location = new Point(130, 55);
            _outline.Width = 180;
            _outline.SelectedIndex = Math.Max(0, Math.Min(3, paragraph.OutlineLevel));
            general.Controls.Add(_outline);

            GroupBox indentation = new GroupBox();
            indentation.Text = "Sangría";
            indentation.Location = new Point(10, 108);
            indentation.Size = new Size(520, 105);
            indent.Controls.Add(indentation);
            AddLabel(indentation, "Izquierda:", 12, 28);
            _left = Numeric(12, 46, 0, 10, 0.1m);
            _left.Value = TwipsToCm(paragraph.LeftIndentTwips);
            indentation.Controls.Add(_left);
            AddLabel(indentation, "Derecha:", 260, 28);
            _right = Numeric(355, 46, 0, 10, 0.1m);
            _right.Value = TwipsToCm(paragraph.RightIndentTwips);
            indentation.Controls.Add(_right);

            AddLabel(indentation, "Especial:", 12, 70);
            _special = new ComboBox();
            _special.DropDownStyle = ComboBoxStyle.DropDownList;
            _special.Items.AddRange(new string[] { "Ninguna", "Primera línea", "Sangría francesa" });
            _special.Location = new Point(100, 67);
            _special.Width = 150;
            _special.SelectedIndex = paragraph.FirstLineIndentTwips > 0 ? 1 : paragraph.FirstLineIndentTwips < 0 ? 2 : 0;
            indentation.Controls.Add(_special);
            AddLabel(indentation, "En:", 270, 70);
            _by = Numeric(310, 67, 0, 10, 0.1m);
            _by.Value = TwipsToCm(Math.Abs(paragraph.FirstLineIndentTwips));
            indentation.Controls.Add(_by);

            GroupBox spacing = new GroupBox();
            spacing.Text = "Espaciado";
            spacing.Location = new Point(10, 220);
            spacing.Size = new Size(520, 115);
            indent.Controls.Add(spacing);
            AddLabel(spacing, "Anterior:", 12, 28);
            _before = Numeric(100, 25, 0, 100, 1);
            _before.Value = Math.Max(0, paragraph.SpaceBeforeTwips / 20m);
            spacing.Controls.Add(_before);
            AddLabel(spacing, "Posterior:", 12, 70);
            _after = Numeric(100, 67, 0, 100, 1);
            _after.Value = Math.Max(0, paragraph.SpaceAfterTwips / 20m);
            spacing.Controls.Add(_after);
            AddLabel(spacing, "Interlineado:", 250, 28);
            _lineSpacing = new ComboBox();
            _lineSpacing.DropDownStyle = ComboBoxStyle.DropDownList;
            _lineSpacing.Items.AddRange(new string[] { "Múltiple", "Sencillo", "1,5 líneas", "Doble" });
            _lineSpacing.Location = new Point(350, 25);
            _lineSpacing.Width = 150;
            _lineSpacing.SelectedIndex = paragraph.LineSpacing == 1 ? 2 : paragraph.LineSpacing >= 2 ? 3 : 0;
            spacing.Controls.Add(_lineSpacing);
            AddLabel(spacing, "En:", 350, 55);
            NumericUpDown lineAt = Numeric(390, 52, 0, 10, 0.1m);
            lineAt.Value = 1.15m;
            lineAt.Enabled = false;
            spacing.Controls.Add(lineAt);
            _noSpace = new CheckBox();
            _noSpace.Text = "No agregar espacio entre párrafos del mismo estilo";
            _noSpace.Checked = paragraph.NoSpaceBetweenSameStyle;
            _noSpace.Location = new Point(12, 94);
            _noSpace.AutoSize = true;
            spacing.Controls.Add(_noSpace);

            TabPage lines = new TabPage("Líneas y saltos de página");
            tabs.TabPages.Add(lines);
            GroupBox pagination = new GroupBox();
            pagination.Text = "Paginación";
            pagination.Location = new Point(12, 12);
            pagination.Size = new Size(520, 160);
            lines.Controls.Add(pagination);
            _pageBreak = Check("Salto de página antes del párrafo", 16, 28, paragraph.PageBreakBefore, pagination);
            _keepLines = Check("No dividir el párrafo entre páginas", 16, 58, paragraph.KeepLinesTogether, pagination);
            _keepNext = Check("Mantener con el párrafo siguiente", 16, 88, paragraph.KeepWithNext, pagination);
            _widow = Check("Control de viudas y huérfanas", 16, 118, paragraph.WidowControl, pagination);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 48;
            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(382, 10);
            ok.Width = 80;
            bottom.Controls.Add(ok);
            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(468, 10);
            cancel.Width = 80;
            bottom.Controls.Add(cancel);
            Controls.Add(bottom);
            tabs.BringToFront();
            AcceptButton = ok;
            CancelButton = cancel;
        }

        private static NumericUpDown Numeric(int x, int y, decimal min, decimal max, decimal step)
        {
            NumericUpDown box = new NumericUpDown();
            box.Location = new Point(x, y);
            box.Width = 110;
            box.Minimum = min;
            box.Maximum = max;
            box.DecimalPlaces = step < 1 ? 2 : 0;
            box.Increment = step;
            return box;
        }

        private static void AddLabel(Control parent, string text, int x, int y)
        {
            Label label = new Label();
            label.Text = text;
            label.Location = new Point(x, y);
            label.AutoSize = true;
            parent.Controls.Add(label);
        }

        private static CheckBox Check(string text, int x, int y, bool value, Control parent)
        {
            CheckBox box = new CheckBox();
            box.Text = text;
            box.Checked = value;
            box.Location = new Point(x, y);
            box.AutoSize = true;
            parent.Controls.Add(box);
            return box;
        }

        private static decimal TwipsToCm(int twips)
        {
            return Math.Round((decimal)twips / 1440m * 2.54m, 2);
        }

        private static int CmToTwips(decimal cm)
        {
            return (int)Math.Round(cm / 2.54m * 1440m);
        }
    }
}
