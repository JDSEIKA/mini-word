using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class NativePageSetupDialog : Form
    {
        private readonly ComboBox _paper;
        private readonly RadioButton _portrait;
        private readonly RadioButton _landscape;
        private readonly NumericUpDown _left;
        private readonly NumericUpDown _right;
        private readonly NumericUpDown _top;
        private readonly NumericUpDown _bottom;

        public int PageWidthTwips;
        public int PageHeightTwips;
        public int MarginLeftTwips;
        public int MarginRightTwips;
        public int MarginTopTwips;
        public int MarginBottomTwips;

        public NativePageSetupDialog(DocumentModel model)
        {
            Text = "Configurar página";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(420, 310);

            Label paperLabel = new Label();
            paperLabel.Text = "Tamaño de papel:";
            paperLabel.Location = new Point(18, 20);
            paperLabel.AutoSize = true;
            Controls.Add(paperLabel);

            _paper = new ComboBox();
            _paper.DropDownStyle = ComboBoxStyle.DropDownList;
            _paper.Location = new Point(150, 17);
            _paper.Width = 240;
            _paper.Items.Add("A4 (210 × 297 mm)");
            _paper.Items.Add("Carta (8,5 × 11 in)");
            _paper.Items.Add("Oficio (8,5 × 14 in)");
            _paper.Items.Add("A5 (148 × 210 mm)");
            _paper.SelectedIndex = 0;
            Controls.Add(_paper);

            GroupBox orientation = new GroupBox();
            orientation.Text = "Orientación";
            orientation.Location = new Point(18, 55);
            orientation.Size = new Size(372, 62);
            Controls.Add(orientation);

            _portrait = new RadioButton();
            _portrait.Text = "Vertical";
            _portrait.Location = new Point(14, 24);
            _portrait.AutoSize = true;
            orientation.Controls.Add(_portrait);

            _landscape = new RadioButton();
            _landscape.Text = "Horizontal";
            _landscape.Location = new Point(150, 24);
            _landscape.AutoSize = true;
            orientation.Controls.Add(_landscape);

            GroupBox margins = new GroupBox();
            margins.Text = "Márgenes (mm)";
            margins.Location = new Point(18, 128);
            margins.Size = new Size(372, 125);
            Controls.Add(margins);

            _left = AddMargin(margins, "Izquierdo:", 14, 24);
            _right = AddMargin(margins, "Derecho:", 190, 24);
            _top = AddMargin(margins, "Superior:", 14, 68);
            _bottom = AddMargin(margins, "Inferior:", 190, 68);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(226, 267);
            ok.Width = 80;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(312, 267);
            cancel.Width = 80;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;

            LoadModel(model);
            ok.Click += delegate { Store(); };
        }

        private static NumericUpDown AddMargin(Control parent, string label, int x, int y)
        {
            Label l = new Label();
            l.Text = label;
            l.Location = new Point(x, y + 3);
            l.AutoSize = true;
            parent.Controls.Add(l);
            NumericUpDown n = new NumericUpDown();
            n.Location = new Point(x + 78, y);
            n.Width = 62;
            n.Minimum = 0;
            n.Maximum = 100;
            n.DecimalPlaces = 1;
            parent.Controls.Add(n);
            return n;
        }

        private void LoadModel(DocumentModel model)
        {
            int w = Math.Min(model.PageWidthTwips, model.PageHeightTwips);
            int h = Math.Max(model.PageWidthTwips, model.PageHeightTwips);
            if (Math.Abs(w - 11906) < 100) _paper.SelectedIndex = 0;
            else if (Math.Abs(w - 12240) < 100 && h > 18000) _paper.SelectedIndex = 2;
            else if (Math.Abs(w - 12240) < 100) _paper.SelectedIndex = 1;
            else if (Math.Abs(w - 8391) < 100) _paper.SelectedIndex = 3;
            _portrait.Checked = !model.IsLandscape;
            _landscape.Checked = model.IsLandscape;
            _left.Value = TwipsToMm(model.MarginLeftTwips);
            _right.Value = TwipsToMm(model.MarginRightTwips);
            _top.Value = TwipsToMm(model.MarginTopTwips);
            _bottom.Value = TwipsToMm(model.MarginBottomTwips);
        }

        private void Store()
        {
            int w = 11906;
            int h = 16838;
            switch (_paper.SelectedIndex)
            {
                case 1: w = 12240; h = 15840; break;
                case 2: w = 12240; h = 20160; break;
                case 3: w = 8391; h = 11906; break;
            }
            if (_landscape.Checked)
            {
                int t = w; w = h; h = t;
            }
            PageWidthTwips = w;
            PageHeightTwips = h;
            MarginLeftTwips = MmToTwips(_left.Value);
            MarginRightTwips = MmToTwips(_right.Value);
            MarginTopTwips = MmToTwips(_top.Value);
            MarginBottomTwips = MmToTwips(_bottom.Value);
        }

        private static int TwipsToMm(int twips)
        {
            return (int)Math.Round(twips / 1440.0 * 25.4);
        }

        private static int MmToTwips(decimal mm)
        {
            return (int)Math.Round((double)mm / 25.4 * 1440.0);
        }
    }
}
