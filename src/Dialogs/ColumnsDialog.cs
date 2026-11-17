using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class ColumnsDialog : Form
    {
        private readonly NumericUpDown _count;
        private readonly NumericUpDown _spacing;
        private readonly CheckBox _separator;

        public int ColumnCount { get { return (int)_count.Value; } }
        public int SpacingTwips { get { return (int)_spacing.Value; } }
        public bool Separator { get { return _separator.Checked; } }

        public ColumnsDialog(int count, int spacing, bool separator)
        {
            Text = "Columnas";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(330, 165);

            Label numberLabel = new Label();
            numberLabel.Text = "Número de columnas:";
            numberLabel.Location = new Point(18, 22);
            numberLabel.AutoSize = true;
            Controls.Add(numberLabel);

            _count = new NumericUpDown();
            _count.Minimum = 1;
            _count.Maximum = 4;
            _count.Value = Math.Max(1, Math.Min(4, count));
            _count.Location = new Point(180, 19);
            _count.Width = 60;
            Controls.Add(_count);

            Label spacingLabel = new Label();
            spacingLabel.Text = "Separación (puntos):";
            spacingLabel.Location = new Point(18, 56);
            spacingLabel.AutoSize = true;
            Controls.Add(spacingLabel);

            _spacing = new NumericUpDown();
            _spacing.Minimum = 6;
            _spacing.Maximum = 144;
            _spacing.Value = Math.Max(6, Math.Min(144, spacing / 20));
            _spacing.Location = new Point(180, 53);
            _spacing.Width = 60;
            Controls.Add(_spacing);

            _separator = new CheckBox();
            _separator.Text = "Línea entre columnas";
            _separator.Checked = separator;
            _separator.Location = new Point(18, 88);
            _separator.AutoSize = true;
            Controls.Add(_separator);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(146, 124);
            ok.Width = 76;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(232, 124);
            cancel.Width = 76;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
