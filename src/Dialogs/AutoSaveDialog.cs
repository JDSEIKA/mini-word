using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class AutoSaveDialog : Form
    {
        private readonly CheckBox _enabled;
        private readonly NumericUpDown _minutes;

        public bool AutoSaveEnabled { get { return _enabled.Checked; } }
        public int IntervalMinutes { get { return (int)_minutes.Value; } }

        public AutoSaveDialog(bool enabled, int minutes)
        {
            Text = "Autoguardado";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(390, 165);

            _enabled = new CheckBox();
            _enabled.Text = "Activar autoguardado";
            _enabled.Checked = enabled;
            _enabled.Location = new Point(18, 20);
            _enabled.AutoSize = true;
            Controls.Add(_enabled);

            Label label = new Label();
            label.Text = "Guardar cada:";
            label.Location = new Point(18, 60);
            label.AutoSize = true;
            Controls.Add(label);

            _minutes = new NumericUpDown();
            _minutes.Minimum = 1;
            _minutes.Maximum = 120;
            _minutes.Value = Math.Max(1, Math.Min(120, minutes));
            _minutes.Location = new Point(120, 57);
            _minutes.Width = 64;
            Controls.Add(_minutes);

            Label suffix = new Label();
            suffix.Text = "minutos";
            suffix.Location = new Point(190, 60);
            suffix.AutoSize = true;
            Controls.Add(suffix);

            Label hint = new Label();
            hint.Text = "Si el documento aún no tiene nombre, se guarda una copia de recuperación.";
            hint.Location = new Point(18, 92);
            hint.Size = new Size(350, 20);
            hint.ForeColor = Color.FromArgb(90, 100, 110);
            Controls.Add(hint);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(206, 124);
            ok.Width = 76;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(292, 124);
            cancel.Width = 76;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
