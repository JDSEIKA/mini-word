using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class GoToDialog : Form
    {
        private readonly NumericUpDown _page;

        public int PageNumber
        {
            get { return (int)_page.Value; }
        }

        public GoToDialog(int pageCount)
        {
            Text = "Ir a página";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(260, 100);

            Label label = new Label();
            label.Text = "Número de página:";
            label.Location = new Point(14, 18);
            label.AutoSize = true;
            Controls.Add(label);

            _page = new NumericUpDown();
            _page.Minimum = 1;
            _page.Maximum = Math.Max(1, pageCount);
            _page.Value = 1;
            _page.Location = new Point(130, 15);
            _page.Width = 110;
            Controls.Add(_page);

            Button ok = new Button();
            ok.Text = "Ir";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(78, 58);
            ok.Width = 76;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(160, 58);
            cancel.Width = 80;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
