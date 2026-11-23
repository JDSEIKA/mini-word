using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class FieldDialog : Form
    {
        private readonly ComboBox _kind;

        public FieldDialog()
        {
            Text = "Insertar campo";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(360, 150);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            Label label = new Label();
            label.Text = "Campo:";
            label.Location = new Point(16, 22);
            label.AutoSize = true;
            Controls.Add(label);

            _kind = new ComboBox();
            _kind.DropDownStyle = ComboBoxStyle.DropDownList;
            _kind.Location = new Point(90, 18);
            _kind.Width = 240;
            _kind.Items.Add("Número de página (PAGE)");
            _kind.Items.Add("Total de páginas (NUMPAGES)");
            _kind.Items.Add("Fecha (DATE)");
            _kind.Items.Add("Hora (TIME)");
            _kind.Items.Add("Título del documento (TITLE)");
            _kind.Items.Add("Autor (AUTHOR)");
            _kind.SelectedIndex = 0;
            Controls.Add(_kind);

            Button ok = new Button();
            ok.Text = "Insertar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(166, 88);
            ok.Width = 80;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(252, 88);
            cancel.Width = 80;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        public FieldKind SelectedKind
        {
            get
            {
                switch (_kind.SelectedIndex)
                {
                    case 1: return FieldKind.NumPages;
                    case 2: return FieldKind.Date;
                    case 3: return FieldKind.Time;
                    case 4: return FieldKind.Title;
                    case 5: return FieldKind.Author;
                    default: return FieldKind.Page;
                }
            }
        }
    }
}
