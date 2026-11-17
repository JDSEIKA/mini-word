using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Controls;

namespace WordPadClone.Dialogs
{
    internal sealed class NativeFindDialog : Form
    {
        private readonly NativeDocumentEditor _editor;
        private readonly TextBox _find;
        private readonly TextBox _replace;
        private readonly CheckBox _matchCase;
        private readonly Label _status;

        public NativeFindDialog(NativeDocumentEditor editor)
        {
            _editor = editor;
            Text = "Buscar y reemplazar";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(450, 220);

            Label findLabel = new Label();
            findLabel.Text = "Buscar:";
            findLabel.Location = new Point(14, 18);
            findLabel.AutoSize = true;
            Controls.Add(findLabel);

            _find = new TextBox();
            _find.Location = new Point(90, 15);
            _find.Width = 340;
            Controls.Add(_find);

            Label replaceLabel = new Label();
            replaceLabel.Text = "Reemplazar:";
            replaceLabel.Location = new Point(14, 52);
            replaceLabel.AutoSize = true;
            Controls.Add(replaceLabel);

            _replace = new TextBox();
            _replace.Location = new Point(90, 49);
            _replace.Width = 340;
            Controls.Add(_replace);

            _matchCase = new CheckBox();
            _matchCase.Text = "Coincidir mayúsculas y minúsculas";
            _matchCase.Location = new Point(90, 82);
            _matchCase.AutoSize = true;
            Controls.Add(_matchCase);

            _status = new Label();
            _status.Text = "";
            _status.Location = new Point(14, 112);
            _status.Size = new Size(416, 22);
            Controls.Add(_status);

            Button find = new Button();
            find.Text = "Buscar siguiente";
            find.Location = new Point(90, 145);
            find.Width = 120;
            find.Click += delegate { FindNext(); };
            Controls.Add(find);

            Button replace = new Button();
            replace.Text = "Reemplazar";
            replace.Location = new Point(218, 145);
            replace.Width = 100;
            replace.Click += delegate { ReplaceOne(); };
            Controls.Add(replace);

            Button all = new Button();
            all.Text = "Reemplazar todo";
            all.Location = new Point(326, 145);
            all.Width = 104;
            all.Click += delegate { ReplaceAll(); };
            Controls.Add(all);

            Button close = new Button();
            close.Text = "Cerrar";
            close.DialogResult = DialogResult.Cancel;
            close.Location = new Point(326, 180);
            close.Width = 104;
            Controls.Add(close);
            CancelButton = close;
            AcceptButton = find;
            Shown += delegate { _find.Focus(); };
        }

        private void FindNext()
        {
            bool found = _editor.FindNext(_find.Text, _matchCase.Checked, true);
            _status.Text = found ? "Coincidencia encontrada." : "No se encontró el texto.";
        }

        private void ReplaceOne()
        {
            if (_editor.HasSelection && _editor.SelectedText == _find.Text)
            {
                _editor.ReplaceSelection(_replace.Text);
                FindNext();
            }
            else FindNext();
        }

        private void ReplaceAll()
        {
            if (string.IsNullOrEmpty(_find.Text)) return;
            int count = 0;
            _editor.SelectRange(0, 0, 0);
            while (_editor.FindNext(_find.Text, _matchCase.Checked, true))
            {
                _editor.ReplaceSelection(_replace.Text);
                count++;
                if (count > 10000) break;
            }
            _status.Text = count + " reemplazo(s).";
        }
    }
}
