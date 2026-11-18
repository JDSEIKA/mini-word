using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class HyperlinkDialog : Form
    {
        private readonly TextBox _url;
        private readonly Label _error;

        public string Url { get { return _url.Text.Trim(); } }

        public HyperlinkDialog(string url)
        {
            Text = string.IsNullOrEmpty(url) ? "Insertar hipervínculo" : "Editar hipervínculo";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(470, 150);
            AccessibleName = Text;

            Label label = new Label();
            label.Text = "Destino:";
            label.Location = new Point(16, 18);
            label.AutoSize = true;
            Controls.Add(label);

            _url = new TextBox();
            _url.Text = url ?? string.Empty;
            _url.Location = new Point(16, 40);
            _url.Width = 438;
            _url.AccessibleName = "Destino del hipervínculo";
            Controls.Add(_url);

            Label help = new Label();
            help.Text = "Se abrirá solo después de confirmar el destino.";
            help.Location = new Point(16, 70);
            help.AutoSize = true;
            help.ForeColor = Color.FromArgb(92, 104, 118);
            Controls.Add(help);

            _error = new Label();
            _error.Location = new Point(16, 92);
            _error.Width = 438;
            _error.Height = 18;
            _error.ForeColor = Color.FromArgb(180, 30, 30);
            Controls.Add(_error);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(286, 116);
            ok.Width = 76;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(372, 116);
            cancel.Width = 76;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
            ok.Click += delegate
            {
                string value = _url.Text.Trim();
                Uri uri;
                if (!Uri.TryCreate(value, UriKind.Absolute, out uri)
                    || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps
                        && uri.Scheme != Uri.UriSchemeMailto))
                {
                    _error.Text = "Usa un destino http, https o mailto válido.";
                    DialogResult = DialogResult.None;
                    _url.Focus();
                    _url.SelectAll();
                }
            };
            Shown += delegate { _url.Focus(); _url.SelectAll(); };
        }
    }
}
