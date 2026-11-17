using System;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class CommentsDialog : Form
    {
        private readonly ListBox _list;
        private readonly TextBox _detail;

        public CommentsDialog(DocumentModel document)
        {
            Text = "Comentarios";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(540, 330);
            MinimumSize = new Size(420, 260);

            _list = new ListBox();
            _list.Dock = DockStyle.Left;
            _list.Width = 210;
            _list.IntegralHeight = false;
            for (int i = 0; i < document.Comments.Count; i++)
            {
                DocumentComment comment = document.Comments[i];
                _list.Items.Add("#" + comment.Id + "  " + (string.IsNullOrEmpty(comment.Author) ? "Comentario" : comment.Author));
            }
            _list.SelectedIndexChanged += delegate { ShowSelected(document); };
            Controls.Add(_list);

            _detail = new TextBox();
            _detail.Dock = DockStyle.Fill;
            _detail.Multiline = true;
            _detail.ReadOnly = true;
            _detail.ScrollBars = ScrollBars.Vertical;
            _detail.BackColor = Color.White;
            Controls.Add(_detail);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 42;
            Button close = new Button();
            close.Text = "Cerrar";
            close.DialogResult = DialogResult.OK;
            close.Location = new Point(448, 8);
            close.Width = 80;
            bottom.Controls.Add(close);
            Controls.Add(bottom);
            AcceptButton = close;
            CancelButton = close;

            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            else _detail.Text = "Este documento no contiene comentarios.";
        }

        private void ShowSelected(DocumentModel document)
        {
            int index = _list.SelectedIndex;
            if (index < 0 || index >= document.Comments.Count) { _detail.Text = string.Empty; return; }
            DocumentComment comment = document.Comments[index];
            _detail.Text = "Autor: " + comment.Author + Environment.NewLine
                + "Fecha: " + comment.Date + Environment.NewLine + Environment.NewLine
                + comment.Text;
        }
    }
}
