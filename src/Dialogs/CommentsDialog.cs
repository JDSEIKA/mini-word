using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class CommentsDialog : Form
    {
        private readonly ListBox _list;
        private readonly TextBox _detail;
        private readonly TextBox _reply;
        private readonly List<DocumentComment> _comments = new List<DocumentComment>();

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
            _reply = new TextBox();
            _reply.Multiline = true;
            _reply.ScrollBars = ScrollBars.Vertical;
            _reply.Location = new Point(8, 80);
            _reply.Width = 520;
            _reply.Height = 56;
            bottom.Height = 146;
            bottom.Controls.Add(_reply);

            Button reply = new Button();
            reply.Text = "Responder";
            reply.Location = new Point(8, 8);
            reply.Width = 90;
            reply.Click += delegate
            {
                int index = _list.SelectedIndex;
                if (index < 0 || index >= _comments.Count || string.IsNullOrWhiteSpace(_reply.Text)) return;
                _comments[index].Replies.Add(new CommentReply
                {
                    Text = _reply.Text.Trim(),
                    Author = "Tú",
                    Date = DateTime.Now.ToString("s")
                });
                _reply.Clear();
                RefreshLabels();
                ShowSelected(document);
            };
            bottom.Controls.Add(reply);

            Button resolve = new Button();
            resolve.Text = "Resolver";
            resolve.Location = new Point(104, 8);
            resolve.Width = 90;
            resolve.Click += delegate
            {
                int index = _list.SelectedIndex;
                if (index < 0 || index >= _comments.Count) return;
                _comments[index].Resolved = !_comments[index].Resolved;
                RefreshLabels();
                ShowSelected(document);
            };
            bottom.Controls.Add(resolve);
            Button close = new Button();
            close.Text = "Cerrar";
            close.DialogResult = DialogResult.OK;
            close.Location = new Point(448, 108);
            close.Width = 80;
            bottom.Controls.Add(close);
            Controls.Add(bottom);
            AcceptButton = close;
            CancelButton = close;

            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            else _detail.Text = "Este documento no contiene comentarios.";
        }

        private void RefreshLabels()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            for (int i = 0; i < _comments.Count; i++)
            {
                DocumentComment c = _comments[i];
                _list.Items.Add("#" + c.Id + "  " + (string.IsNullOrEmpty(c.Author) ? "Comentario" : c.Author)
                    + (c.Resolved ? "  ✓" : string.Empty) + (c.Replies.Count > 0 ? "  (" + c.Replies.Count + ")" : string.Empty));
            }
            if (_list.Items.Count > 0 && _list.SelectedIndex < 0) _list.SelectedIndex = 0;
            _list.EndUpdate();
        }
        private void ShowSelected(DocumentModel document)
        {
            int index = _list.SelectedIndex;
            if (index < 0 || index >= document.Comments.Count) { _detail.Text = string.Empty; return; }
            DocumentComment comment = document.Comments[index];
            StringBuilder text = new StringBuilder();
            text.Append("Autor: ").Append(comment.Author).Append(Environment.NewLine);
            text.Append("Fecha: ").Append(comment.Date).Append(Environment.NewLine);
            if (comment.Resolved) text.Append("Estado: resuelto").Append(Environment.NewLine);
            text.Append(Environment.NewLine).Append(comment.Text);
            for (int r = 0; r < comment.Replies.Count; r++)
                text.Append(Environment.NewLine).Append(Environment.NewLine)
                    .Append("Respuesta de ").Append(comment.Replies[r].Author).Append(":")
                    .Append(Environment.NewLine).Append(comment.Replies[r].Text);
            _detail.Text = text.ToString();
        }
    }
}
