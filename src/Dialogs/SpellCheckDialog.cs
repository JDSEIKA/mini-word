using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using WordPadClone.Controls;

namespace WordPadClone.Dialogs
{
    internal sealed class SpellCheckDialog : Form
    {
        private readonly NativeDocumentEditor _editor;
        private readonly Label _word;
        private readonly Label _language;
        private readonly ListBox _suggestions;
        private readonly TextBox _replacement;
        private readonly Label _status;
        private readonly Button _change;
        private readonly Button _ignore;
        private readonly Button _ignoreAll;
        private readonly Button _add;
        private readonly ToolTip _wordTip = new ToolTip();
        private Font _wordFont;
        private SystemSpellChecker _checker;
        private List<SpellError> _errors = new List<SpellError>();
        private int _index;
        private int _version;
        private int _startupTicks;
        private bool _checking;
        private readonly Timer _startupTimer;

        public SpellCheckDialog(NativeDocumentEditor editor)
        {
            _editor = editor;
            Text = "Corrector ortográfico";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(640, 410);
            MinimumSize = new Size(600, 360);

            Label found = new Label();
            found.Text = "Palabra no encontrada:";
            found.Location = new Point(16, 20);
            found.AutoSize = true;
            Controls.Add(found);

            _word = new Label();
            _word.ForeColor = Color.FromArgb(190, 30, 45);
            _word.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
            _word.Location = new Point(170, 12);
            _word.Size = new Size(320, 42);
            _word.AutoSize = false;
            _word.AutoEllipsis = false;
            _word.TextAlign = ContentAlignment.MiddleLeft;
            Controls.Add(_word);

            _language = new Label();
            _language.ForeColor = Color.Gray;
            _language.Location = new Point(508, 20);
            _language.AutoSize = true;
            Controls.Add(_language);

            Label suggestionsLabel = new Label();
            suggestionsLabel.Text = "Sugerencias";
            suggestionsLabel.Location = new Point(16, 60);
            suggestionsLabel.AutoSize = true;
            Controls.Add(suggestionsLabel);

            _suggestions = new ListBox();
            _suggestions.Location = new Point(16, 80);
            _suggestions.Size = new Size(290, 220);
            _suggestions.IntegralHeight = false;
            _suggestions.SelectedIndexChanged += delegate
            {
                if (_suggestions.SelectedItem != null) _replacement.Text = _suggestions.SelectedItem.ToString();
            };
            Controls.Add(_suggestions);

            Label replaceLabel = new Label();
            replaceLabel.Text = "Cambiar por:";
            replaceLabel.Location = new Point(322, 72);
            replaceLabel.AutoSize = true;
            Controls.Add(replaceLabel);

            _replacement = new TextBox();
            _replacement.Location = new Point(322, 92);
            _replacement.Width = 230;
            Controls.Add(_replacement);

            _change = new Button();
            _change.Text = "Cambiar";
            _change.Location = new Point(322, 130);
            _change.Width = 110;
            _change.Click += delegate { ChangeWord(); };
            Controls.Add(_change);

            _ignore = new Button();
            _ignore.Text = "Ignorar";
            _ignore.Location = new Point(442, 130);
            _ignore.Width = 110;
            _ignore.Click += delegate { IgnoreWord(false); };
            Controls.Add(_ignore);

            _ignoreAll = new Button();
            _ignoreAll.Text = "Ignorar todas";
            _ignoreAll.Location = new Point(322, 168);
            _ignoreAll.Width = 110;
            _ignoreAll.Click += delegate { IgnoreWord(true); };
            Controls.Add(_ignoreAll);

            _add = new Button();
            _add.Text = "Añadir al diccionario";
            _add.Location = new Point(427, 168);
            _add.Width = 160;
            _add.Click += delegate { AddWord(); };
            Controls.Add(_add);

            _status = new Label();
            _status.Location = new Point(16, 304);
            _status.Size = new Size(400, 22);
            _status.ForeColor = Color.FromArgb(80, 90, 100);
            Controls.Add(_status);

            Button close = new Button();
            close.Text = "Cerrar";
            close.Location = new Point(472, 330);
            close.Width = 80;
            close.DialogResult = DialogResult.Cancel;
            close.Click += delegate { Close(); };
            Controls.Add(close);
            CancelButton = close;

            _startupTimer = new Timer();
            _startupTimer.Interval = 120;
            _startupTimer.Tick += delegate { TryStart(); };
            Shown += delegate { _startupTimer.Start(); };
            FormClosed += delegate
            {
                _startupTimer.Stop();
                _startupTimer.Dispose();
                _wordTip.Dispose();
                if (_wordFont != null) _wordFont.Dispose();
                if (_checker != null) _checker.Dispose();
                _editor.ClearSpellingHighlight();
            };
            Shown += delegate
            {
                _language.Text = "Idioma: " + _editor.SpellLanguage;
                if (Owner != null)
                {
                    Rectangle area = Screen.FromControl(Owner).WorkingArea;
                    int x = Math.Max(area.Left + 8, Math.Min(area.Right - Width - 8, Owner.Right - Width - 16));
                    int y = Math.Max(area.Top + 8, Owner.Top + 16);
                    Location = new Point(x, y);
                }
            };
        }

        private void TryStart()
        {
            if (_checker == null) _checker = new SystemSpellChecker(_editor.SpellLanguage);
            _startupTicks++;
            if (_checker.IsAvailable)
            {
                _startupTimer.Stop();
                Recheck();
            }
            else if (_startupTicks > 30)
            {
                _startupTimer.Stop();
                _status.Text = "No hay un corrector ortográfico disponible para " + _editor.SpellLanguage + ".";
                SetButtons(false);
            }
            else
            {
                _status.Text = "Iniciando el corrector de Windows...";
            }
        }

        private void Recheck()
        {
            if (_checker == null || !_checker.IsAvailable || _checking) return;
            _checking = true;
            SetButtons(false);
            _status.Text = "Revisando...";
            int version = ++_version;
            _checker.Check(_editor.GetSpellParagraphs(), version, delegate(int resultVersion, List<SpellError> errors)
            {
                if (resultVersion != _version || IsDisposed) return;
                _checking = false;
                _errors = errors ?? new List<SpellError>();
                _index = 0;
                ShowCurrent();
            });
        }

        private void SetWordText(string word)
        {
            if (_wordFont != null) _wordFont.Dispose();
            _wordFont = null;
            string display = word;
            float size = 12f;
            while (size >= 7f)
            {
                Font candidate = new Font("Segoe UI", size, FontStyle.Bold);
                string fitted;
                if (TryFitWord(word, candidate, _word.ClientSize.Width - 6, out fitted))
                {
                    _wordFont = candidate;
                    display = fitted;
                    break;
                }
                candidate.Dispose();
                size -= 0.5f;
            }
            if (_wordFont == null) _wordFont = new Font("Segoe UI", 7f, FontStyle.Bold);
            _word.Font = _wordFont;
            _word.Text = display;
            _wordTip.SetToolTip(_word, word);
        }

        private static bool TryFitWord(string word, Font font, int maxWidth, out string display)
        {
            display = word;
            if (TextRenderer.MeasureText(word, font).Width <= maxWidth) return true;
            for (int split = word.Length / 2; split > 0 && split < word.Length; split--)
            {
                string first = word.Substring(0, split);
                string second = word.Substring(split);
                if (TextRenderer.MeasureText(first, font).Width <= maxWidth
                    && TextRenderer.MeasureText(second, font).Width <= maxWidth)
                {
                    display = first + "\n" + second;
                    return true;
                }
            }
            for (int split = word.Length / 2 + 1; split < word.Length; split++)
            {
                string first = word.Substring(0, split);
                string second = word.Substring(split);
                if (TextRenderer.MeasureText(first, font).Width <= maxWidth
                    && TextRenderer.MeasureText(second, font).Width <= maxWidth)
                {
                    display = first + "\n" + second;
                    return true;
                }
            }
            return false;
        }

        private void ShowCurrent()
        {
            if (_errors == null || _errors.Count == 0)
            {
                _word.Text = string.Empty;
                _suggestions.Items.Clear();
                _replacement.Text = string.Empty;
                _status.Text = "No se encontraron errores ortográficos.";
                SetButtons(false);
                return;
            }
            if (_index < 0) _index = 0;
            if (_index >= _errors.Count) _index = _errors.Count - 1;
            SpellError error = _errors[_index];
            SetWordText(error.Word);
            _editor.SelectSpellingError(error);
            _editor.Refresh();
            _status.Text = "Sugerencia " + (_index + 1) + " de " + _errors.Count;
            SetButtons(true);
            _suggestions.Items.Clear();
            _replacement.Text = string.Empty;
            _checker.Suggest(error.Word, delegate(List<string> suggestions)
            {
                if (IsDisposed) return;
                _suggestions.BeginUpdate();
                for (int i = 0; i < suggestions.Count; i++) _suggestions.Items.Add(suggestions[i]);
                _suggestions.EndUpdate();
                if (_suggestions.Items.Count > 0) _suggestions.SelectedIndex = 0;
            });
        }

        private void ChangeWord()
        {
            if (_errors.Count == 0 || _index < 0 || _index >= _errors.Count) return;
            string replacement = _replacement.Text.Trim();
            if (replacement.Length == 0) return;
            SpellError error = _errors[_index];
            _editor.ReplaceSpelling(error, replacement);
            Recheck();
        }

        private void IgnoreWord(bool all)
        {
            if (_errors.Count == 0 || _index < 0 || _index >= _errors.Count) return;
            SpellError error = _errors[_index];
            string word = error.Word;
            if (_checker != null) _checker.Ignore(word);
            if (all)
            {
                for (int i = _errors.Count - 1; i >= 0; i--)
                    if (string.Equals(_errors[i].Word, word, StringComparison.OrdinalIgnoreCase)) _errors.RemoveAt(i);
                _index = 0;
                ShowCurrent();
            }
            else
            {
                _errors.RemoveAt(_index);
                if (_index >= _errors.Count) _index = Math.Max(0, _errors.Count - 1);
                ShowCurrent();
            }
        }

        private void AddWord()
        {
            if (_errors.Count == 0 || _index < 0 || _index >= _errors.Count) return;
            SpellError error = _errors[_index];
            if (_checker != null) _checker.Add(error.Word);
            _errors.RemoveAt(_index);
            if (_index >= _errors.Count) _index = Math.Max(0, _errors.Count - 1);
            ShowCurrent();
        }

        private void SetButtons(bool enabled)
        {
            _change.Enabled = enabled;
            _ignore.Enabled = enabled;
            _ignoreAll.Enabled = enabled;
            _add.Enabled = enabled;
        }
    }
}
