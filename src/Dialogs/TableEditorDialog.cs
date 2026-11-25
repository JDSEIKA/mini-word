using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class TableEditorDialog : Form
    {
        private readonly DataGridView _grid;
        private readonly Label _hint;
        public string[][] Rows;
        public bool DeleteTable;

        public TableEditorDialog(string[][] source)
        {
            Text = source == null ? "Insertar tabla" : "Editar tabla";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(680, 460);
            MinimumSize = new Size(460, 300);

            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AllowUserToResizeRows = true;
            _grid.RowHeadersWidth = 52;
            _grid.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid.RowHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.FixedSingle;
            _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            _grid.MultiSelect = false;
            _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            Controls.Add(_grid);

            Panel tools = new Panel();
            tools.Dock = DockStyle.Bottom;
            tools.Height = 86;
            Controls.Add(tools);

            Button addRow = MakeButton("Añadir fila", 12, 10);
            addRow.Click += delegate { _grid.Rows.Add(); UpdateLabels(); };
            tools.Controls.Add(addRow);
            Button delRow = MakeButton("Quitar fila", 120, 10);
            delRow.Click += delegate
            {
                if (_grid.CurrentRow != null && _grid.Rows.Count > 1)
                {
                    _grid.Rows.Remove(_grid.CurrentRow);
                    UpdateLabels();
                }
            };
            tools.Controls.Add(delRow);
            Button addCol = MakeButton("Añadir columna", 228, 10);
            addCol.Width = 124;
            addCol.Click += delegate
            {
                int index = _grid.Columns.Add("c" + _grid.Columns.Count, "Columna " + (_grid.Columns.Count + 1));
                _grid.Columns[index].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                UpdateLabels();
            };
            tools.Controls.Add(addCol);
            Button delCol = MakeButton("Quitar columna", 358, 10);
            delCol.Width = 124;
            delCol.Click += delegate
            {
                if (_grid.CurrentCell != null && _grid.Columns.Count > 1)
                {
                    _grid.Columns.RemoveAt(_grid.CurrentCell.ColumnIndex);
                    UpdateLabels();
                }
            };
            tools.Controls.Add(delCol);

            _hint = new Label();
            _hint.Text = "Edite las celdas y use los botones para cambiar filas o columnas.";
            _hint.Location = new Point(12, 48);
            _hint.Size = new Size(420, 22);
            _hint.ForeColor = Color.FromArgb(90, 100, 110);
            tools.Controls.Add(_hint);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(492, 48);
            ok.Width = 78;
            tools.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(578, 48);
            cancel.Width = 78;
            tools.Controls.Add(cancel);

            if (source != null)
            {
                Button delete = new Button();
                delete.Text = "Eliminar tabla";
                delete.Location = new Point(492, 10);
                delete.Width = 164;
                delete.Click += delegate { DeleteTable = true; DialogResult = DialogResult.OK; Close(); };
                tools.Controls.Add(delete);
            }

            AcceptButton = ok;
            CancelButton = cancel;
            Shown += delegate { Fill(source); };
            FormClosing += delegate
            {
                if (DialogResult == DialogResult.OK && !DeleteTable) CaptureRows();
            };
        }

        private static Button MakeButton(string text, int x, int y)
        {
            Button button = new Button();
            button.Text = text;
            button.Location = new Point(x, y);
            button.Width = 104;
            button.Height = 28;
            return button;
        }

        private void Fill(string[][] source)
        {
            int rows = source == null ? 2 : Math.Max(1, source.Length);
            int columns = 2;
            if (source != null)
                for (int i = 0; i < source.Length; i++)
                    if (source[i] != null && source[i].Length > columns) columns = source[i].Length;
            for (int c = 0; c < columns; c++)
                _grid.Columns.Add("c" + c, "Columna " + (c + 1));
            _grid.Rows.Add(rows);
            UpdateLabels();
            if (source != null)
            {
                for (int r = 0; r < source.Length; r++)
                {
                    for (int c = 0; c < columns && c < source[r].Length; c++)
                        _grid.Rows[r].Cells[c].Value = source[r][c];
                }
            }
            if (_grid.Rows.Count > 0 && _grid.Columns.Count > 0)
                _grid.CurrentCell = _grid.Rows[0].Cells[0];
        }

        private void UpdateLabels()
        {
            for (int r = 0; r < _grid.Rows.Count; r++)
                _grid.Rows[r].HeaderCell.Value = (r + 1).ToString();
            _hint.Text = "Filas: " + _grid.Rows.Count + "  ·  Columnas: " + _grid.Columns.Count
                + ". Edite las celdas con los botones inferiores.";
        }

        private void CaptureRows()
        {
            if (_grid.Rows.Count == 0 || _grid.Columns.Count == 0) return;
            Rows = new string[_grid.Rows.Count][];
            for (int r = 0; r < _grid.Rows.Count; r++)
            {
                Rows[r] = new string[_grid.Columns.Count];
                for (int c = 0; c < _grid.Columns.Count; c++)
                    Rows[r][c] = Convert.ToString(_grid.Rows[r].Cells[c].Value) ?? string.Empty;
            }
        }
    }
}
