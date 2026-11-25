using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Drawing.Printing;
using System.IO;
using Mutex = System.Threading.Mutex;
using System.Windows.Forms;
using Microsoft.Win32;
using WordPadClone.Controls;
using WordPadClone.Dialogs;
using WordPadClone.Formats;
using WordPadClone.Util;

namespace WordPadClone
{
    /// <summary>
    /// Word-like shell around the native page editor.  All document editing and
    /// rendering is performed by NativeDocumentEditor; no RichEdit control is
    /// involved in the document view or in DOCX serialization.
    /// </summary>
    internal sealed class NativeEditorForm : Form
    {
        private readonly NativeDocumentEditor _editor;
        private readonly RulerBar _ruler;
        private readonly StatusStrip _status;
        private readonly ToolStripStatusLabel _pageLabel;
        private readonly ToolStripStatusLabel _wordLabel;
        private readonly ToolStripStatusLabel _stateLabel;
        private ToolStripComboBox _fontBox;
        private ToolStripComboBox _sizeBox;
        private ToolStripButton _boldButton;
        private ToolStripButton _italicButton;
        private ToolStripButton _underlineButton;
        private ToolStripButton _leftButton;
        private ToolStripButton _centerButton;
        private ToolStripButton _rightButton;
        private RibbonGroup _fontGroup;
        private RibbonGroup _paragraphGroup;
        private RibbonGroup _insertGroup;
        private FlowLayoutPanel _homeFlow;
        private ToolStripButton _spellToggle;
        private SpellCheckDialog _spellDialog;
        private ToolStripButton _rulerButton;
        private ToolStripComboBox _spellLanguageBox;
        private string _path = string.Empty;
        private DocumentFormat _format = DocumentFormat.Docx;
        private bool _dirty;
        private int _printPage;
        private bool _updatingToolbar;
        private Timer _autoSaveTimer;
        private bool _autoSaveEnabled;
        private int _autoSaveMinutes = 5;
        private ToolStripButton _autoSaveButton;
        private ToolStripButton _trackChangesButton;
        private readonly List<string> _recentFiles = new List<string>();
        private RibbonGroup _recentGroup;
        private string _recoveryPath;
        private bool _lossWarningAccepted;
        private static Mutex _recoveryMutex;
        private bool _ownsRecoveryMutex;
        private RunFormat _formatPainterFormat;
        private bool _applyingFormatPainter;
        private FileStream _recoveryLock;

        public NativeEditorForm(string[] args)
        {
            _recoveryMutex = new Mutex(false, "Local\\MiniWord.Recovery");
            try { _ownsRecoveryMutex = _recoveryMutex.WaitOne(0, false); }
            catch (System.Threading.AbandonedMutexException) { _ownsRecoveryMutex = true; }
            Text = "Mini Word " + AppInfo.Version;
            AccessibleName = "Mini Word";
            AccessibleDescription = "Procesador de textos Mini Word";
            Font = new Font(FontResolver.Resolve("Segoe UI", "Español"), 9f);
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1100, 820);
            MinimumSize = new Size(720, 520);
            WindowState = FormWindowState.Maximized;
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            _editor = new NativeDocumentEditor();
            _editor.SpellLanguage = DetectSystemSpellLanguage();
            _editor.Dock = DockStyle.Fill;
            _editor.TextChanged += EditorTextChanged;
            _editor.SelectionChanged += EditorSelectionChanged;
            _editor.PageCountChanged += EditorPageCountChanged;
            _editor.Scroll += delegate { UpdateStatus(); };
            _editor.ViewChanged += delegate { UpdateViewStatus(); };
            Controls.Add(_editor);

            _ruler = new RulerBar(_editor);
            Controls.Add(_ruler);
            LoadRecentFiles();
            Control ribbon = BuildRibbon();
            Controls.Add(ribbon);

            _status = new StatusStrip();
            _status.Dock = DockStyle.Bottom;
            _pageLabel = new ToolStripStatusLabel("Pág. 1/1");
            _wordLabel = new ToolStripStatusLabel("0 palabras");
            _stateLabel = new ToolStripStatusLabel("Listo");
            _stateLabel.Spring = true;
            _stateLabel.TextAlign = ContentAlignment.MiddleRight;
            _status.Items.Add(_pageLabel);
            _status.Items.Add(new ToolStripSeparator());
            _status.Items.Add(_wordLabel);
            _status.Items.Add(_stateLabel);
            Controls.Add(_status);

            _autoSaveTimer = new Timer();
            _autoSaveTimer.Interval = _autoSaveMinutes * 60000;
            _autoSaveTimer.Tick += delegate { AutoSaveNow(); };
            LoadAutoSaveSettings();

            Load += delegate
            {
                bool recovered = TryRestoreRecovery();
                if (!recovered)
                {
                    if (args != null && args.Length > 0 && File.Exists(args[0]))
                        OpenFile(args[0]);
                    else
                        NewDocument();
                }
                FocusEditor();
            };
            FormClosing += NativeFormClosing;
        }

        private Control BuildRibbon()
        {
            Panel host = new Panel();
            host.Dock = DockStyle.Top;
            host.Height = 154;
            host.BackColor = Color.FromArgb(238, 242, 247);
            host.Padding = new Padding(0);

            TabControl tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            tabs.Font = new Font("Segoe UI", 9f);
            tabs.ItemSize = new Size(88, 26);
            tabs.SizeMode = TabSizeMode.Fixed;
            tabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabs.Padding = new Point(14, 4);
            tabs.DrawItem += DrawRibbonTab;
            host.Controls.Add(tabs);

            TabPage file = NewRibbonPage("Archivo");
            FlowLayoutPanel fileFlow = NewRibbonFlow(file);
            fileFlow.Controls.Add(BuildFileTabGroup());
            _recentGroup = BuildRecentGroup();
            fileFlow.Controls.Add(_recentGroup);
            fileFlow.Controls.Add(BuildPrintGroup());
            fileFlow.Controls.Add(BuildAutoSaveGroup());
            tabs.TabPages.Add(file);

            TabPage home = NewRibbonPage("Inicio");
            _homeFlow = NewRibbonFlow(home);
            _homeFlow.Controls.Add(BuildClipboardGroup());
            _fontGroup = BuildFontGroup();
            _homeFlow.Controls.Add(_fontGroup);
            _paragraphGroup = BuildParagraphGroup();
            _homeFlow.Controls.Add(_paragraphGroup);
            _insertGroup = BuildInsertGroup();
            _homeFlow.Controls.Add(_insertGroup);
            tabs.TabPages.Add(home);

            TabPage insert = NewRibbonPage("Insertar");
            FlowLayoutPanel insertFlow = NewRibbonFlow(insert);
            insertFlow.Controls.Add(BuildImageGroup());
            insertFlow.Controls.Add(BuildObjectGroup());
            insertFlow.Controls.Add(BuildHyperlinkGroup());
            insertFlow.Controls.Add(BuildFieldGroup());
            tabs.TabPages.Add(insert);
            TabPage tables = NewRibbonPage("Tablas");
            FlowLayoutPanel tablesFlow = NewRibbonFlow(tables);
            tablesFlow.Controls.Add(BuildTableGroup());
            tabs.TabPages.Add(tables);


            TabPage layout = NewRibbonPage("Disposición");
            FlowLayoutPanel layoutFlow = NewRibbonFlow(layout);
            layoutFlow.Controls.Add(BuildPageGroup());
            layoutFlow.Controls.Add(BuildParagraphDialogGroup());
            layoutFlow.Controls.Add(BuildSectionGroup());
            layoutFlow.Controls.Add(BuildNavigationGroup());
            tabs.TabPages.Add(layout);

            TabPage review = NewRibbonPage("Revisar");
            FlowLayoutPanel reviewFlow = NewRibbonFlow(review);
            reviewFlow.Controls.Add(BuildSpellGroup());
            reviewFlow.Controls.Add(BuildCommentsGroup());
            reviewFlow.Controls.Add(BuildFindGroup());
            reviewFlow.Controls.Add(BuildRevisionGroup());
            tabs.TabPages.Add(review);

            TabPage view = NewRibbonPage("Ver");
            FlowLayoutPanel viewFlow = NewRibbonFlow(view);
            viewFlow.Controls.Add(BuildRulerGroup());
            viewFlow.Controls.Add(BuildZoomGroup());
            viewFlow.Controls.Add(BuildNavigationGroup());
            tabs.TabPages.Add(view);

            TabPage tools = NewRibbonPage("Ayuda");
            FlowLayoutPanel toolsFlow = NewRibbonFlow(tools);
            toolsFlow.Controls.Add(BuildToolsGroup());
            tabs.TabPages.Add(tools);

            tabs.SelectedIndex = 1;
            Resize += delegate { UpdateRibbonVisibility(); };
            UpdateRibbonVisibility();
            return host;
        }

        private void UpdateRibbonVisibility()
        {
            int width = ClientSize.Width;
            if (_fontGroup != null) _fontGroup.Visible = width >= 700;
            if (_paragraphGroup != null) _paragraphGroup.Visible = width >= 920;
            if (_insertGroup != null) _insertGroup.Visible = width >= 1080;
            if (_homeFlow != null) _homeFlow.PerformLayout();
        }

        private static TabPage NewRibbonPage(string text)
        {
            TabPage page = new TabPage(text);
            page.BackColor = Color.FromArgb(249, 250, 252);
            page.Padding = new Padding(0);
            return page;
        }

        private static FlowLayoutPanel NewRibbonFlow(TabPage page)
        {
            FlowLayoutPanel flow = new FlowLayoutPanel();
            flow.Dock = DockStyle.Fill;
            flow.WrapContents = false;
            flow.AutoScroll = true;
            flow.Padding = new Padding(6, 4, 6, 2);
            flow.BackColor = Color.FromArgb(249, 250, 252);
            page.Controls.Add(flow);
            return flow;
        }

        private void DrawRibbonTab(object sender, DrawItemEventArgs e)
        {
            TabControl tabs = (TabControl)sender;
            bool selected = e.Index == tabs.SelectedIndex;
            Rectangle bounds = e.Bounds;
            if (e.Index == 0)
            {
                using (Brush fileBack = new SolidBrush(Color.FromArgb(38, 78, 140)))
                    e.Graphics.FillRectangle(fileBack, bounds);
                TextRenderer.DrawText(e.Graphics, tabs.TabPages[e.Index].Text, tabs.Font, bounds,
                    Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                return;
            }
            using (Brush back = new SolidBrush(selected ? Color.FromArgb(249, 250, 252) : Color.FromArgb(238, 242, 247)))
                e.Graphics.FillRectangle(back, bounds);
            if (selected)
            {
                using (Brush accent = new SolidBrush(Color.FromArgb(38, 78, 140)))
                    e.Graphics.FillRectangle(accent, bounds.Left + 12, bounds.Bottom - 3, Math.Max(1, bounds.Width - 24), 3);
            }
            TextRenderer.DrawText(e.Graphics, tabs.TabPages[e.Index].Text, tabs.Font, bounds,
                selected ? Color.FromArgb(31, 48, 70) : Color.FromArgb(76, 88, 102),
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private RibbonGroup BuildFileTabGroup()
        {
            RibbonGroup group = NewGroup("Documento");
            group.Strip.Items.Add(LargeButton("Nuevo", "new", delegate { NewDocument(); }));
            group.Strip.Items.Add(LargeButton("Abrir", "open", delegate { OpenDialog(); }));
            group.Strip.Items.Add(LargeButton("Guardar", "save", delegate { Save(false); }));
            group.Strip.Items.Add(LargeButton("Guardar como", "save", delegate { Save(true); }));
            group.Strip.Items.Add(LargeButton("Cerrar", "close", delegate { Close(); }));
            return group;
        }

        private RibbonGroup BuildPrintGroup()
        {
            RibbonGroup group = NewGroup("Imprimir");
            group.Strip.Items.Add(LargeButton("Imprimir", "print", delegate { PrintDocument(); }));
            group.Strip.Items.Add(LargeButton("Vista previa", "print", delegate { PrintPreview(); }));
            return group;
        }

        private RibbonGroup BuildAutoSaveGroup()
        {
            RibbonGroup group = NewGroup("Autoguardado");
            _autoSaveButton = LargeButton("Autoguardado", "save", delegate { ConfigureAutoSave(); });
            _autoSaveButton.CheckOnClick = false;
            group.Strip.Items.Add(_autoSaveButton);
            return group;
        }

        private RibbonGroup BuildRecentGroup()
        {
            RibbonGroup group = NewGroup("Recientes");
            PopulateRecentGroup(group);
            return group;
        }

        private void PopulateRecentGroup(RibbonGroup group)
        {
            group.Strip.Items.Clear();
            ToolStripDropDownButton recent = new ToolStripDropDownButton("Abrir reciente");
            recent.Image = RibbonIcons.Get("open");
            recent.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            recent.TextImageRelation = TextImageRelation.ImageAboveText;
            recent.TextAlign = ContentAlignment.BottomCenter;
            recent.AutoSize = true;
            recent.Padding = new Padding(4, 1, 4, 1);
            recent.Font = new Font("Segoe UI", 8.5f);
            recent.ToolTipText = "Últimos 5 documentos";
            if (_recentFiles.Count == 0)
            {
                ToolStripMenuItem empty = new ToolStripMenuItem("Sin documentos recientes");
                empty.Enabled = false;
                recent.DropDownItems.Add(empty);
            }
            else
            {
                for (int i = 0; i < _recentFiles.Count && i < 5; i++)
                {
                    string path = _recentFiles[i];
                    string recentPath = path;
                    string name = Path.GetFileName(path);
                    if (string.IsNullOrEmpty(name)) name = path;
                    if (name.Length > 48) name = name.Substring(0, 45) + "...";
                    name = name.Replace("&", "&&");
                    ToolStripMenuItem item = new ToolStripMenuItem(name);
                    item.ToolTipText = path;
                    item.Click += delegate { OpenRecentFile(recentPath); };
                    recent.DropDownItems.Add(item);
                }
            }
            group.Strip.Items.Add(recent);
        }

        private void OpenRecentFile(string path)
        {
            if (File.Exists(path))
            {
                OpenFile(path);
                return;
            }
            RemoveRecentFile(path);
            MessageBox.Show(this, "El documento reciente ya no existe.",
                "Documentos recientes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RefreshRecentGroup()
        {
            if (_recentGroup != null) PopulateRecentGroup(_recentGroup);
        }

        private void LoadRecentFiles()
        {
            _recentFiles.Clear();
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\MiniWord"))
                {
                    if (key != null)
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            string path = key.GetValue("Recent" + i) as string;
                            if (string.IsNullOrEmpty(path) || !File.Exists(path)) continue;
                            bool exists = false;
                            for (int j = 0; j < _recentFiles.Count; j++)
                                if (string.Equals(_recentFiles[j], path, StringComparison.OrdinalIgnoreCase))
                                    exists = true;
                            if (!exists) _recentFiles.Add(path);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CrashLog.Write("Cargar documentos recientes", ex);
            }
            RefreshRecentGroup();
        }

        private void SaveRecentFiles()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\MiniWord"))
                {
                    if (key == null) return;
                    for (int i = 0; i < 5; i++)
                    {
                        if (i < _recentFiles.Count) key.SetValue("Recent" + i, _recentFiles[i]);
                        else key.DeleteValue("Recent" + i, false);
                    }
                }
            }
            catch (Exception ex)
            {
                CrashLog.Write("Guardar documentos recientes", ex);
            }
        }
        private void AddRecentFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            try
            {
                string full = Path.GetFullPath(path);
                if (!File.Exists(full)) return;
                for (int i = _recentFiles.Count - 1; i >= 0; i--)
                    if (string.Equals(_recentFiles[i], full, StringComparison.OrdinalIgnoreCase))
                        _recentFiles.RemoveAt(i);
                _recentFiles.Insert(0, full);
                while (_recentFiles.Count > 5) _recentFiles.RemoveAt(_recentFiles.Count - 1);
                SaveRecentFiles();
                RefreshRecentGroup();
            }
            catch { }
        }

        private void RemoveRecentFile(string path)
        {
            for (int i = _recentFiles.Count - 1; i >= 0; i--)
                if (string.Equals(_recentFiles[i], path, StringComparison.OrdinalIgnoreCase))
                    _recentFiles.RemoveAt(i);
            SaveRecentFiles();
            RefreshRecentGroup();
        }

        private RibbonGroup BuildClipboardGroup()
        {
            RibbonGroup group = NewGroup("Portapapeles");
            group.Strip.Items.Add(LargeButton("Pegar", "paste", delegate { _editor.Paste(); FocusEditor(); }));
            ToolStripButton plainPaste = LargeButton("Pegar sin formato", "paste", delegate { _editor.PastePlainText(); FocusEditor(); });
            plainPaste.ToolTipText = "Solo texto: sin fuentes, tamaños, negritas ni colores";
            group.Strip.Items.Add(plainPaste);
            group.Strip.Items.Add(IconButton("Cortar", "cut", delegate { _editor.Cut(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Copiar", "copy", delegate { _editor.Copy(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Deshacer", "undo", delegate { _editor.Undo(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Rehacer", "redo", delegate { _editor.Redo(); FocusEditor(); }));
            return group;
        }

        private static string[] SystemFontNames()
        {
            List<string> names = new List<string>();
            try
            {
                using (InstalledFontCollection fonts = new InstalledFontCollection())
                {
                    for (int i = 0; i < fonts.Families.Length; i++)
                    {
                        string name = fonts.Families[i].Name;
                        if (!string.IsNullOrEmpty(name) && !name.StartsWith("@", StringComparison.Ordinal))
                            names.Add(name);
                    }
                }
            }
            catch
            {
                names.AddRange(new string[] { "Calibri", "Arial", "Segoe UI", "Tahoma" });
            }
            names.Sort(StringComparer.CurrentCultureIgnoreCase);
            string[] preferred = new string[] { "Calibri", "Arial", "Times New Roman", "Cambria", "Georgia", "Verdana", "Courier New", "Segoe UI", "Garamond", "Tahoma" };
            List<string> ordered = new List<string>();
            for (int i = 0; i < preferred.Length; i++)
            {
                int index = names.FindIndex(delegate(string value)
                {
                    return string.Equals(value, preferred[i], StringComparison.OrdinalIgnoreCase);
                });
                if (index >= 0)
                {
                    ordered.Add(names[index]);
                    names.RemoveAt(index);
                }
            }
            ordered.AddRange(names);
            return ordered.ToArray();
        }

        private RibbonGroup BuildFontGroup()
        {
            RibbonGroup group = NewGroup("Fuente");
            _fontBox = new ToolStripComboBox("font");
            _fontBox.AutoSize = false;
            _fontBox.Width = 170;
            _fontBox.DropDownWidth = 280;
            _fontBox.MaxDropDownItems = 30;
            _fontBox.DropDownStyle = ComboBoxStyle.DropDown;
            _fontBox.Items.AddRange(SystemFontNames());
            _fontBox.Text = "Calibri";
            _fontBox.SelectedIndexChanged += delegate { ApplyFontFamily(); };
            _fontBox.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { ApplyFontFamily(); e.SuppressKeyPress = true; } };
            group.Strip.Items.Add(_fontBox);

            _sizeBox = new ToolStripComboBox("size");
            _sizeBox.AutoSize = false;
            _sizeBox.Width = 46;
            for (int i = 8; i <= 72; i++) _sizeBox.Items.Add(i.ToString());
            _sizeBox.Text = "11";
            _sizeBox.SelectedIndexChanged += delegate { ApplyFontSize(); };
            _sizeBox.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { ApplyFontSize(); e.SuppressKeyPress = true; } };
            group.Strip.Items.Add(_sizeBox);

            _boldButton = IconToggle("Negrita", "bold", delegate { _editor.ToggleBold(); SyncToolbar(); if (_editor.IsEditingTableCell) _editor.FocusCellEditor(); else FocusEditor(); });
            _italicButton = IconToggle("Cursiva", "italic", delegate { _editor.ToggleItalic(); SyncToolbar(); if (_editor.IsEditingTableCell) _editor.FocusCellEditor(); else FocusEditor(); });
            _underlineButton = IconToggle("Subrayado", "underline", delegate { _editor.ToggleUnderline(); SyncToolbar(); if (_editor.IsEditingTableCell) _editor.FocusCellEditor(); else FocusEditor(); });
            group.Strip.Items.Add(_boldButton);
            group.Strip.Items.Add(_italicButton);
            group.Strip.Items.Add(_underlineButton);
            group.Strip.Items.Add(IconButton("Color del texto", "color", delegate { ChooseColor(); }));
            return group;
        }

        private RibbonGroup BuildParagraphGroup()
        {
            RibbonGroup group = NewGroup("Párrafo");
            _leftButton = IconToggle("Izquierda", "left", delegate { _editor.SetAlignment(ParagraphAlignment.Left); SyncToolbar(); FocusEditor(); });
            _centerButton = IconToggle("Centrar", "center", delegate { _editor.SetAlignment(ParagraphAlignment.Center); SyncToolbar(); FocusEditor(); });
            _rightButton = IconToggle("Derecha", "right", delegate { _editor.SetAlignment(ParagraphAlignment.Right); SyncToolbar(); FocusEditor(); });
            ToolStripButton justify = IconButton("Justificar", "justify", delegate { _editor.SetAlignment(ParagraphAlignment.Justify); SyncToolbar(); FocusEditor(); });
            group.Strip.Items.Add(_leftButton);
            group.Strip.Items.Add(_centerButton);
            group.Strip.Items.Add(_rightButton);
            group.Strip.Items.Add(justify);
            group.Strip.Items.Add(IconButton("Aumentar sangría", "indent", delegate { _editor.IncreaseIndent(); SyncToolbar(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Disminuir sangría", "outdent", delegate { _editor.DecreaseIndent(); SyncToolbar(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Viñeta", "bullet", delegate { _editor.ToggleBullet(); SyncToolbar(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Numeración", "numbered", delegate { _editor.ToggleNumberedList(); SyncToolbar(); FocusEditor(); }));
            ToolStripDropDownButton listLevels = new ToolStripDropDownButton("Lista");
            listLevels.Image = RibbonIcons.Get("list-level");
            listLevels.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            listLevels.TextImageRelation = TextImageRelation.ImageAboveText;
            listLevels.TextAlign = ContentAlignment.BottomCenter;
            listLevels.ToolTipText = "Nivel de lista";
            listLevels.DropDownItems.Add(new ToolStripMenuItem("Aumentar nivel", null, delegate { _editor.IncreaseListLevel(); SyncToolbar(); FocusEditor(); }));
            listLevels.DropDownItems.Add(new ToolStripMenuItem("Disminuir nivel", null, delegate { _editor.DecreaseListLevel(); SyncToolbar(); FocusEditor(); }));
            group.Strip.Items.Add(listLevels);
            ToolStripDropDownButton styles = new ToolStripDropDownButton("Estilos");
            styles.Image = RibbonIcons.Get("paragraph");
            styles.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            styles.TextImageRelation = TextImageRelation.ImageAboveText;
            styles.TextAlign = ContentAlignment.BottomCenter;
            styles.ToolTipText = "Aplicar estilo de párrafo";
            for (int si = 0; si < DocumentStyles.Names.Length; si++)
            {
                string styleName = DocumentStyles.Names[si];
                ToolStripMenuItem item = new ToolStripMenuItem(styleName);
                item.Click += delegate { _editor.ApplyParagraphStyle(styleName); SyncToolbar(); FocusEditor(); };
                styles.DropDownItems.Add(item);
            }
            group.Strip.Items.Add(styles);
            return group;
        }

        private RibbonGroup BuildTableGroup()
        {
            RibbonGroup group = NewGroup("Tabla");
            group.Strip.Items.Add(LargeButton("Insertar tabla", "table", delegate { InsertTable(); }));
            group.Strip.Items.Add(LargeButton("Editar tabla", "table", delegate { EditTableAtCaret(); }));
            group.Strip.Items.Add(LargeButton("Agregar fila", "addrow", delegate { if (!_editor.AddTableRow()) MessageBox.Show(this, "Seleccione una tabla.", "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information); FocusEditor(); }));
            group.Strip.Items.Add(LargeButton("Eliminar fila", "removerow", delegate { if (!_editor.RemoveTableRow()) MessageBox.Show(this, "Seleccione una tabla simple, sin celdas combinadas.", "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information); FocusEditor(); }));
            group.Strip.Items.Add(LargeButton("Fusionar celdas", "merge", delegate
            {
                if (_editor.CanMergeTableCells) _editor.MergeSelectedTableCells();
                else MessageBox.Show(this, "Seleccione varias celdas con el ratón para fusionarlas.",
                    "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }));
            group.Strip.Items.Add(LargeButton("Dividir celda", "split", delegate { _editor.SplitSelectedTableCell(); }));
            group.Strip.Items.Add(LargeButton("Añadir columna", "addcolumn", delegate { if (!_editor.AddTableColumn()) MessageBox.Show(this, "Seleccione una tabla.", "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information); FocusEditor(); }));
            group.Strip.Items.Add(LargeButton("Quitar columna", "removecolumn", delegate { if (!_editor.RemoveTableColumn()) MessageBox.Show(this, "Seleccione una tabla simple, sin celdas combinadas.", "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information); FocusEditor(); }));
            group.Strip.Items.Add(LargeButton("Repetir encabezado", "table", delegate { if (!_editor.ToggleRepeatTableHeader()) MessageBox.Show(this, "Seleccione una tabla.", "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information); FocusEditor(); }));
            return group;
        }

        private RibbonGroup BuildImageGroup()
        {
            RibbonGroup group = NewGroup("Imagen y fecha");
            group.Strip.Items.Add(LargeButton("Imagen", "image", delegate { InsertImage(); }));
            group.Strip.Items.Add(LargeButton("Fecha", "calendar", delegate { InsertDate(); }));
            group.Strip.Items.Add(LargeButton("Ajuste de texto", "wrap", delegate
            {
                if (!_editor.ToggleSelectedImageWrap())
                {
                    MessageBox.Show(this, "Seleccione primero una imagen insertada.", "Ajuste de texto",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                FocusEditor();
            }));
            return group;
        }

        private RibbonGroup BuildObjectGroup()
        {
            RibbonGroup group = NewGroup("Objetos");
            group.Strip.Items.Add(LargeButton("Cuadro de texto", "textbox", delegate
            {
                _editor.InsertTextBox();
                ActiveControl = _editor;
                _editor.Select();
                FocusEditor();
                BeginInvoke(new Action(delegate { ActiveControl = _editor; _editor.Select(); FocusEditor(); }));
            }));
            group.Strip.Items.Add(LargeButton("Formato de texto", "textbox", delegate
            {
                if (!_editor.EditSelectedTextBox())
                {
                    MessageBox.Show(this, "Seleccione primero un cuadro de texto.", "Formato de texto",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                FocusEditor();
            }));
            group.Strip.Items.Add(LargeButton("Forma", "shape", delegate
            {
                using (ShapePickerDialog dialog = new ShapePickerDialog())
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    _editor.InsertShape(dialog.SelectedKind);
                }
                ActiveControl = _editor;
                _editor.Select();
                FocusEditor();
                BeginInvoke(new Action(delegate { ActiveControl = _editor; _editor.Select(); FocusEditor(); }));
            }));
            group.Strip.Items.Add(LargeButton("Color de forma", "shapecolor", delegate
            {
                if (!_editor.EditSelectedShape())
                {
                    MessageBox.Show(this, "Seleccione primero una forma insertada.", "Formato de forma",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                FocusEditor();
            }));
            group.Strip.Items.Add(LargeButton("Traer al frente", "shape", delegate
            {
                if (!_editor.BringSelectedFloatingToFront()) FocusEditor();
            }));
            group.Strip.Items.Add(LargeButton("Enviar al fondo", "shape", delegate
            {
                if (!_editor.SendSelectedFloatingToBack()) FocusEditor();
            }));
            group.Strip.Items.Add(LargeButton("Rotar 90°", "shape", delegate
            {
                if (!_editor.RotateSelectedFloating(90)) FocusEditor();
            }));
            return group;
        }

        private RibbonGroup BuildInsertGroup()
        {
            RibbonGroup group = NewGroup("Insertar");
            group.Strip.Items.Add(LargeButton("Encabezado y pie", "header", delegate { HeaderFooter(); }));
            group.Strip.Items.Add(LargeButton("Salto de página", "pagebreak", delegate { _editor.InsertPageBreak(); FocusEditor(); }));
            return group;
        }
        private RibbonGroup BuildFieldGroup()
        {
            RibbonGroup group = NewGroup("Campos e índice");
            group.Strip.Items.Add(LargeButton("Campo", "page", delegate { InsertFieldDialog(); }));
            group.Strip.Items.Add(LargeButton("Índice", "paragraph", delegate { InsertToc(); }));
            group.Strip.Items.Add(LargeButton("Actualizar índice", "find", delegate { UpdateToc(); }));
            return group;
        }

        private void InsertFieldDialog()
        {
            using (FieldDialog dialog = new FieldDialog())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _editor.InsertField(dialog.SelectedKind);
                _dirty = true;
                UpdateTitle();
                FocusEditor();
            }
        }

        private void InsertToc()
        {
            if (!_editor.InsertTableOfContents(3))
            {
                MessageBox.Show(this, "El documento no tiene títulos con nivel de esquema.", "Índice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void UpdateToc()
        {
            if (!_editor.UpdateTableOfContents(3))
            {
                MessageBox.Show(this, "Inserte primero un índice.", "Índice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void SyncTrackChangesButton()
        {
            if (_trackChangesButton != null) _trackChangesButton.Checked = _editor != null && _editor.TrackChanges;
        }
        private RibbonGroup BuildRevisionGroup()
        {
            RibbonGroup group = NewGroup("Cambios");
            _trackChangesButton = LargeButton("Control de cambios", "comment", delegate
            {
                bool enabled = !_editor.TrackChanges;
                _editor.SetTrackChanges(enabled);
                _trackChangesButton.Checked = enabled;
                _status.Text = enabled ? "Control de cambios activado" : "Control de cambios desactivado";
                FocusEditor();
            });
            _trackChangesButton.CheckOnClick = false;
            group.Strip.Items.Add(_trackChangesButton);
            group.Strip.Items.Add(LargeButton("Aceptar cambio", "check", delegate { AcceptOneRevision(); }));
            group.Strip.Items.Add(LargeButton("Rechazar cambio", "close", delegate { RejectOneRevision(); }));
            group.Strip.Items.Add(LargeButton("Aceptar todo", "check", delegate { AcceptRevisions(); }));
            group.Strip.Items.Add(LargeButton("Rechazar todo", "close", delegate { RejectRevisions(); }));
            group.Strip.Items.Add(LargeButton("Comparar documentos", "find", delegate { CompareDocuments(); }));
            return group;
        }

        private void AcceptOneRevision()
        {
            if (!_editor.AcceptRevisionAtCaret())
                MessageBox.Show(this, "Coloque el cursor sobre un cambio.", "Control de cambios",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void RejectOneRevision()
        {
            if (!_editor.RejectRevisionAtCaret())
                MessageBox.Show(this, "Coloque el cursor sobre un cambio.", "Control de cambios",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }
        private void AcceptRevisions()
        {
            if (!_editor.AcceptAllRevisions())
            {
                MessageBox.Show(this, "No hay cambios pendientes.", "Control de cambios",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void RejectRevisions()
        {
            if (!_editor.RejectAllRevisions())
            {
                MessageBox.Show(this, "No hay cambios pendientes.", "Control de cambios",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void CompareDocuments()
        {
            if (!ConfirmDiscard()) return;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Comparar con documento";
                dialog.Filter = DocumentFormats.Filter;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    DocumentModel other = DocumentFormats.ReadModel(dialog.FileName,
                        DocumentFormats.Detect(dialog.FileName));
                    List<string> formatNotes = new List<string>();
                    DocumentModel compared = DocumentCompare.Compare(_editor.Document, other, formatNotes);
                    _editor.Document = compared;
                    _path = string.Empty;
                    _format = DocumentFormat.Docx;
                    _dirty = true;
                    _lossWarningAccepted = false;
                    UpdateTitle();
                    UpdateStatus();
                    _status.Text = "Documento comparado con cambios controlados"
                        + (formatNotes.Count > 0 ? "; " + formatNotes.Count + " cambios de formato" : string.Empty);
                    if (formatNotes.Count > 0)
                        MessageBox.Show(this, "La comparación detectó " + formatNotes.Count
                            + " cambios de formato. Se aplicaron como formato del documento nuevo.",
                            "Comparar documentos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    FocusEditor();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No se pudo comparar el documento.\n\n" + ex.Message,
                        "Comparar documentos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private RibbonGroup BuildHyperlinkGroup()
        {
            RibbonGroup group = NewGroup("Insertar");
            group.Strip.Items.Add(LargeButton("Hipervínculo", "link", delegate { EditHyperlink(); }));
            return group;
        }

        private void EditHyperlink()
        {
            using (HyperlinkDialog dialog = new HyperlinkDialog(_editor.SelectedHyperlinkUrl))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                if (_editor.ApplyHyperlink(dialog.Url))
                {
                    _dirty = true;
                    UpdateTitle();
                }
                FocusEditor();
            }
        }

        private RibbonGroup BuildPageGroup()
        {
            RibbonGroup group = NewGroup("Página");
            group.Strip.Items.Add(LargeButton("Configurar página", "setup", delegate { PageSetup(); }));
            return group;
        }

        private RibbonGroup BuildRulerGroup()
        {
            RibbonGroup group = NewGroup("Regla");
            _rulerButton = LargeButton("Regla", "ruler", delegate
            {
                _ruler.Visible = !_ruler.Visible;
                _rulerButton.Checked = _ruler.Visible;
                FocusEditor();
            });
            _rulerButton.CheckOnClick = false;
            _rulerButton.Checked = _ruler.Visible;
            group.Strip.Items.Add(_rulerButton);
            return group;
        }

        private RibbonGroup BuildZoomGroup()
        {
            RibbonGroup group = NewGroup("Zoom");
            group.Strip.Items.Add(IconButton("Acercar", "zoom-in", delegate { SetZoom(_editor.Zoom + 0.1f); }));
            group.Strip.Items.Add(IconButton("Alejar", "zoom-out", delegate { SetZoom(_editor.Zoom - 0.1f); }));
            group.Strip.Items.Add(IconButton("100%", "zoom-reset", delegate { SetZoom(1f); }));
            return group;
        }

        private RibbonGroup BuildParagraphDialogGroup()
        {
            RibbonGroup group = NewGroup("Párrafo");
            group.Strip.Items.Add(LargeButton("Párrafo...", "paragraph", delegate { EditParagraph(); }));
            return group;
        }

        private void EditParagraph()
        {
            using (ParagraphDialog dialog = new ParagraphDialog(_editor.CurrentParagraph))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _editor.ApplyParagraphFormatting(dialog.Alignment, dialog.LeftIndentTwips,
                    dialog.RightIndentTwips, dialog.FirstLineIndentTwips,
                    dialog.SpaceBeforeTwips, dialog.SpaceAfterTwips, dialog.LineSpacingValue,
                    dialog.PageBreakBefore, dialog.NoSpaceSameStyle, dialog.OutlineLevel,
                    dialog.KeepLinesTogether, dialog.KeepWithNext, dialog.WidowControl);
                _dirty = true;
                UpdateTitle();
                FocusEditor();
            }
        }

        private RibbonGroup BuildSectionGroup()
        {
            RibbonGroup group = NewGroup("Sección");
            group.Strip.Items.Add(LargeButton("Columnas", "columns", delegate
            {
                DocumentSection section = _editor.Document.SectionForParagraph(_editor.CaretParagraph);
                using (ColumnsDialog dialog = new ColumnsDialog(section.ColumnCount, section.ColumnSpacingTwips, section.ColumnsSeparator))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    _editor.SetCurrentSectionColumns(dialog.ColumnCount, dialog.SpacingTwips, dialog.Separator);
                    _dirty = true;
                    UpdateTitle();
                }
                FocusEditor();
            }));
            group.Strip.Items.Add(LargeButton("Salto de sección", "section", delegate
            {
                _editor.InsertSectionBreak();
                _dirty = true;
                UpdateTitle();
                FocusEditor();
            }));
            return group;
        }

        private RibbonGroup BuildNavigationGroup()
        {
            RibbonGroup group = NewGroup("Navegación");
            group.Strip.Items.Add(IconButton("Página anterior", "previous", delegate { _editor.GoToPreviousPage(); FocusEditor(); }));
            group.Strip.Items.Add(IconButton("Página siguiente", "next", delegate { _editor.GoToNextPage(); FocusEditor(); }));
            return group;
        }

        private RibbonGroup BuildToolsGroup()
        {
            RibbonGroup group = NewGroup("Ayuda");
            group.Strip.Items.Add(LargeButton("Ayuda", "help", delegate { ShowHelp(); }));
            group.Strip.Items.Add(LargeButton("Acerca de", "info", delegate { ShowAbout(); }));
            return group;
        }

        private void ShowSpellCheck()
        {
            _editor.SpellCheckEnabled = true;
            if (_spellToggle != null) _spellToggle.Checked = true;
            if (_spellDialog != null && !_spellDialog.IsDisposed)
            {
                _spellDialog.Activate();
                return;
            }
            _spellDialog = new SpellCheckDialog(_editor);
            _spellDialog.FormClosed += delegate
            {
                _spellDialog = null;
                FocusEditor();
            };
            _spellDialog.Show(this);
        }

        private void ShowHelp()
        {
            using (HelpDialog dialog = new HelpDialog())
                dialog.ShowDialog(this);
        }

        private void ShowAbout()
        {
            MessageBox.Show(this,
                "Mini Word " + AppInfo.Version + "\nEditor nativo de documentos DOCX.\n"
                + "Motor tipográfico y de paginación propio.\n"
                + "Secciones, columnas, objetos flotantes, notas, comentarios y estilos de tabla.",
                "Acerca de Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private RibbonGroup BuildCommentsGroup()
        {
            RibbonGroup group = NewGroup("Comentarios");
            group.Strip.Items.Add(LargeButton("Comentarios", "comment", delegate
            {
                using (CommentsDialog dialog = new CommentsDialog(_editor.Document))
                {
                    dialog.ShowDialog(this);
                    _dirty = true;
                    UpdateTitle();
                }
            }));
            return group;
        }

        private RibbonGroup BuildFindGroup()
        {
            RibbonGroup group = NewGroup("Edición");
            group.Strip.Items.Add(LargeButton("Buscar y reemplazar", "find", delegate { FindReplace(); }));
            return group;
        }

        private static readonly string[] KnownSpellLanguages =
            new string[] { "es-MX", "es-ES", "en-US", "en-GB", "fr-FR", "de-DE" };

        private static string DetectSystemSpellLanguage()
        {
            CultureInfo culture = CultureInfo.CurrentUICulture;
            if (culture == null || string.IsNullOrEmpty(culture.Name) || culture.Name == "iv")
                culture = CultureInfo.InstalledUICulture;
            string name = culture == null ? string.Empty : culture.Name;
            if (string.IsNullOrEmpty(name)) return "es-MX";

            for (int i = 0; i < KnownSpellLanguages.Length; i++)
                if (string.Equals(KnownSpellLanguages[i], name, StringComparison.OrdinalIgnoreCase))
                    return KnownSpellLanguages[i];

            // Windows can expose a neutral culture (for example "es" or
            // "en").  Prefer the matching regional variant when one exists.
            if (name.StartsWith("es", StringComparison.OrdinalIgnoreCase)) return "es-ES";
            if (name.StartsWith("en", StringComparison.OrdinalIgnoreCase)) return "en-US";
            if (name.StartsWith("fr", StringComparison.OrdinalIgnoreCase)) return "fr-FR";
            if (name.StartsWith("de", StringComparison.OrdinalIgnoreCase)) return "de-DE";

            // Keep the real system language visible even if this build does
            // not ship an explicit entry for it.  The Windows spell checker
            // will report availability at review time.
            return name;
        }

        private RibbonGroup BuildSpellGroup()
        {
            RibbonGroup group = NewGroup("Ortografía");
            _spellToggle = LargeButton("Activar", "find", delegate
            {
                _editor.SpellCheckEnabled = !_editor.SpellCheckEnabled;
                _spellToggle.Checked = _editor.SpellCheckEnabled;
                FocusEditor();
            });
            _spellToggle.Checked = true;
            group.Strip.Items.Add(_spellToggle);

            _spellLanguageBox = new ToolStripComboBox("language");
            _spellLanguageBox.AutoSize = false;
            _spellLanguageBox.Width = 112;
            _spellLanguageBox.DropDownStyle = ComboBoxStyle.DropDownList;
            List<string> languages = new List<string>(KnownSpellLanguages);
            string systemLanguage = _editor.SpellLanguage;
            if (!languages.Contains(systemLanguage))
                languages.Insert(0, systemLanguage);
            _spellLanguageBox.Items.AddRange(languages.ToArray());
            _spellLanguageBox.SelectedItem = systemLanguage;
            if (_spellLanguageBox.SelectedIndex < 0) _spellLanguageBox.SelectedIndex = 0;
            _spellLanguageBox.SelectedIndexChanged += delegate
            {
                if (_spellLanguageBox.Text.Length > 0) _editor.SpellLanguage = _spellLanguageBox.Text;
                FocusEditor();
            };
            group.Strip.Items.Add(_spellLanguageBox);
            group.Strip.Items.Add(LargeButton("Ortografía", "spell", delegate { ShowSpellCheck(); }));
            return group;
        }

        private RibbonGroup NewGroup(string caption)
        {
            return new RibbonGroup(caption);
        }

        private ToolStripButton LargeButton(string text, string icon, EventHandler click)
        {
            ToolStripButton button = new ToolStripButton(text);
            button.Image = RibbonIcons.Get(icon);
            button.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            button.TextImageRelation = TextImageRelation.ImageAboveText;
            button.TextAlign = ContentAlignment.BottomCenter;
            button.AutoSize = true;
            button.ImageAlign = ContentAlignment.TopCenter;
            button.Padding = new Padding(4, 1, 4, 1);
            button.Font = new Font("Segoe UI", 8.5f);
            button.ToolTipText = text;
            button.Click += click;
            return button;
        }

        private ToolStripButton IconButton(string tooltip, string icon, EventHandler click)
        {
            ToolStripButton button = new ToolStripButton();
            button.Image = RibbonIcons.Get(icon);
            button.DisplayStyle = ToolStripItemDisplayStyle.Image;
            button.AutoSize = false;
            button.Size = new Size(38, 36);
            button.ImageAlign = ContentAlignment.MiddleCenter;
            button.ToolTipText = tooltip;
            button.AccessibleName = tooltip;
            button.Click += click;
            return button;
        }

        private ToolStripButton IconToggle(string tooltip, string icon, EventHandler click)
        {
            ToolStripButton button = IconButton(tooltip, icon, click);
            button.CheckOnClick = false;
            return button;
        }

        private void EditorTextChanged(object sender, EventArgs e)
        {
            _dirty = true;
            UpdateTitle();
            UpdateStatus();
        }

        private void EditorSelectionChanged(object sender, EventArgs e)
        {
            SyncToolbar();
            UpdateStatus();
            if (_formatPainterFormat != null && !_applyingFormatPainter && _editor.HasSelection)
            {
                _applyingFormatPainter = true;
                try
                {
                    RunFormat copied = _formatPainterFormat;
                    _formatPainterFormat = null;
                    _editor.ApplySelectionFormat(copied);
                    _status.Text = "Formato aplicado";
                }
                finally { _applyingFormatPainter = false; }
            }
        }

        private void EditorPageCountChanged(object sender, EventArgs e)
        {
            UpdateStatus();
        }

        private void SyncToolbar()
        {
            if (_editor == null) return;
            _updatingToolbar = true;
            try
            {
                RunFormat format = _editor.IsEditingTableCell && _editor.CellEditorFormat != null
                    ? _editor.CellEditorFormat : _editor.CurrentFormat;
                _fontBox.Text = format.FontFamily;
                _sizeBox.Text = (format.SizeHalfPoints / 2f).ToString("0.#");
                _boldButton.Checked = format.Bold;
                _italicButton.Checked = format.Italic;
                _underlineButton.Checked = format.Underline;
                ParagraphAlignment alignment = _editor.CurrentAlignment;
                _leftButton.Checked = alignment == ParagraphAlignment.Left;
                _centerButton.Checked = alignment == ParagraphAlignment.Center;
                _rightButton.Checked = alignment == ParagraphAlignment.Right;
            }
            finally
            {
                _updatingToolbar = false;
            }
        }

        private void FocusEditor()
        {
            if (_editor != null && _editor.IsEditingTableCell) _editor.FocusCellEditor();
            else if (_editor != null) _editor.Focus();
        }
        private void ChangeFontSize(float delta)
        {
            RunFormat format = _editor.CurrentFormat;
            float points = Math.Max(1f, format.SizeHalfPoints / 2f + delta);
            _editor.SetFontSize(points);
            SyncToolbar();
            FocusEditor();
        }
        private void ApplyFontFamily()
        {
            if (_updatingToolbar) return;
            _editor.SetFontFamily(_fontBox.Text);
            if (_editor.IsEditingTableCell) _editor.FocusCellEditor(); else FocusEditor();
        }

        private void ApplyFontSize()
        {
            if (_updatingToolbar) return;
            float size;
            if (float.TryParse(_sizeBox.Text, out size)) _editor.SetFontSize(size);
            if (_editor.IsEditingTableCell) _editor.FocusCellEditor(); else FocusEditor();
        }

        private void UpdateViewStatus()
        {
            if (_editor == null) return;
            _pageLabel.Text = "Pág. " + (_editor.ViewedPage + 1) + "/" + _editor.PageCount;
        }

        private void UpdateStatus()
        {
            if (_editor == null) return;
            _pageLabel.Text = "Pág. " + (_editor.ViewedPage + 1) + "/" + _editor.PageCount;
            string text = _editor.Document.PlainText;
            int words = 0;
            bool inWord = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsWhiteSpace(text[i])) inWord = false;
                else if (!inWord) { words++; inWord = true; }
            }
            _wordLabel.Text = words + (words == 1 ? " palabra" : " palabras");
            _stateLabel.Text = _dirty ? "Modificado" : "Listo";
        }

        private void LoadAutoSaveSettings()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\MiniWord"))
                {
                    if (key != null)
                    {
                        object enabled = key.GetValue("AutoSaveEnabled");
                        object minutes = key.GetValue("AutoSaveMinutes");
                        if (enabled != null) _autoSaveEnabled = Convert.ToInt32(enabled) != 0;
                        if (minutes != null) _autoSaveMinutes = Math.Max(1, Math.Min(120, Convert.ToInt32(minutes)));
                    }
                }
            }
            catch { }
            if (_autoSaveButton != null) _autoSaveButton.Checked = _autoSaveEnabled;
            if (_autoSaveTimer != null) _autoSaveTimer.Interval = _autoSaveMinutes * 60000;
            if (_autoSaveTimer != null)
            {
                _autoSaveTimer.Stop();
                _autoSaveTimer.Start();
            }
        }

        private void SaveAutoSaveSettings()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\MiniWord"))
                {
                    key.SetValue("AutoSaveEnabled", _autoSaveEnabled ? 1 : 0);
                    key.SetValue("AutoSaveMinutes", _autoSaveMinutes);
                }
            }
            catch { }
        }

        private void ConfigureAutoSave()
        {
            using (AutoSaveDialog dialog = new AutoSaveDialog(_autoSaveEnabled, _autoSaveMinutes))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _autoSaveEnabled = dialog.AutoSaveEnabled;
                _autoSaveMinutes = dialog.IntervalMinutes;
                _autoSaveButton.Checked = _autoSaveEnabled;
                if (_autoSaveTimer != null)
                {
                    _autoSaveTimer.Stop();
                    _autoSaveTimer.Interval = _autoSaveMinutes * 60000;
                    _autoSaveTimer.Start();
                }
                SaveAutoSaveSettings();
                _stateLabel.Text = _autoSaveEnabled
                    ? "Autoguardado cada " + _autoSaveMinutes + " min"
                    : "Autoguardado desactivado; recuperación activa";
            }
        }

        private static string RecoveryFolder
        {
            get
            {
                string overridden = Environment.GetEnvironmentVariable("MINIWORD_RECOVERY_DIR");
                if (!string.IsNullOrEmpty(overridden))
                {
                    try { return Path.GetFullPath(overridden); }
                    catch { }
                }
                return Path.Combine(Path.GetTempPath(), "MiniWord", "Recovery");
            }
        }

        private void CloseRecoveryLock()
        {
            if (_recoveryLock != null)
            {
                try { _recoveryLock.Dispose(); } catch { }
                _recoveryLock = null;
            }
        }

        private bool AcquireRecoveryLock(string path)
        {
            CloseRecoveryLock();
            try
            {
                _recoveryLock = new FileStream(path + ".lock", FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                byte[] pid = System.Text.Encoding.UTF8.GetBytes(System.Diagnostics.Process.GetCurrentProcess().Id.ToString());
                _recoveryLock.Write(pid, 0, pid.Length);
                _recoveryLock.Flush();
                return true;
            }
            catch
            {
                _recoveryLock = null;
                return false;
            }
        }

        private static void PurgeSkippedRecovery()
        {
            try
            {
                string skipped = Path.Combine(RecoveryFolder, "Skipped");
                if (!Directory.Exists(skipped)) return;
                string[] files = Directory.GetFiles(skipped, "*.skipped*");
                for (int i = 0; i < files.Length; i++)
                {
                    try
                    {
                        if (File.GetLastWriteTimeUtc(files[i]) < DateTime.UtcNow.AddDays(-7))
                            File.Delete(files[i]);
                    }
                    catch { }
                }
            }
            catch { }
        }
        private string EnsureRecoveryPath()
        {
            if (!string.IsNullOrEmpty(_recoveryPath)) return _recoveryPath;
            Directory.CreateDirectory(RecoveryFolder);
            _recoveryPath = Path.Combine(RecoveryFolder,
                "recovery-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-"
                + Guid.NewGuid().ToString("N").Substring(0, 8) + ".docx");
            if (!AcquireRecoveryLock(_recoveryPath))
                throw new IOException("No se pudo bloquear la copia de recuperaciÃ³n.");
            return _recoveryPath;
        }

        private void WriteRecoveryMetadata(string documentPath)
        {
            string metadata = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(documentPath ?? string.Empty))
                + Environment.NewLine
                + DateTime.UtcNow.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + Environment.NewLine;
            AtomicFile.Save(documentPath + ".meta", delegate(string temporary)
            {
                File.WriteAllText(temporary, metadata, new System.Text.UTF8Encoding(false));
            });
        }

        private void ClearRecoveryFiles()
        {
            if (string.IsNullOrEmpty(_recoveryPath))
            {
                return;
            }
            string document = _recoveryPath;
            CloseRecoveryLock();
            try
            {
                try { File.Delete(document); } catch { }
                try { File.Delete(document + ".lock"); } catch { }
                try { File.Delete(document + ".bak"); } catch { }
                try { File.Delete(document + ".meta"); } catch { }
                try { File.Delete(document + ".meta.bak"); } catch { }
            }
            finally
            {
                _recoveryPath = null;
            }
        }

        private void QuarantineRecoveryFile(string document)
        {
            if (string.IsNullOrEmpty(document)) return;
            try
            {
                string folder = Path.Combine(RecoveryFolder, "Skipped");
                Directory.CreateDirectory(folder);
                string name = Path.GetFileName(document);
                string dest = Path.Combine(folder, name + ".skipped");
                if (File.Exists(document)) File.Move(document, dest);
                string meta = document + ".meta";
                if (File.Exists(meta)) File.Move(meta, dest + ".meta");
                string backup = document + ".bak";
                if (File.Exists(backup)) File.Move(backup, dest + ".bak");
                string lockPath = document + ".lock";
                if (File.Exists(lockPath)) { if (string.Equals(_recoveryPath, document, StringComparison.OrdinalIgnoreCase)) CloseRecoveryLock(); File.Delete(lockPath); }
            }
            catch
            {
                // Si no se puede mover, se deja la copia intacta para no perderla.
            }
        }
        private bool TryRestoreRecovery()
        {
            if (!_ownsRecoveryMutex) return false;
            try
            {
                string folder = RecoveryFolder;
                PurgeSkippedRecovery();
                if (!Directory.Exists(folder)) return false;
                string[] files = Directory.GetFiles(folder, "recovery-*.docx");
                Array.Sort(files, delegate(string a, string b)
                {
                    return File.GetLastWriteTimeUtc(b).CompareTo(File.GetLastWriteTimeUtc(a));
                });
                for (int i = 0; i < files.Length; i++)
                {
                    string document = files[i];
                    if (!AcquireRecoveryLock(document)) continue;
                    string meta = document + ".meta";
                    if (!File.Exists(meta)) { CloseRecoveryLock(); continue; }
                    string originalPath = string.Empty;
                    try
                    {
                        string[] lines = File.ReadAllLines(meta, System.Text.Encoding.UTF8);
                        if (lines.Length > 0 && !string.IsNullOrEmpty(lines[0]))
                            originalPath = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(lines[0]));
                    }
                    catch
                    {
                    }
                    if (!string.IsNullOrEmpty(originalPath))
                    {
                        try { originalPath = Path.GetFullPath(originalPath); }
                        catch { originalPath = string.Empty; }
                    }
                    string display = string.IsNullOrEmpty(originalPath) ? Path.GetFileName(document) : originalPath;
                    DialogResult answer = MessageBox.Show(this,
                        "Mini Word encontró un documento sin guardar.\n\n"
                        + display + "\n\n¿Desea recuperarlo?",
                        "Recuperación automática", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer != DialogResult.Yes)
                    {
                        CloseRecoveryLock();
                        QuarantineRecoveryFile(document);
                        return false;
                    }
                    DocumentModel model = DocumentFormats.ReadModel(document, DocumentFormat.Docx);
                    _editor.Document = model;
                    SyncTrackChangesButton();
                    _path = originalPath;
                    _format = DocumentFormat.Docx;
                    _dirty = true;
                    _recoveryPath = document;
                    UpdateTitle();
                    UpdateStatus();
                    _stateLabel.Text = "Documento recuperado";
                    return true;
                }
            }
            catch (Exception ex)
            {
                CloseRecoveryLock();
                CrashLog.Write("Recuperar documento", ex);
                MessageBox.Show(this, "No se pudo recuperar el documento.\n\n" + ex.Message,
                    "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            CloseRecoveryLock();
            return false;
        }

        private void AutoSaveNow()
        {
            if (!_dirty) return;
            try
            {
                // La recuperación se escribe siempre, aunque el autoguardado
                // programado esté desactivado. El original nunca se sobrescribe.
                string recovery = EnsureRecoveryPath();
                DocumentFormats.SaveDocx(recovery, _editor.Document);
                WriteRecoveryMetadata(recovery);
                _stateLabel.Text = "Copia de recuperación: " + recovery;
            }
            catch (Exception ex)
            {
                CrashLog.Write("Autoguardado", ex);
                _stateLabel.Text = "No se pudo autoguardar";
            }
        }

        private void UpdateTitle()
        {
            string name = string.IsNullOrEmpty(_path) ? "Sin título" : Path.GetFileName(_path);
            Text = (string.IsNullOrEmpty(name) ? "Mini Word " + AppInfo.Version : name + " - Mini Word " + AppInfo.Version) + (_dirty ? " *" : "");
        }

        private void NewDocument()
        {
            _editor.CommitPendingCellEdit();
            if (!ConfirmDiscard()) return;
            ClearRecoveryFiles();
            _recoveryPath = null;
            DocumentModel doc = new DocumentModel();
            doc.PageWidthTwips = 11906;
            doc.PageHeightTwips = 16838;
            doc.MarginLeftTwips = 1417;
            doc.MarginRightTwips = 1417;
            doc.MarginTopTwips = 1417;
            doc.MarginBottomTwips = 1417;
            Paragraph first = new Paragraph();
            first.LineSpacing = 1;
            doc.Paragraphs.Add(first);
            _editor.Document = doc;
            SyncTrackChangesButton();
            _path = string.Empty;
            _format = DocumentFormat.Docx;
            _dirty = false;
            _lossWarningAccepted = false;
            UpdateTitle();
            UpdateStatus();
            FocusEditor();
        }

        private void OpenDialog()
        {
            _editor.CommitPendingCellEdit();
            if (!ConfirmDiscard()) return;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = DocumentFormats.Filter;
                dialog.Title = "Abrir documento";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                OpenFile(dialog.FileName);
            }
        }

        private void OpenFile(string path)
        {
            _editor.CommitPendingCellEdit();
            try
            {
                DocumentFormat format = DocumentFormats.Detect(path);
                DocumentModel model = DocumentFormats.ReadModel(path, format);
                if (format == DocumentFormat.Doc && !model.UnsupportedFeatures.Contains("Formato DOC 97-2003 (importación)"))
                    model.UnsupportedFeatures.Add("Formato DOC 97-2003 (importación)");
                _editor.Document = model;
                SyncTrackChangesButton();
                ClearRecoveryFiles();
                _recoveryPath = null;
                _path = path;
                bool importedDoc = format == DocumentFormat.Doc;
                _format = importedDoc ? DocumentFormat.Docx : format;
                _dirty = false;
                _lossWarningAccepted = false;
                UpdateTitle();
                UpdateStatus();
                if (importedDoc)
                    _stateLabel.Text = "DOC 97-2003 importado; al guardar se creará un DOCX";
                if (model.SecurityWarnings.Count > 0)
                {
                    MessageBox.Show(this,
                        "Advertencia de seguridad:\n\n• "
                        + string.Join("\n• ", model.SecurityWarnings.ToArray())
                        + "\n\nMini Word no ejecuta macros, scripts ni controles ActiveX.",
                        "Contenido potencialmente activo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                AddRecentFile(path);


            }
            catch (Exception ex)
            {
                CrashLog.Write("Abrir documento", ex);
                MessageBox.Show(this, "No se pudo abrir el documento.\n\n" + ex.Message,
                    "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool Save(bool saveAs)
        {
            _editor.CommitPendingCellEdit();
            if (!saveAs && !string.IsNullOrEmpty(_path) && _format != DocumentFormat.Doc &&
                !string.Equals(Path.GetExtension(_path), ".doc", StringComparison.OrdinalIgnoreCase))
                return SaveTo(_path, _format);
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = DocumentFormats.SaveFilter;
                dialog.FilterIndex = _format == DocumentFormat.Odt ? 2 : (_format == DocumentFormat.Rtf ? 3 : (_format == DocumentFormat.Text ? 4 : 1));
                dialog.Title = "Guardar documento";
                if (!string.IsNullOrEmpty(_path)) dialog.FileName = Path.GetFileNameWithoutExtension(_path);
                else dialog.FileName = "Documento";
                if (dialog.ShowDialog(this) != DialogResult.OK) return false;
                DocumentFormat format = DocumentFormats.FromExtension(dialog.FileName);
                if (format == DocumentFormat.Doc) format = DocumentFormat.Docx;
                return SaveTo(dialog.FileName, format);
            }
        }

        private bool SaveTo(string path, DocumentFormat format)
        {
            try
            {
                if (_editor.Document.UnsupportedFeatures.Count > 0 && !_lossWarningAccepted)
                {
                    string features = string.Join("\n• ", _editor.Document.UnsupportedFeatures.ToArray());
                    DialogResult answer = MessageBox.Show(this,
                        "Este documento contiene elementos que Mini Word no conserva por completo:\n\n• "
                        + features + "\n\nSi guarda, esas partes pueden perderse. Se recomienda conservar el original.\n\n¿Desea guardar de todos modos?",
                        "Compatibilidad limitada", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (answer != DialogResult.Yes) return false;
                    _lossWarningAccepted = true;
                }
                switch (format)
                {
                    case DocumentFormat.Docx:
                        DocumentFormats.SaveDocx(path, _editor.Document);
                        break;
                    case DocumentFormat.Odt:
                        DocumentFormats.SaveOdt(path, _editor.Document);
                        break;
                    case DocumentFormat.Rtf:
                        DocumentFormats.SaveRtf(path, RtfWriter.Write(_editor.Document, false));
                        break;
                    case DocumentFormat.Text:
                        DocumentFormats.SaveText(path, _editor.Document.PlainText, TextEncodingKind.Unicode);
                        break;
                    default:
                        DocumentFormats.SaveDocx(path, _editor.Document);
                        format = DocumentFormat.Docx;
                        break;
                }
                _path = path;
                _format = format;
                _dirty = false;
                ClearRecoveryFiles();
                UpdateTitle();
                UpdateStatus();
                AddRecentFile(path);
                return true;
            }
            catch (Exception ex)
            {
                CrashLog.Write("Guardar documento", ex);
                MessageBox.Show(this, "No se pudo guardar el documento.\n\n" + ex.Message,
                    "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool ConfirmDiscard()
        {
            if (!_dirty) return true;
            DialogResult result = MessageBox.Show(this, "¿Desea guardar los cambios en el documento?",
                "Mini Word", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Cancel) return false;
            if (result == DialogResult.Yes) return Save(false);
            return true;
        }

        private void NativeFormClosing(object sender, FormClosingEventArgs e)
        {
            _editor.CommitPendingCellEdit();
            if (!ConfirmDiscard())
            {
                e.Cancel = true;
                return;
            }
            if (_autoSaveTimer != null) _autoSaveTimer.Stop();
            SaveAutoSaveSettings();
            ClearRecoveryFiles();
            if (_ownsRecoveryMutex)
            {
                try { _recoveryMutex.ReleaseMutex(); } catch { }
                _ownsRecoveryMutex = false;
            }
        }

        private void HeaderFooter()
        {
            DocumentSection section = _editor.CurrentSection;
            using (HeaderFooterDialog dialog = new HeaderFooterDialog(section.HeaderText,
                section.FooterText, section.FirstPageHeaderText, section.FirstPageFooterText,
                section.EvenPageHeaderText, section.EvenPageFooterText,
                section.HeaderDistanceTwips, section.FooterDistanceTwips,
                new RunFormat[] { section.HeaderFormat, section.FirstPageHeaderFormat, section.EvenPageHeaderFormat },
                new RunFormat[] { section.FooterFormat, section.FirstPageFooterFormat, section.EvenPageFooterFormat },
                new ParagraphAlignment[] { section.HeaderAlignment, section.FirstPageHeaderAlignment, section.EvenPageHeaderAlignment },
                new ParagraphAlignment[] { section.FooterAlignment, section.FirstPageFooterAlignment, section.EvenPageFooterAlignment },
                new List<ImageInline>[] { section.HeaderImages, section.FirstPageHeaderImages, section.EvenPageHeaderImages },
                new List<ImageInline>[] { section.FooterImages, section.FirstPageFooterImages, section.EvenPageFooterImages },
                section.TitlePage, section.DifferentOddEven))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _editor.SetCurrentHeaderFooter(dialog.HeaderText, dialog.FooterText,
                    dialog.FirstPageHeaderText, dialog.FirstPageFooterText,
                    dialog.EvenPageHeaderText, dialog.EvenPageFooterText,
                    dialog.DifferentFirstPage, dialog.DifferentOddEven,
                    dialog.HeaderDistanceTwips, dialog.FooterDistanceTwips,
                    new RunFormat[] { dialog.GetHeaderFormat(0), dialog.GetHeaderFormat(1), dialog.GetHeaderFormat(2) },
                    new RunFormat[] { dialog.GetFooterFormat(0), dialog.GetFooterFormat(1), dialog.GetFooterFormat(2) },
                    new ParagraphAlignment[] { dialog.GetHeaderAlignment(0), dialog.GetHeaderAlignment(1), dialog.GetHeaderAlignment(2) },
                    new ParagraphAlignment[] { dialog.GetFooterAlignment(0), dialog.GetFooterAlignment(1), dialog.GetFooterAlignment(2) },
                    new List<ImageInline>[] { dialog.GetHeaderImages(0), dialog.GetHeaderImages(1), dialog.GetHeaderImages(2) },
                    new List<ImageInline>[] { dialog.GetFooterImages(0), dialog.GetFooterImages(1), dialog.GetFooterImages(2) });
                _dirty = true;
                UpdateTitle();
            }
        }

        private void InsertTable()
        {
            using (TableEditorDialog dialog = new TableEditorDialog(null))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Rows == null) return;
                _editor.InsertTable(dialog.Rows);
                _dirty = true;
                UpdateTitle();
                FocusEditor();
            }
        }

        private void EditSelectedTableCell()
        {
            if (_editor.HasTableCellSelection)
            {
                _editor.EditSelectedTableCell();
                return;
            }
            EditTableAtCaret();
        }

        private void InsertImage()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Imágenes|*.png;*.jpg;*.jpeg;*.bmp;*.gif|Todos los archivos|*.*";
                dialog.Title = "Insertar imagen";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    int maxWidthTwips = _editor.GetImageMaxWidthTwips();
                    byte[] data = File.ReadAllBytes(dialog.FileName);
                    int width, height;
                    using (Image image = Image.FromFile(dialog.FileName))
                    {
                        width = image.Width;
                        height = image.Height;
                    }
                    string extension = Path.GetExtension(dialog.FileName).TrimStart('.').ToLowerInvariant();
                    if (extension == "jpg") extension = "jpeg";
                    _editor.InsertImage(data, extension, width, height, maxWidthTwips);
                    _dirty = true;
                    UpdateTitle();
                    ActiveControl = _editor;
                    _editor.Select();
                    FocusEditor();
                    BeginInvoke(new Action(delegate { ActiveControl = _editor; _editor.Select(); FocusEditor(); }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No se pudo insertar la imagen.\n\n" + ex.Message,
                        "Imagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void InsertDate()
        {
            string text = DateTime.Now.ToString("d 'de' MMMM 'de' yyyy",
                System.Globalization.CultureInfo.CurrentCulture);
            _editor.InsertText(text);
            _dirty = true;
            UpdateTitle();
            FocusEditor();
        }

        private void EditTableAtCaret()
        {
            int paragraph = _editor.TableAtCaret;
            if (paragraph < 0)
            {
                MessageBox.Show(this, "Haga clic en una celda; después puede escribir directamente en ella.",
                    "Tabla", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            EditTable(paragraph);
        }

        private void EditTable(int paragraph)
        {
            string[][] rows = _editor.GetTableRows(paragraph);
            if (rows == null) return;
            using (TableEditorDialog dialog = new TableEditorDialog(rows))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                if (dialog.DeleteTable)
                {
                    _editor.DeleteTable(paragraph);
                }
                else if (dialog.Rows != null)
                {
                    _editor.ReplaceTable(paragraph, dialog.Rows);
                }
                _dirty = true;
                UpdateTitle();
                FocusEditor();
            }
        }

        private void PageSetup()
        {
            using (NativePageSetupDialog dialog = new NativePageSetupDialog(_editor.Document))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _editor.SetPageSize(dialog.PageWidthTwips, dialog.PageHeightTwips);
                _editor.SetMargins(dialog.MarginLeftTwips, dialog.MarginTopTwips,
                    dialog.MarginRightTwips, dialog.MarginBottomTwips);
                _ruler.Invalidate();
                _dirty = true;
                UpdateTitle();
            }
        }

        private void ChooseColor()
        {
            bool cellEditing = _editor.IsEditingTableCell;
            if (cellEditing) _editor.SetCellEditorCommitSuppressed(true);
            try
            {
                using (ColorDialog dialog = new ColorDialog())
                {
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        _editor.SetForeColor(dialog.Color);
                        _dirty = true;
                    }
                }
            }
            finally
            {
                if (cellEditing)
                {
                    _editor.SetCellEditorCommitSuppressed(false);
                    _editor.FocusCellEditor();
                }
            }
        }

        private void ShowGoTo()
        {
            using (GoToDialog dialog = new GoToDialog(_editor.PageCount))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    _editor.GoToPage(dialog.PageNumber);
            }
        }

        private void PasteSpecial()
        {
            if (!Clipboard.ContainsText())
            {
                MessageBox.Show(this, "El portapapeles no contiene texto.", "Pegado especial",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            DialogResult answer = MessageBox.Show(this,
                "¿Pegar como texto sin formato?\n\nSí = sin formato   ·   No = con formato",
                "Pegado especial", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (answer == DialogResult.Yes) _editor.PastePlainText();
            else if (answer == DialogResult.No) _editor.Paste();
            FocusEditor();
        }

        private void ShowComments()
        {
            using (CommentsDialog dialog = new CommentsDialog(_editor.Document))
            {
                dialog.ShowDialog(this);
                _dirty = true;
                UpdateTitle();
            }
        }

        private void ToggleTrackChanges()
        {
            bool enabled = !_editor.TrackChanges;
            _editor.SetTrackChanges(enabled);
            if (_trackChangesButton != null) _trackChangesButton.Checked = enabled;
            _status.Text = enabled ? "Control de cambios activado" : "Control de cambios desactivado";
            FocusEditor();
        }

        private void CopyFormatForPainter()
        {
            _formatPainterFormat = _editor.CurrentFormat;
            _status.Text = "Formato copiado; seleccione el texto de destino";
            FocusEditor();
        }
        private void FindReplace()
        {
            using (NativeFindDialog dialog = new NativeFindDialog(_editor))
            {
                dialog.ShowDialog(this);
            }
        }

        private void SetZoom(float zoom)
        {
            _editor.Zoom = zoom;
            if (_ruler != null) _ruler.Invalidate();
            _stateLabel.Text = "Zoom " + Math.Round(zoom * 100) + "%";
        }

        private void PrintDocument()
        {
            _editor.CommitPendingCellEdit();
            float zoom = _editor.Zoom;
            try
            {
                _editor.Zoom = 1f;
                using (PrintDialog dialog = new PrintDialog())
                {
                    PrintDocument document = CreatePrintDocument();
                    dialog.Document = document;
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    try { document.Print(); }
                    catch (Exception ex) { MessageBox.Show(this, ex.Message, "Imprimir", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
            finally { _editor.Zoom = zoom; }
        }

        private void PrintPreview()
        {
            _editor.CommitPendingCellEdit();
            float zoom = _editor.Zoom;
            try
            {
                _editor.Zoom = 1f;
                using (PrintPreviewDialog dialog = new PrintPreviewDialog())
                {
                    dialog.Document = CreatePrintDocument();
                    dialog.Width = 1000;
                    dialog.Height = 760;
                    dialog.Shown += delegate { AttachPreviewWheel(dialog); };
                    dialog.ShowDialog(this);
                }
            }
            finally { _editor.Zoom = zoom; }
        }

        private void AttachPreviewWheel(PrintPreviewDialog dialog)
        {
            PrintPreviewControl preview = FindPreviewControl(dialog);
            if (preview == null || preview.Tag != null) return;
            preview.Tag = "wheel";
            preview.MouseWheel += delegate(object sender, MouseEventArgs e)
            {
                int step = e.Delta > 0 ? -1 : 1;
                int max = Math.Max(0, _editor.PageCount - 1);
                int target = preview.StartPage + step;
                if (target < 0) target = 0;
                if (target > max) target = max;
                preview.StartPage = target;
                HandledMouseEventArgs handled = e as HandledMouseEventArgs;
                if (handled != null) handled.Handled = true;
            };
        }

        private static PrintPreviewControl FindPreviewControl(Control parent)
        {
            if (parent is PrintPreviewControl) return (PrintPreviewControl)parent;
            for (int i = 0; i < parent.Controls.Count; i++)
            {
                PrintPreviewControl found = FindPreviewControl(parent.Controls[i]);
                if (found != null) return found;
            }
            return null;
        }

        private PrintDocument CreatePrintDocument()
        {
            PrintDocument document = new PrintDocument();
            DocumentModel model = _editor.Document;
            document.DefaultPageSettings.PaperSize = new PaperSize("Documento",
                Math.Max(100, model.PageWidthTwips * 100 / 1440),
                Math.Max(100, model.PageHeightTwips * 100 / 1440));
            document.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            document.DefaultPageSettings.Landscape = false;
            document.PrintPage += PrintPage;
            document.BeginPrint += delegate { _printPage = 0; };
            return document;
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_printPage >= _editor.PageCount) { e.HasMorePages = false; return; }
            _editor.DrawPageForPrint(e.Graphics, _printPage, e.PageBounds);
            _printPage++;
            e.HasMorePages = _printPage < _editor.PageCount;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.W)) { Close(); return true; }
            if (keyData == (Keys.Control | Keys.H)) { FindReplace(); return true; }
            if (keyData == (Keys.Control | Keys.K)) { EditHyperlink(); return true; }
            if (keyData == Keys.F7) { ShowSpellCheck(); return true; }
            if (keyData == (Keys.Control | Keys.E)) { _editor.SetAlignment(ParagraphAlignment.Center); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.L)) { _editor.SetAlignment(ParagraphAlignment.Left); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.R)) { _editor.SetAlignment(ParagraphAlignment.Right); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.J)) { _editor.SetAlignment(ParagraphAlignment.Justify); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.M)) { _editor.IncreaseIndent(); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.M)) { _editor.DecreaseIndent(); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.L)) { _editor.ToggleBullet(); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.D7) || keyData == (Keys.Control | Keys.Shift | Keys.NumPad7))
            { _editor.ToggleNumberedList(); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.N)) { _editor.ApplyParagraphStyle("Normal"); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.D1) || keyData == (Keys.Control | Keys.Alt | Keys.NumPad1))
            { _editor.ApplyParagraphStyle("Título 1"); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.D2) || keyData == (Keys.Control | Keys.Alt | Keys.NumPad2))
            { _editor.ApplyParagraphStyle("Título 2"); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.D3) || keyData == (Keys.Control | Keys.Alt | Keys.NumPad3))
            { _editor.ApplyParagraphStyle("Título 3"); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.E)) { ToggleTrackChanges(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.M)) { ShowComments(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.C)) { CopyFormatForPainter(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.V)) { PasteSpecial(); return true; }
            if (keyData == Keys.F9) { _editor.UpdateFields(); return true; }
            if (keyData == Keys.F11) { _editor.SelectNextField((Control.ModifierKeys & Keys.Shift) == 0); return true; }
            if (keyData == Keys.F5 || keyData == (Keys.Control | Keys.G)) { ShowGoTo(); return true; }
            if (keyData == (Keys.Control | Keys.D1) || keyData == (Keys.Control | Keys.NumPad1)) { _editor.SetCurrentLineSpacing(0); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.D5) || keyData == (Keys.Control | Keys.NumPad5)) { _editor.SetCurrentLineSpacing(1); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.D2) || keyData == (Keys.Control | Keys.NumPad2)) { _editor.SetCurrentLineSpacing(2); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.OemPeriod)) { ChangeFontSize(2f); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.Oemcomma)) { ChangeFontSize(-2f); return true; }
            if (keyData == (Keys.Control | Keys.OemCloseBrackets)) { ChangeFontSize(1f); return true; }
            if (keyData == (Keys.Control | Keys.OemOpenBrackets)) { ChangeFontSize(-1f); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.Space)) { _editor.InsertText("\u00A0"); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.Z)) { _editor.Redo(); return true; }            if (keyData == (Keys.Control | Keys.N)) { NewDocument(); return true; }
            if (keyData == (Keys.Control | Keys.O)) { OpenDialog(); return true; }
            if (keyData == (Keys.Control | Keys.S)) { Save(false); return true; }
            if (keyData == (Keys.Control | Keys.Shift | Keys.S) || keyData == Keys.F12) { Save(true); return true; }
            if (keyData == (Keys.Control | Keys.P)) { PrintDocument(); return true; }
            if (keyData == (Keys.Control | Keys.F)) { FindReplace(); return true; }
            if (keyData == (Keys.Control | Keys.Enter)) { _editor.InsertPageBreak(); return true; }
            if (keyData == (Keys.Control | Keys.Add) || keyData == (Keys.Control | Keys.Oemplus) ||
                keyData == (Keys.Control | Keys.Subtract) || keyData == (Keys.Control | Keys.OemMinus))
            {
                SetZoom(_editor.Zoom + (keyData == (Keys.Control | Keys.Add) || keyData == (Keys.Control | Keys.Oemplus) ? 0.1f : -0.1f));
                return true;
            }
            if (keyData == (Keys.Control | Keys.D0)) { _editor.ApplyParagraphStyle("Normal"); SyncToolbar(); FocusEditor(); return true; }
            if (keyData == (Keys.Control | Keys.Alt | Keys.D0)) { SetZoom(1f); return true; }
            if (keyData == Keys.Enter && _editor != null && !_editor.ContainsFocus && ActiveControl is ToolStrip)
            {
                ActiveControl = _editor;
                _editor.Select();
                FocusEditor();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}


