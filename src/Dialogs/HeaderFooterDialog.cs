using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using WordPadClone.Formats;

namespace WordPadClone.Dialogs
{
    internal sealed class HeaderFooterDialog : Form
    {
        private readonly TextBox _header;
        private readonly TextBox _footer;
        private readonly TextBox _firstHeader;
        private readonly TextBox _firstFooter;
        private readonly TextBox _evenHeader;
        private readonly TextBox _evenFooter;
        private readonly CheckBox _differentFirst;
        private readonly CheckBox _differentEven;
        private readonly NumericUpDown _headerDistance;
        private readonly NumericUpDown _footerDistance;
        private readonly NumericUpDown _fontSize;
        private readonly Button _fontColor;
        private readonly ComboBox _alignment;
        private readonly ComboBox _variant;
        private readonly List<ImageInline>[] _headerImages = new List<ImageInline>[3];
        private readonly List<ImageInline>[] _footerImages = new List<ImageInline>[3];
        private readonly RunFormat[] _headerFormats = new RunFormat[3];
        private readonly RunFormat[] _footerFormats = new RunFormat[3];
        private readonly ParagraphAlignment[] _headerAlignments = new ParagraphAlignment[3];
        private readonly ParagraphAlignment[] _footerAlignments = new ParagraphAlignment[3];
        private int _activeVariant;
        private bool _loadingFormat;
        private Color _color;

        public string HeaderText { get { return _header.Text; } }
        public string FooterText { get { return _footer.Text; } }
        public string FirstPageHeaderText { get { return _firstHeader.Text; } }
        public string FirstPageFooterText { get { return _firstFooter.Text; } }
        public string EvenPageHeaderText { get { return _evenHeader.Text; } }
        public string EvenPageFooterText { get { return _evenFooter.Text; } }
        public bool DifferentFirstPage { get { return _differentFirst.Checked; } }
        public bool DifferentOddEven { get { return _differentEven.Checked; } }
        public int HeaderDistanceTwips { get { return (int)Math.Round((double)_headerDistance.Value / 2.54 * 1440.0); } }
        public int FooterDistanceTwips { get { return (int)Math.Round((double)_footerDistance.Value / 2.54 * 1440.0); } }

        public RunFormat HeaderFormat { get { SaveActiveVariant(); return _headerFormats[0]; } }
        public RunFormat FooterFormat { get { SaveActiveVariant(); return _footerFormats[0]; } }
        public ParagraphAlignment HeaderAlignment { get { SaveActiveVariant(); return _headerAlignments[0]; } }
        public ParagraphAlignment FooterAlignment { get { SaveActiveVariant(); return _footerAlignments[0]; } }
        public RunFormat GetHeaderFormat(int index) { SaveActiveVariant(); return _headerFormats[index]; }
        public RunFormat GetFooterFormat(int index) { SaveActiveVariant(); return _footerFormats[index]; }
        public ParagraphAlignment GetHeaderAlignment(int index) { SaveActiveVariant(); return _headerAlignments[index]; }
        public ParagraphAlignment GetFooterAlignment(int index) { SaveActiveVariant(); return _footerAlignments[index]; }
        public List<ImageInline> GetHeaderImages(int index) { return _headerImages[index]; }
        public List<ImageInline> GetFooterImages(int index) { return _footerImages[index]; }

        public HeaderFooterDialog(string header, string footer, int headerTwips, int footerTwips,
            RunFormat format, ParagraphAlignment alignment)
            : this(header, footer, string.Empty, string.Empty, string.Empty, string.Empty,
                headerTwips, footerTwips,
                new RunFormat[] { format, format, format },
                new RunFormat[] { format, format, format },
                new ParagraphAlignment[] { alignment, alignment, alignment },
                new ParagraphAlignment[] { alignment, alignment, alignment },
                null, null, false, false)
        {
        }

        public HeaderFooterDialog(string header, string footer, string firstHeader, string firstFooter,
            string evenHeader, string evenFooter, int headerTwips, int footerTwips,
            RunFormat format, ParagraphAlignment alignment, bool differentFirstPage, bool differentOddEven)
            : this(header, footer, firstHeader, firstFooter, evenHeader, evenFooter,
                headerTwips, footerTwips,
                new RunFormat[] { format, format, format },
                new RunFormat[] { format, format, format },
                new ParagraphAlignment[] { alignment, alignment, alignment },
                new ParagraphAlignment[] { alignment, alignment, alignment },
                null, null, differentFirstPage, differentOddEven)
        {
        }

        public HeaderFooterDialog(string header, string footer, string firstHeader, string firstFooter,
            string evenHeader, string evenFooter, int headerTwips, int footerTwips,
            RunFormat[] headerFormats, RunFormat[] footerFormats,
            ParagraphAlignment[] headerAlignments, ParagraphAlignment[] footerAlignments,
            bool differentFirstPage, bool differentOddEven)
            : this(header, footer, firstHeader, firstFooter, evenHeader, evenFooter,
                headerTwips, footerTwips, headerFormats, footerFormats,
                headerAlignments, footerAlignments, null, null, differentFirstPage, differentOddEven)
        {
        }

        public HeaderFooterDialog(string header, string footer, string firstHeader, string firstFooter,
            string evenHeader, string evenFooter, int headerTwips, int footerTwips,
            RunFormat[] headerFormats, RunFormat[] footerFormats,
            ParagraphAlignment[] headerAlignments, ParagraphAlignment[] footerAlignments,
            List<ImageInline>[] headerImages, List<ImageInline>[] footerImages,
            bool differentFirstPage, bool differentOddEven)
        {
            for (int i = 0; i < 3; i++)
            {
                _headerFormats[i] = headerFormats != null && headerFormats.Length > i && headerFormats[i] != null
                    ? headerFormats[i].Clone() : new RunFormat("Calibri", 18);
                _footerFormats[i] = footerFormats != null && footerFormats.Length > i && footerFormats[i] != null
                    ? footerFormats[i].Clone() : new RunFormat("Calibri", 18);
                _headerAlignments[i] = headerAlignments != null && headerAlignments.Length > i
                    ? headerAlignments[i] : ParagraphAlignment.Left;
                _footerAlignments[i] = footerAlignments != null && footerAlignments.Length > i
                    ? footerAlignments[i] : ParagraphAlignment.Left;
                _headerImages[i] = CloneImages(headerImages, i);
                _footerImages[i] = CloneImages(footerImages, i);
            }
            RunFormat format = _headerFormats[0];
            ParagraphAlignment alignment = _headerAlignments[0];
            _color = format != null && format.ForeColor.HasValue ? format.ForeColor.Value : Color.Black;
            Text = "Encabezado y pie de página";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(600, 650);

            Label intro = new Label();
            intro.Text = "Use {PAGE} y {NUMPAGES} para insertar el número de página y el total. Puede tener encabezado distinto en la primera página y en las páginas pares.";
            intro.Location = new Point(14, 10);
            intro.Size = new Size(570, 36);
            Controls.Add(intro);

            TabControl tabs = new TabControl();
            tabs.Location = new Point(14, 48);
            tabs.Size = new Size(566, 300);
            Controls.Add(tabs);

            TabPage generalPage = new TabPage("Principal");
            TabPage firstPage = new TabPage("Primera página");
            TabPage evenPage = new TabPage("Páginas pares");
            tabs.TabPages.Add(generalPage);
            tabs.TabPages.Add(firstPage);
            tabs.TabPages.Add(evenPage);

            _header = AddTextGroup(generalPage, "Encabezado", header, 4, _headerImages[0]);
            _footer = AddTextGroup(generalPage, "Pie de página", footer, 136, _footerImages[0]);
            _firstHeader = AddTextGroup(firstPage, "Encabezado de la primera página", firstHeader, 4, _headerImages[1]);
            _firstFooter = AddTextGroup(firstPage, "Pie de la primera página", firstFooter, 136, _footerImages[1]);
            _evenHeader = AddTextGroup(evenPage, "Encabezado de páginas pares", evenHeader, 4, _headerImages[2]);
            _evenFooter = AddTextGroup(evenPage, "Pie de páginas pares", evenFooter, 136, _footerImages[2]);

            _differentFirst = new CheckBox();
            _differentFirst.Text = "Primera página diferente";
            _differentFirst.Location = new Point(14, 354);
            _differentFirst.AutoSize = true;
            _differentFirst.Checked = differentFirstPage;
            Controls.Add(_differentFirst);

            _differentEven = new CheckBox();
            _differentEven.Text = "Páginas pares diferentes";
            _differentEven.Location = new Point(220, 354);
            _differentEven.AutoSize = true;
            _differentEven.Checked = differentOddEven;
            Controls.Add(_differentEven);
            GroupBox formatGroup = new GroupBox();
            formatGroup.Text = "Formato de encabezado y pie";
            formatGroup.Location = new Point(14, 382);
            formatGroup.Size = new Size(566, 112);
            Controls.Add(formatGroup);

            Label variantLabel = new Label();
            variantLabel.Text = "Variante:";
            variantLabel.Location = new Point(10, 58);
            variantLabel.AutoSize = true;
            formatGroup.Controls.Add(variantLabel);

            _variant = new ComboBox();
            _variant.AccessibleName = "Variante de encabezado";
            _variant.DropDownStyle = ComboBoxStyle.DropDownList;
            _variant.Location = new Point(72, 54);
            _variant.Width = 170;
            _variant.Items.AddRange(new object[] { "Principal", "Primera página", "Páginas pares" });
            _variant.SelectedIndex = 0;
            _variant.SelectedIndexChanged += delegate
            {
                if (_loadingFormat) return;
                SaveActiveVariant();
                _activeVariant = Math.Max(0, _variant.SelectedIndex);
                LoadVariant();
            };
            formatGroup.Controls.Add(_variant);

            Label sizeLabel = new Label();
            sizeLabel.Text = "Tamaño:";
            sizeLabel.Location = new Point(10, 27);
            sizeLabel.AutoSize = true;
            formatGroup.Controls.Add(sizeLabel);

            _fontSize = new NumericUpDown();
            _fontSize.Location = new Point(70, 23);
            _fontSize.Width = 62;
            _fontSize.Minimum = 6m;
            _fontSize.Maximum = 48m;
            _fontSize.DecimalPlaces = 1;
            _fontSize.Increment = 0.5m;
            _fontSize.Value = Math.Max(6m, Math.Min(48m,
                (format == null ? 18 : format.SizeHalfPoints) / 2m));
            formatGroup.Controls.Add(_fontSize);

            Label colorLabel = new Label();
            colorLabel.Text = "Color:";
            colorLabel.Location = new Point(158, 27);
            colorLabel.AutoSize = true;
            formatGroup.Controls.Add(colorLabel);

            _fontColor = new Button();
            _fontColor.Text = "Color";
            _fontColor.Location = new Point(205, 23);
            _fontColor.Size = new Size(58, 26);
            _fontColor.BackColor = _color;
            _fontColor.ForeColor = Contrast(_color);
            _fontColor.FlatStyle = FlatStyle.Flat;
            _fontColor.Click += delegate
            {
                using (ColorDialog dialog = new ColorDialog())
                {
                    dialog.FullOpen = true;
                    dialog.Color = _color;
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    _color = dialog.Color;
                    _fontColor.BackColor = dialog.Color;
                    _fontColor.ForeColor = Contrast(dialog.Color);
                }
            };
            formatGroup.Controls.Add(_fontColor);

            Label alignLabel = new Label();
            alignLabel.Text = "Alineación:";
            alignLabel.Location = new Point(292, 27);
            alignLabel.AutoSize = true;
            formatGroup.Controls.Add(alignLabel);

            _alignment = new ComboBox();
            _alignment.AccessibleName = "Alineación del encabezado y pie";
            _alignment.DropDownStyle = ComboBoxStyle.DropDownList;
            _alignment.Location = new Point(375, 23);
            _alignment.Width = 150;
            _alignment.Items.AddRange(new object[] { "Izquierda", "Centro", "Derecha" });
            _alignment.SelectedIndex = AlignmentIndex(alignment);
            formatGroup.Controls.Add(_alignment);

            Label hd = new Label();
            hd.Text = "Distancia del encabezado:";
            hd.Location = new Point(14, 510);
            hd.AutoSize = true;
            Controls.Add(hd);

            _headerDistance = MakeDistance(headerTwips, 166, 506);
            Controls.Add(_headerDistance);
            Label hdUnit = new Label();
            hdUnit.Text = "cm";
            hdUnit.Location = new Point(265, 510);
            hdUnit.AutoSize = true;
            Controls.Add(hdUnit);

            Label fd = new Label();
            fd.Text = "Distancia del pie:";
            fd.Location = new Point(292, 510);
            fd.AutoSize = true;
            Controls.Add(fd);

            _footerDistance = MakeDistance(footerTwips, 404, 506);
            Controls.Add(_footerDistance);
            Label fdUnit = new Label();
            fdUnit.Text = "cm";
            fdUnit.Location = new Point(505, 510);
            fdUnit.AutoSize = true;
            Controls.Add(fdUnit);

            CheckBox different = new CheckBox();
            different.Text = "Mantener encabezado y pie en la vista de documento y en la impresión";
            different.Checked = true;
            different.Enabled = false;
            different.Location = new Point(14, 550);
            different.AutoSize = true;
            Controls.Add(different);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(404, 585);
            ok.Size = new Size(75, 26);
            ok.UseVisualStyleBackColor = true;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(485, 585);
            cancel.Size = new Size(81, 26);
            cancel.UseVisualStyleBackColor = true;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private static TextBox AddTextGroup(Control parent, string title, string text, int y,
            List<ImageInline> images)
        {
            GroupBox group = new GroupBox();
            group.Text = title;
            group.AccessibleName = title;
            group.Location = new Point(8, y);
            group.Size = new Size(548, 126);
            parent.Controls.Add(group);
            TextBox box = new TextBox();
            box.AccessibleName = title;
            box.AccessibleDescription = "Texto de " + title;
            box.Multiline = true;
            box.ScrollBars = ScrollBars.Vertical;
            box.AcceptsReturn = true;
            box.Location = new Point(10, 22);
            box.Size = new Size(400, 94);
            box.Text = text ?? string.Empty;
            group.Controls.Add(box);

            Label count = new Label();
            count.Location = new Point(418, 84);
            count.Size = new Size(120, 20);
            count.AutoEllipsis = true;
            group.Controls.Add(count);

            Button add = new Button();
            add.Text = "Añadir imagen...";
            add.AccessibleName = "Añadir imagen a " + title;
            add.Location = new Point(418, 20);
            add.Size = new Size(120, 26);
            add.Click += delegate
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
                    dialog.Filter = "Imágenes|*.png;*.jpg;*.jpeg;*.bmp;*.gif|Todos los archivos|*.*";
                    dialog.Multiselect = true;
                    if (dialog.ShowDialog(parent.FindForm()) != DialogResult.OK) return;
                    for (int i = 0; i < dialog.FileNames.Length; i++)
                    {
                        try
                        {
                            FileInfo info = new FileInfo(dialog.FileNames[i]);
                            if (info.Length > 20 * 1024 * 1024)
                                throw new InvalidOperationException("Imagen mayor de 20 MB.");
                            byte[] data = File.ReadAllBytes(dialog.FileNames[i]);
                            int width, height;
                            using (Image image = Image.FromFile(dialog.FileNames[i]))
                            {
                                width = image.Width;
                                height = image.Height;
                            }
                            string extension = Path.GetExtension(dialog.FileNames[i]).TrimStart('.').ToLowerInvariant();
                            if (extension == "jpg") extension = "jpeg";
                            images.Add(new ImageInline(data, extension, width * 15, height * 15,
                                new RunFormat("Calibri", 18)));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("No se pudo añadir la imagen.\r\n\r\n" + ex.Message,
                                "Mini Word", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    count.Text = images.Count + (images.Count == 1 ? " imagen" : " imágenes");
                }
            };
            group.Controls.Add(add);

            Button clear = new Button();
            clear.Text = "Quitar todas";
            clear.Location = new Point(418, 50);
            clear.Size = new Size(120, 26);
            clear.Click += delegate
            {
                images.Clear();
                count.Text = "0 imágenes";
            };
            group.Controls.Add(clear);

            count.Text = images.Count + (images.Count == 1 ? " imagen" : " imágenes");
            return box;
        }

        private static List<ImageInline> CloneImages(List<ImageInline>[] sets, int index)
        {
            List<ImageInline> result = new List<ImageInline>();
            if (sets == null || sets.Length <= index || sets[index] == null) return result;
            List<ImageInline> source = sets[index];
            for (int i = 0; i < source.Count; i++)
            {
                ImageInline image = source[i];
                result.Add(new ImageInline(image.Data, image.Extension, image.WidthTwips,
                    image.HeightTwips, image.Format.Clone()));
            }
            return result;
        }

        private void SaveActiveVariant()
        {
            if (_fontSize == null || _alignment == null) return;
            RunFormat format = BuildFormat();
            _headerFormats[_activeVariant] = format.Clone();
            _footerFormats[_activeVariant] = format.Clone();
            _headerAlignments[_activeVariant] = SelectedAlignment();
            _footerAlignments[_activeVariant] = SelectedAlignment();
        }

        private void LoadVariant()
        {
            RunFormat format = _headerFormats[_activeVariant] ?? new RunFormat("Calibri", 18);
            _loadingFormat = true;
            try
            {
                _color = format.ForeColor ?? Color.Black;
                _fontSize.Value = Math.Max(_fontSize.Minimum, Math.Min(_fontSize.Maximum, format.SizeHalfPoints / 2m));
                _fontColor.BackColor = _color;
                _fontColor.ForeColor = Contrast(_color);
                _alignment.SelectedIndex = AlignmentIndex(_headerAlignments[_activeVariant]);
            }
            finally
            {
                _loadingFormat = false;
            }
        }

        private RunFormat BuildFormat()
        {
            RunFormat result = new RunFormat("Calibri", Math.Max(2, (int)Math.Round(_fontSize.Value * 2m)));
            result.ForeColor = _color;
            return result;
        }

        private ParagraphAlignment SelectedAlignment()
        {
            switch (_alignment.SelectedIndex)
            {
                case 1: return ParagraphAlignment.Center;
                case 2: return ParagraphAlignment.Right;
                case 3: return ParagraphAlignment.Justify;
                default: return ParagraphAlignment.Left;
            }
        }

        private static int AlignmentIndex(ParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case ParagraphAlignment.Center: return 1;
                case ParagraphAlignment.Right: return 2;
                case ParagraphAlignment.Justify: return 3;
                default: return 0;
            }
        }

        private static Color Contrast(Color color)
        {
            int brightness = (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
            return brightness > 140 ? Color.Black : Color.White;
        }

        private static NumericUpDown MakeDistance(int twips, int x, int y)
        {
            NumericUpDown box = new NumericUpDown();
            box.DecimalPlaces = 1;
            box.Increment = 0.1m;
            box.Minimum = 0.0m;
            box.Maximum = 10.0m;
            box.Value = Math.Max(box.Minimum, Math.Min(box.Maximum, (decimal)(twips / 1440.0 * 2.54)));
            box.Location = new Point(x, y);
            box.Size = new Size(90, 24);
            box.TextAlign = HorizontalAlignment.Left;
            return box;
        }
    }
}
