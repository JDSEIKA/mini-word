using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WordPadClone.Dialogs;
using WordPadClone.Formats;

namespace WordPadClone.Controls
{
	internal sealed class NativeDocumentEditor : ScrollableControl
	{
		[DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		private static extern IntPtr ShellExecute(IntPtr hwnd, string operation, string file,
			string parameters, string directory, int showCommand);

		private sealed class LayoutPage
		{
			public int Index;

			public float Left;

			public float Top;

			public float Width;

			public float Height;

			public float BodyLeft;

			public float BodyTop;

			public float BodyRight;

			public float BodyBottom;

			public float FullBodyLeft;

			public float FullBodyRight;

			public float ColumnWidth;

			public int ColumnIndex;

			public int ColumnCount = 1;

			public DocumentSection Section;

			public readonly List<LayoutLine> Lines = new List<LayoutLine>();

			public readonly List<LayoutTable> Tables = new List<LayoutTable>();

			public readonly List<LayoutFloating> Floatings = new List<LayoutFloating>();
		}

		private sealed class LayoutLine
		{
			public int Page;

			public int Paragraph;

			public int Start;

			public int End;

			public float X;

			public float Y;

			public float Width;

			public float Height;

			public bool First;

			public bool Bullet;

			public string Number;

			public float BulletX;

			public readonly List<LayoutSegment> Segments = new List<LayoutSegment>();
		}

		private sealed class LayoutSegment
		{
			public string Text;

			public RunFormat Format;

			public int Start;

			public int End;

			public float X;

			public float Y;

			public float Width;

			public float Height;

			public ImageInline Image;

			public DocumentNote Note;

			public DocumentComment Comment;

			public Font Font;

			public bool IsSpace;
		}

		private sealed class LayoutFloating
		{
			public FloatingObject Object;

			public float X;

			public float Y;

			public float Width;

			public float Height;

			public int ColumnIndex;
		}

		private sealed class LayoutTable
		{
			public int Page;

			public int Paragraph;

			public float X;

			public float Y;

			public float Width;

			public float Height;

			public float RowHeight;

			public readonly List<LayoutTableRow> Rows = new List<LayoutTableRow>();

			public readonly List<LayoutTableCell> Cells = new List<LayoutTableCell>();
		}

		private sealed class LayoutTableRow
		{
			public float Y;

			public float Height;
		}

		private sealed class LayoutTableCell
		{
			public TableCell Cell;

			public int Row;

			public int Column;

			public float X;

			public float Y;

			public float Width;

			public float Height;
		}

		private sealed class CaretPos
		{
			public int P;

			public int Offset;

			public CaretPos(int p, int offset)
			{
				P = p;
				Offset = offset;
			}
		}

		private sealed class LineBuilder
		{
			public readonly List<LayoutSegment> Segments = new List<LayoutSegment>();

			public float Width;

			public float Height;

			public float X;

			public float Y;

			public int Start = -1;

			public int End;

			public bool HasContent;

			public bool Bullet;

			public string Number;

			public float BulletX;
		}

		private sealed class Token
		{
			public string Text;

			public RunFormat Format;

			public int Start;

			public int End;

			public bool Tab;

			public bool Newline;

			public ImageInline Image;

			public DocumentNote Note;

			public DocumentComment Comment;

			public bool IsSpace;
		}

		private sealed class FontCache : IDisposable
		{
			private readonly Dictionary<string, Font> _fonts = new Dictionary<string, Font>(StringComparer.Ordinal);

			public Font Get(string family, float pixels, FontStyle style)
			{
				if (string.IsNullOrEmpty(family))
				{
					family = "Calibri";
				}
				if (pixels < 1f)
				{
					pixels = 1f;
				}
				string key = family.ToLowerInvariant() + "|" + pixels.ToString("0.##") + "|" + (int)style;
				Font value;
				if (_fonts.TryGetValue(key, out value))
				{
					return value;
				}
				try
				{
					value = new Font(family, pixels, style, GraphicsUnit.Pixel);
				}
				catch
				{
					value = new Font("Calibri", pixels, style, GraphicsUnit.Pixel);
				}
				_fonts[key] = value;
				return value;
			}

			public void Clear()
			{
				foreach (Font value in _fonts.Values)
				{
					value.Dispose();
				}
				_fonts.Clear();
			}

			public void Dispose()
			{
				Clear();
			}
		}

		private const float PageGap = 18f;

		private const float PxPerTwip = 1f / 15f;

		private static readonly StringFormat MeasureFormat = CreateMeasureFormat();

		private DocumentModel _document;

		private readonly List<LayoutPage> _pages = new List<LayoutPage>();

		private readonly List<LayoutLine> _lines = new List<LayoutLine>();

		private readonly List<LayoutTable> _tables = new List<LayoutTable>();

		private readonly FontCache _fonts = new FontCache();

		private readonly Timer _blinkTimer;

		private readonly Stack<DocumentModel> _undo = new Stack<DocumentModel>();

		private readonly Stack<DocumentModel> _redo = new Stack<DocumentModel>();

		private float _zoom = 1f;
		private bool _showRuler = true;

		private Graphics _layoutGraphics;
		private Graphics _measureGraphics;

		private CaretPos _caret;

		private CaretPos _anchor;

		private bool _selecting;

		private bool _caretOn;

		private bool _rebuilding;

		private bool _dirtyLayout = true;

		private RunFormat _typingFormat = new RunFormat("Arial", 22);

		private int _preferredX = -1;

		private int _currentPage;

		private int _lastViewedPage = -1;

		private SystemSpellChecker _spellChecker;

		private Timer _spellTimer;

		private readonly List<SpellError> _spellErrors = new List<SpellError>();

		private int _spellVersion;

		private bool _spellEnabled = true;

		private string _spellLanguage = "es-MX";

		private TextBox _cellEditor;

		private LayoutTableCell _cellEdit;

		private int _tableSelectionParagraph = -1;

		private int _tableSelectionStartRow;

		private int _tableSelectionStartColumn;

		private int _tableSelectionEndRow;

		private int _tableSelectionEndColumn;

		private bool _tableDragging;

		private bool _tableMoved;

		private Point _tableClickScroll;

		private int _cellEditParagraph = -1;

		private bool _suppressCellLeave;

		private bool _suppressCellCommit;

		private string _cellOriginalText = string.Empty;

		private ImageInline _selectedImage;

		private FloatingObject _selectedFloating;

		private int _spellingParagraph = -1;

		private int _spellingStart;

		private int _spellingLength;

		private sealed class FormattedRun
		{
			public string Text;
			public RunFormat Format;

			public FormattedRun(string text, RunFormat format)
			{
				Text = text;
				Format = format;
			}
		}

		private FloatingObject _clipboardFloating;

		private bool _clipboardIsShape;

		private bool _movingFloating;

		private bool _floatingUndoPushed;

		private bool _resizingFloating;

		private int _floatingHandle = -1;

		private Point _floatingDragStart;

		private int _floatingStartX;

		private int _floatingStartY;

		private int _floatingStartWidth;

		private int _floatingStartHeight;

		private float _floatingGrabX;

		private float _floatingGrabY;

		private bool _resizingImage;

		private DateTime _lastTextBoxInsertUtc;

		private DateTime _lastShapeInsertUtc;

		private DateTime _lastImageInsertUtc;

		private int _resizeHandle = -1;

		private Point _resizeStart;

		private float _resizeStartWidth;

		private float _resizeStartHeight;

		private int _spellUnavailableAttempts;

		public DocumentModel Document
		{
			get
			{
				return _document;
			}
			set
			{
				_document = value ?? CreateDefaultDocument();
				EnsureParagraphs();
				_document.EnsureSections();
				NormalizeDocumentText(_document);
				_caret = new CaretPos(0, 0);
				_anchor = _caret;
				_preferredX = -1;
				_undo.Clear();
				_redo.Clear();
				_typingFormat = GetDefaultFormat();
				InvalidateLayout(true);
				_typingFormat = FormatAt(_caret);
				_spellErrors.Clear();
				ScheduleSpellCheck();
				OnSelectionChanged();
			}
		}

		public float Zoom
		{
			get
			{
				return _zoom;
			}
			set
			{
				if (value < 0.25f)
				{
					value = 0.25f;
				}
				if (value > 4f)
				{
					value = 4f;
				}
				if (!(Math.Abs(value - _zoom) < 0.001f))
				{
					CommitCellEditor(false);
					_zoom = value;
					InvalidateLayout(true);
				}
			}
		}

		public int PageCount
		{
			get
			{
				EnsureLayout();
				return Math.Max(1, _pages.Count);
			}
		}

		public int CaretParagraph
		{
			get
			{
				if (_caret == null)
				{
					return 0;
				}
				return _caret.P;
			}
		}

		public Paragraph CurrentParagraph
		{
			get
			{
				EnsureParagraphs();
				return _document.Paragraphs[Math.Max(0, Math.Min(_document.Paragraphs.Count - 1, _caret.P))];
			}
		}

		public int CurrentPage
		{
			get
			{
				return _currentPage;
			}
		}

		public int ViewedPage
		{
			get
			{
				EnsureLayout();
				if (_pages.Count == 0)
				{
					return 0;
				}
				int num = -base.AutoScrollPosition.Y;
				int val = num + Math.Max(1, base.ClientSize.Height);
				int result = 0;
				int num2 = -1;
				for (int i = 0; i < _pages.Count; i++)
				{
					int num3 = Math.Max(num, (int)_pages[i].Top);
					int num4 = Math.Min(val, (int)(_pages[i].Top + _pages[i].Height));
					int num5 = Math.Max(0, num4 - num3);
					if (num5 > num2)
					{
						num2 = num5;
						result = i;
					}
				}
				return result;
			}
		}

		public ParagraphAlignment CurrentAlignment
		{
			get
			{
				EnsureParagraphs();
				int index = Math.Max(0, Math.Min(_caret.P, _document.Paragraphs.Count - 1));
				return _document.Paragraphs[index].Alignment;
			}
		}

		public bool CanUndo
		{
			get
			{
				return _undo.Count > 0;
			}
		}

		public bool CanRedo
		{
			get
			{
				return _redo.Count > 0;
			}
		}

		public bool HasSelection
		{
			get
			{
				if (_caret.P == _anchor.P)
				{
					return _caret.Offset != _anchor.Offset;
				}
				return true;
			}
		}

		public RunFormat CurrentFormat
		{
			get
			{
				if (!HasSelection)
				{
					return _typingFormat.Clone();
				}
				CaretPos start;
				CaretPos end;
				NormalizeSelection(out start, out end);
				Paragraph paragraph = _document.Paragraphs[start.P];
				for (int i = 0; i < paragraph.Inlines.Count; i++)
				{
					TextInline textInline = paragraph.Inlines[i] as TextInline;
					if (textInline != null)
					{
						return textInline.Format.Clone();
					}
				}
				return _typingFormat.Clone();
			}
		}

		public bool TrackChanges
		{
			get { return _document != null && _document.TrackChanges; }
			set
			{
				if (_document == null) return;
				_document.TrackChanges = value;
				OnSelectionChanged();
				Invalidate();
			}
		}

		public int RevisionCount
		{
			get { return DocumentRevisions.Count(_document); }
		}

		public bool SetTrackChanges(bool enabled)
		{
			TrackChanges = enabled;
			return TrackChanges == enabled;
		}

		public bool AcceptAllRevisions()
		{
			bool changed = DocumentRevisions.AcceptAll(_document);
			if (!changed) return false;
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
			return true;
		}

		public bool RejectAllRevisions()
		{
			bool changed = DocumentRevisions.RejectAll(_document);
			if (!changed) return false;
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
			return true;
		}

		public void InsertField(FieldKind kind)
		{
			if (_document == null) return;
			TextInline field = FieldText.Create(kind, _document, PageForParagraphIndex(_caret.P), PageCount);
			if (string.IsNullOrEmpty(field.Text)) field.Text = " ";
			if (_document.TrackChanges) DocumentRevisions.MarkInsert(field);
			PushUndo();
			InsertAtomAtCaret(field);
		}

		private void InsertAtomAtCaret(TextInline atom)
		{
			if (atom == null) return;
			if (HasSelection) DeleteSelectionInternal();
			Paragraph p = _document.Paragraphs[_caret.P];
			int offset = 0;
			bool inserted = false;
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline run = p.Inlines[i] as TextInline;
				if (run == null) continue;
				int length = run.Text.Length;
				if (_caret.Offset == offset)
				{
					p.Inlines.Insert(i, atom); inserted = true; break;
				}
				if (_caret.Offset < offset + length)
				{
					int local = _caret.Offset - offset;
					TextInline before = run.Clone();
					before.Text = run.Text.Substring(0, local);
					TextInline after = run.Clone();
					after.Text = run.Text.Substring(local);
					p.Inlines.RemoveAt(i);
					p.Inlines.Insert(i, before);
					p.Inlines.Insert(i + 1, atom);
					p.Inlines.Insert(i + 2, after);
					inserted = true; break;
				}
				offset += length;
			}
			if (!inserted) p.Inlines.Add(atom);
			_caret = new CaretPos(_caret.P, _caret.Offset + atom.Text.Length);
			_anchor = _caret;
			InvalidateLayout(true);
			AfterEdit();
		}

		public bool InsertTableOfContents(int levels)
		{
			if (_document == null) return false;
			EnsureHandleLayout();
			int at = Math.Max(0, Math.Min(_document.Paragraphs.Count, _caret.P + 1));
			at = TableOfContents.RemoveExisting(_document, at);
			List<Paragraph> entries = TableOfContents.Build(_document, levels, PageForParagraphIndex);
			int added = entries.Count;
			for (int i = 0; i < entries.Count; i++)
				if (entries[i].TocTargetParagraph >= at) entries[i].TocTargetParagraph += added;
			_document.Paragraphs.InsertRange(at, entries);
			_caret = new CaretPos(Math.Min(at + entries.Count, _document.Paragraphs.Count - 1), 0);
			_anchor = _caret;
			InvalidateLayout(true);
			TableOfContents.RefreshPages(_document, PageForParagraphIndex);
			Invalidate();
			OnTextChanged();
			OnSelectionChanged();
			return entries.Count > 1;
		}

		public bool UpdateTableOfContents(int levels)
		{
			if (_document == null) return false;
			int at = -1;
			for (int i = 0; i < _document.Paragraphs.Count; i++)
				if (_document.Paragraphs[i].IsTocTitle || _document.Paragraphs[i].IsTocEntry) { at = i; break; }
			if (at < 0) return false;
			at = TableOfContents.RemoveExisting(_document, at);
			EnsureHandleLayout();
			List<Paragraph> entries = TableOfContents.Build(_document, levels, PageForParagraphIndex);
			int added = entries.Count;
			for (int i = 0; i < entries.Count; i++)
				if (entries[i].TocTargetParagraph >= at) entries[i].TocTargetParagraph += added;
			_document.Paragraphs.InsertRange(at, entries);
			InvalidateLayout(true);
			TableOfContents.RefreshPages(_document, PageForParagraphIndex);
			Invalidate();
			OnTextChanged();
			OnSelectionChanged();
			return entries.Count > 1;
		}

		private int PageForParagraphIndex(int paragraph)
		{
			EnsureHandleLayout();
			for (int i = 0; i < _lines.Count; i++)
				if (_lines[i].Paragraph == paragraph) return _lines[i].Page + 1;
			return 1;
		}
		public string SelectedText
		{
			get
			{
				CaretPos start;
				CaretPos end;
				NormalizeSelection(out start, out end);
				if (start.P == end.P)
				{
					return GetText(_document.Paragraphs[start.P]).Substring(start.Offset, end.Offset - start.Offset);
				}
				StringBuilder stringBuilder = new StringBuilder();
				stringBuilder.Append(GetText(_document.Paragraphs[start.P]).Substring(start.Offset));
				for (int i = start.P + 1; i < end.P; i++)
				{
					stringBuilder.AppendLine();
					stringBuilder.Append(GetText(_document.Paragraphs[i]));
				}
				stringBuilder.AppendLine();
				stringBuilder.Append(GetText(_document.Paragraphs[end.P]).Substring(0, end.Offset));
				return stringBuilder.ToString();
			}
		}

		public bool ShowRuler
		{
			get { return _showRuler; }
			set
			{
				if (_showRuler == value) return;
				_showRuler = value;
				InvalidateLayout(true);
			}
		}

		public bool SpellCheckEnabled
		{
			get
			{
				return _spellEnabled;
			}
			set
			{
				_spellEnabled = value;
				if (!value)
				{
					_spellErrors.Clear();
				}
				else
				{
					ScheduleSpellCheck();
				}
				Invalidate();
			}
		}

		public string SpellLanguage
		{
			get
			{
				return _spellLanguage;
			}
			set
			{
				if (!string.IsNullOrEmpty(value) && !(value == _spellLanguage))
				{
					_spellLanguage = value;
					if (_spellChecker != null)
					{
						_spellChecker.Dispose();
						_spellChecker = null;
					}
					_spellErrors.Clear();
					ScheduleSpellCheck();
					Invalidate();
				}
			}
		}

		public int TableAtCaret
		{
			get
			{
				if (_caret.P >= 0 && _caret.P < _document.Paragraphs.Count && _document.Paragraphs[_caret.P].TableRows != null)
				{
					return _caret.P;
				}
				return -1;
			}
		}

		public bool HasTableCellSelection
		{
			get
			{
				return _tableSelectionParagraph >= 0;
			}
		}

		public bool CanMergeTableCells
		{
			get
			{
				if (_tableSelectionParagraph >= 0)
				{
					return _tableSelectionStartRow != _tableSelectionEndRow
						|| _tableSelectionStartColumn != _tableSelectionEndColumn;
				}
				return false;
			}
		}

		public new event EventHandler TextChanged;

		public event EventHandler SelectionChanged;

		public event EventHandler PageCountChanged;

		public event EventHandler ViewChanged;

		public NativeDocumentEditor()
		{
			SetStyle(ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
			AccessibleName = "Documento";
			AccessibleDescription = "Área de edición del documento. Use las flechas para moverse y el teclado para escribir.";
			AccessibleRole = AccessibleRole.Document;
			AutoScroll = true;
			BackColor = Color.FromArgb(191, 203, 217);
			base.TabStop = true;
			Font = new Font("Segoe UI", 9f);
			_document = CreateDefaultDocument();
			_caret = new CaretPos(0, 0);
			_anchor = _caret;
			_blinkTimer = new Timer();
			_blinkTimer.Interval = 530;
			Timer blinkTimer = _blinkTimer;
			EventHandler value = delegate
			{
				_caretOn = !_caretOn;
				InvalidateCaret();
			};
			blinkTimer.Tick += value;
			_blinkTimer.Start();
			_spellTimer = new Timer();
			_spellTimer.Interval = 650;
			_spellTimer.Tick += delegate
			{
				_spellTimer.Stop();
				RunSpellCheck();
			};
		}

		private static StringFormat CreateMeasureFormat()
		{
			StringFormat genericTypographic = StringFormat.GenericTypographic;
			genericTypographic.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
			return genericTypographic;
		}

		private static DocumentModel CreateDefaultDocument()
		{
			DocumentModel documentModel = new DocumentModel();
			documentModel.PageWidthTwips = 11906;
			documentModel.PageHeightTwips = 16838;
			documentModel.MarginLeftTwips = 1417;
			documentModel.MarginRightTwips = 1417;
			documentModel.MarginTopTwips = 1417;
			documentModel.MarginBottomTwips = 1417;
			documentModel.HeaderDistanceTwips = 720;
			documentModel.FooterDistanceTwips = 720;
			Paragraph paragraph = new Paragraph();
			paragraph.LineSpacing = 1;
			documentModel.Paragraphs.Add(paragraph);
			return documentModel;
		}

		private void EnsureParagraphs()
		{
			if (_document.Paragraphs.Count == 0)
			{
				Paragraph paragraph = new Paragraph();
				paragraph.LineSpacing = 1;
				_document.Paragraphs.Add(paragraph);
			}
		}

		private void InvalidateLayout(bool resetImages)
		{
			_dirtyLayout = true;
			PerformLayoutNow();
			Invalidate();
		}

		private void EnsureLayout()
		{
			if (_dirtyLayout)
			{
				PerformLayoutNow();
			}
		}

		private void EnsureHandleLayout()
		{
			if (base.IsHandleCreated && _dirtyLayout)
			{
				PerformLayoutNow();
			}
		}

		protected override void OnHandleCreated(EventArgs e)
		{
			base.OnHandleCreated(e);
			PerformLayoutNow();
		}

		protected override void OnResize(EventArgs e)
		{
			base.OnResize(e);
			PerformLayoutNow();
			Invalidate();
		}

		private void PerformLayoutNow()
		{
			if (_rebuilding || !base.IsHandleCreated)
			{
				return;
			}
			_rebuilding = true;
			Graphics graphics = null;
			try
			{
				graphics = (_layoutGraphics = CreateGraphics());
				_fonts.Clear();
				_pages.Clear();
				_lines.Clear();
				_tables.Clear();
				if (_document == null)
				{
					_document = CreateDefaultDocument();
				}
				EnsureParagraphs();
				_typingFormat = _typingFormat ?? GetDefaultFormat();
				_document.EnsureSections();
				LayoutPage page = NewPage(0, _document.SectionForParagraph(0));
				float num = page.BodyTop;
				float num2 = page.Width;
				for (int i = 0; i < _document.Paragraphs.Count; i++)
				{
					Paragraph paragraph = _document.Paragraphs[i];
					DocumentSection documentSection = ((paragraph.SectionIndex >= 0 && paragraph.SectionIndex < _document.Sections.Count) ? _document.Sections[paragraph.SectionIndex] : _document.SectionForParagraph(i));
					if (!object.ReferenceEquals(documentSection, page.Section))
					{
						page = NewPage(_pages.Count, documentSection);
						num = page.BodyTop;
					}
					if (page.Width > num2)
					{
						num2 = page.Width;
					}
					if (paragraph.PageBreakBefore && num > page.BodyTop + 0.5f)
					{
						page = NewPage(_pages.Count, page.Section);
						num = page.BodyTop;
					}
					num += (float)paragraph.SpaceBeforeTwips * (1f / 15f) * _zoom;
					if (num + 8f > page.BodyBottom)
					{
						AdvanceFlow(ref page, ref num);
					}
					num += AddFloatingObjects(paragraph, page, num);
					if (paragraph.TableRows != null)
					{
						AddTable(i, paragraph, ref page, ref num, page.Width, page.FullBodyLeft, page.BodyTop - page.Top, page.Width - page.FullBodyRight, page.Height - (page.BodyBottom - page.Top), page.ColumnWidth);
						continue;
					}
					AddParagraph(i, paragraph, ref page, ref num, page.Width, page.BodyLeft, page.BodyTop - page.Top, page.Width - page.BodyRight, page.Height - (page.BodyBottom - page.Top), page.ColumnWidth);
					num += (float)paragraph.SpaceAfterTwips * (1f / 15f) * _zoom;
				}
				float num3 = ((_pages.Count == 0) ? 18f : (_pages[_pages.Count - 1].Top + _pages[_pages.Count - 1].Height + 18f));
				int num4 = (int)Math.Ceiling(num2 + 36f);
				int num5 = (int)Math.Ceiling(num3);
				Point keptScroll = base.AutoScrollPosition;
				if (base.AutoScrollMinSize.Width != num4 || base.AutoScrollMinSize.Height != num5)
				{
					base.AutoScrollMinSize = new Size(num4, num5);
					// Changing the virtual size can make WinForms reset the
					// viewport.  Keep the user's current scroll position when a
					// paragraph reflow (for example the first indent) changes
					// the document height.
					if (keptScroll.X != 0 || keptScroll.Y != 0)
						base.AutoScrollPosition = new Point(-keptScroll.X, -keptScroll.Y);
				}
				_dirtyLayout = false;
				_currentPage = PageForPosition(_caret);
				if (this.PageCountChanged != null)
				{
					this.PageCountChanged(this, EventArgs.Empty);
				}
			}
			finally
			{
				_layoutGraphics = null;
				if (graphics != null)
				{
					graphics.Dispose();
				}
				_rebuilding = false;
			}
		}

		private LayoutPage NewPage(int index, DocumentSection section)
		{
			float num = (float)section.PageWidthTwips * (1f / 15f) * _zoom;
			float num2 = (float)section.PageHeightTwips * (1f / 15f) * _zoom;
			if (num < 100f)
			{
				num = 612f;
			}
			if (num2 < 100f)
			{
				num2 = 792f;
			}
			LayoutPage layoutPage = new LayoutPage();
			layoutPage.Index = index;
			layoutPage.Section = section;
			layoutPage.Width = num;
			layoutPage.Height = num2;
			layoutPage.Left = 18f;
			if ((float)base.ClientSize.Width > num + 36f)
			{
				layoutPage.Left = ((float)base.ClientSize.Width - num) / 2f;
			}
			layoutPage.Top = ((_pages.Count == 0) ? 18f : (_pages[_pages.Count - 1].Top + _pages[_pages.Count - 1].Height + 18f));
			layoutPage.BodyTop = layoutPage.Top + (float)section.MarginTopTwips * (1f / 15f) * _zoom;
			layoutPage.BodyBottom = layoutPage.Top + num2 - (float)section.MarginBottomTwips * (1f / 15f) * _zoom;
			layoutPage.FullBodyLeft = layoutPage.Left + (float)section.MarginLeftTwips * (1f / 15f) * _zoom;
			layoutPage.FullBodyRight = layoutPage.Left + num - (float)section.MarginRightTwips * (1f / 15f) * _zoom;
			layoutPage.ColumnCount = Math.Max(1, section.ColumnCount);
			layoutPage.ColumnIndex = 0;
			SetColumnBounds(layoutPage, section, 0);
			_pages.Add(layoutPage);
			return layoutPage;
		}

		private void SetColumnBounds(LayoutPage page, DocumentSection section, int columnIndex)
		{
			if (page != null && section != null)
			{
				int num = Math.Max(1, page.ColumnCount);
				if (columnIndex < 0)
				{
					columnIndex = 0;
				}
				if (columnIndex >= num)
				{
					columnIndex = num - 1;
				}
				float num2 = (float)section.ColumnSpacingTwips * (1f / 15f) * _zoom;
				float num3 = Math.Max(40f, page.FullBodyRight - page.FullBodyLeft);
				float num4 = (num3 - (float)(num - 1) * num2) / (float)num;
				if (num4 < 30f)
				{
					num4 = 30f;
				}
				page.ColumnIndex = columnIndex;
				page.ColumnWidth = num4;
				page.BodyLeft = page.FullBodyLeft + (float)columnIndex * (num4 + num2);
				page.BodyRight = page.BodyLeft + num4;
			}
		}

		private bool AdvanceFlow(ref LayoutPage page, ref float y)
		{
			if (page.ColumnIndex + 1 < page.ColumnCount)
			{
				SetColumnBounds(page, page.Section, page.ColumnIndex + 1);
				y = page.BodyTop;
				return true;
			}
			page = NewPage(_pages.Count, page.Section);
			y = page.BodyTop;
			return true;
		}

		private float AddFloatingObjects(Paragraph p, LayoutPage page, float y)
		{
			float num = 0f;
			for (int i = 0; i < p.FloatingObjects.Count; i++)
			{
				FloatingObject floatingObject = p.FloatingObjects[i];
				if (floatingObject != null)
				{
					LayoutFloating layoutFloating = new LayoutFloating();
					layoutFloating.Object = floatingObject;
					layoutFloating.ColumnIndex = page.ColumnIndex;
					layoutFloating.Width = Math.Max(24f, (float)floatingObject.WidthTwips * (1f / 15f) * _zoom);
					layoutFloating.Height = Math.Max(16f, (float)floatingObject.HeightTwips * (1f / 15f) * _zoom);
					if (layoutFloating.Width > page.ColumnWidth)
					{
						float num2 = page.ColumnWidth / layoutFloating.Width;
						layoutFloating.Width = page.ColumnWidth;
						layoutFloating.Height *= num2;
					}
					// Offsets are page-relative.  This is also the coordinate
					// system used by VML page-relative margins and by DrawingML
					// positionH/positionV with relativeFrom="page".
					layoutFloating.X = page.Left + (float)floatingObject.OffsetXTwips * (1f / 15f) * _zoom;
					layoutFloating.Y = page.Top + (float)floatingObject.OffsetYTwips * (1f / 15f) * _zoom;
					if (layoutFloating.X + layoutFloating.Width > page.Left + page.Width)
					{
						layoutFloating.X = Math.Max(page.Left, page.Left + page.Width - layoutFloating.Width);
					}
					if (layoutFloating.X < page.Left)
					{
						layoutFloating.X = page.Left;
					}
					if (floatingObject.Wrap == TextWrapMode.TopAndBottom)
					{
						num = Math.Max(num, layoutFloating.Y - y + layoutFloating.Height + 6f);
					}
					page.Floatings.Add(layoutFloating);
				}
			}
			return num;
		}

		private void GetLineWrap(LayoutPage page, float y, Paragraph p, float contentW, bool firstLine, out float x, out float available)
		{
			float num = (float)p.LeftIndentTwips * (1f / 15f) * _zoom;
			float num2 = (float)p.RightIndentTwips * (1f / 15f) * _zoom;
			if (num < 0f)
			{
				num = 0f;
			}
			if (num2 < 0f)
			{
				num2 = 0f;
			}
			if ((p.Bullet || p.Numbered) && num < 1f)
			{
				num = 28f;
			}
			float num3 = page.BodyLeft + num;
			if (firstLine && !p.Bullet)
			{
				num3 += (float)p.FirstLineIndentTwips * (1f / 15f) * _zoom;
			}
			float num4 = page.BodyRight - num2;
			float defaultLineHeight = GetDefaultLineHeight();
			for (int i = 0; i < page.Floatings.Count; i++)
			{
				LayoutFloating layoutFloating = page.Floatings[i];
				if (layoutFloating.ColumnIndex == page.ColumnIndex && (layoutFloating.Object.Wrap == TextWrapMode.Square || layoutFloating.Object.Wrap == TextWrapMode.Tight || layoutFloating.Object.Wrap == TextWrapMode.Through) && !(layoutFloating.Y + layoutFloating.Height <= y) && !(layoutFloating.Y >= y + defaultLineHeight))
				{
					float num5 = layoutFloating.X;
					float num6 = layoutFloating.X + layoutFloating.Width;
					if (num5 <= page.BodyLeft + (page.BodyRight - page.BodyLeft) * 0.55f)
					{
						num3 = Math.Max(num3, num6 + 8f);
					}
					else
					{
						num4 = Math.Min(num4, num5 - 8f);
					}
				}
			}
			if (num3 < page.BodyLeft)
			{
				num3 = page.BodyLeft;
			}
			if (num4 > page.BodyRight)
			{
				num4 = page.BodyRight;
			}
			x = num3;
			available = Math.Max(30f, num4 - num3);
		}

		private LineBuilder StartWrappedLine(ref LayoutPage page, ref float y, Paragraph p, int paragraphIndex, float contentW, bool firstLine, out float available)
		{
			LayoutPage objB = page;
			float num = y;
			// A list marker belongs to the paragraph, not to each wrapped
			// visual line.  Without this guard every continuation line gets a
			// new bullet/number and looks like a new paragraph.
			bool bullet = (p.Bullet || p.Numbered) && firstLine;
			float num2;
			GetLineWrap(page, y, p, contentW, firstLine, out num2, out available);
			LineBuilder lineBuilder = StartLine(ref page, ref y, num2, bullet, num2 - 18f);
			if (!object.ReferenceEquals(page, objB) || lineBuilder.Y != num)
			{
				GetLineWrap(page, lineBuilder.Y, p, contentW, firstLine, out num2, out available);
				lineBuilder.X = num2;
				lineBuilder.BulletX = num2 - 18f;
			}
			if (firstLine && p.Numbered)
			{
				lineBuilder.Number = GetListNumber(paragraphIndex) + ".";
			}
			return lineBuilder;
		}

		private int GetListNumber(int paragraphIndex)
		{
			int num = 1;
			for (int num2 = paragraphIndex - 1; num2 >= 0; num2--)
			{
				Paragraph paragraph = _document.Paragraphs[num2];
				if (!paragraph.Numbered)
				{
					break;
				}
				num++;
			}
			return num;
		}

		private void AddParagraph(int paragraphIndex, Paragraph p, ref LayoutPage page, ref float y, float pageW, float bodyLeft, float top, float right, float bottom, float contentW)
		{
			List<Token> list = Tokenize(p);
			float available;
			if (p.Inlines.Count == 0)
			{
				LineBuilder build = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, true, out available);
				FinishLine(build, paragraphIndex, p, page, pageW, bodyLeft, contentW, true, ref y);
				return;
			}
			LineBuilder line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, true, out available);
			for (int i = 0; i < list.Count; i++)
			{
				Token token = list[i];
				if (token.Newline)
				{
					FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, false, ref y);
					line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, false, out available);
					continue;
				}
				if (token.Tab)
				{
					float num = TabWidth(line.Width, p, _zoom);
					if (line.HasContent && line.Width + num > available)
					{
						FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, false, ref y);
						line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, false, out available);
					}
					else
					{
						AddGap(line, num, token.Start, token.End);
					}
					continue;
				}
				if (token.Image != null)
				{
					float num2 = (float)token.Image.WidthTwips * (1f / 15f) * _zoom;
					float num3 = (float)token.Image.HeightTwips * (1f / 15f) * _zoom;
					if (num2 > available)
					{
						float num4 = available / num2;
						num2 = available;
						num3 *= num4;
					}
					if (line.HasContent && line.Width + num2 > available)
					{
						FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, false, ref y);
						line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, false, out available);
					}
					EnsureLineSpace(ref page, ref y, num3 + 4f, pageW, bodyLeft, top, right, bottom);
					LayoutSegment layoutSegment = new LayoutSegment();
					layoutSegment.Image = token.Image;
					layoutSegment.X = line.X + line.Width;
					layoutSegment.Y = line.Y + 2f;
					layoutSegment.Width = num2;
					layoutSegment.Height = num3;
					layoutSegment.Start = token.Start;
					layoutSegment.End = token.End;
					line.Segments.Add(layoutSegment);
					line.Width += num2;
					line.HasContent = true;
					if (num3 > line.Height)
					{
						line.Height = num3;
					}
					line.End = token.End;
					continue;
				}
				Font font = GetFont(token.Format, token.Text);
				float num5 = Measure(token.Text, font);
				float fontHeight = font.GetHeight(_layoutGraphics);
				if (token.IsSpace && !line.HasContent)
				{
					continue;
				}
				if (line.HasContent && line.Width + num5 > available)
				{
					FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, false, ref y);
					line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, false, out available);
					if (token.IsSpace)
					{
						continue;
					}
				}
				if (num5 > available && !line.HasContent)
				{
					AddLongToken(token, ref line, paragraphIndex, p, ref page, ref y, pageW, bodyLeft, top, right, bottom, contentW, ref available);
					if (!line.HasContent)
					{
					}
				}
				else if (num5 > 0f)
				{
					AddTextSegment(line, token, font, fontHeight);
				}
			}
			if (line.HasContent || line.Segments.Count == 0)
			{
				FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, true, ref y);
			}
		}

		private void AddLongToken(Token token, ref LineBuilder line, int paragraphIndex, Paragraph p, ref LayoutPage page, ref float y, float pageW, float bodyLeft, float top, float right, float bottom, float contentW, ref float available)
		{
			string text = token.Text ?? string.Empty;
			for (int i = 0; i < text.Length; i++)
			{
				char c = text[i];
				string text2 = c.ToString();
				Font font = GetFont(token.Format, text2);
				float num = Measure(text2, font);
				if (line.HasContent && line.Width + num > available)
				{
					FinishLine(line, paragraphIndex, p, page, pageW, bodyLeft, contentW, false, ref y);
					line = StartWrappedLine(ref page, ref y, p, paragraphIndex, contentW, false, out available);
				}
				Token token2 = new Token();
				token2.Text = text2;
				token2.Format = token.Format;
				token2.Start = token.Start + i;
				token2.End = token2.Start + 1;
				token2.IsSpace = char.IsWhiteSpace(c);
				AddTextSegment(line, token2, font, font.GetHeight(_layoutGraphics));
			}
			line.End = token.End;
		}

		private void AddTextSegment(LineBuilder line, Token token, Font font, float fontHeight)
		{
			LayoutSegment layoutSegment = new LayoutSegment();
			layoutSegment.Text = token.Text;
			layoutSegment.Format = token.Format;
			layoutSegment.Note = token.Note;
			layoutSegment.Comment = token.Comment;
			layoutSegment.Start = token.Start;
			layoutSegment.End = token.End;
			layoutSegment.X = line.X + line.Width;
			layoutSegment.Y = line.Y;
			if (token.Format.Script > 0)
			{
				layoutSegment.Y -= fontHeight * 0.35f;
			}
			else if (token.Format.Script < 0)
			{
				layoutSegment.Y += fontHeight * 0.15f;
			}
			layoutSegment.Width = Measure(token.Text, font);
			layoutSegment.Height = fontHeight;
			layoutSegment.Font = font;
			layoutSegment.IsSpace = token.IsSpace;
			line.Segments.Add(layoutSegment);
			line.Width += layoutSegment.Width;
			if (fontHeight > line.Height)
			{
				line.Height = fontHeight;
			}
			if (!token.IsSpace)
			{
				line.HasContent = true;
			}
			if (line.Start < 0)
			{
				line.Start = token.Start;
			}
			line.End = token.End;
		}

		private void AddGap(LineBuilder line, float width, int start, int end)
		{
			LayoutSegment layoutSegment = new LayoutSegment();
			layoutSegment.Text = string.Empty;
			layoutSegment.Width = width;
			layoutSegment.Height = line.Height;
			layoutSegment.X = line.X + line.Width;
			layoutSegment.Start = start;
			layoutSegment.End = end;
			line.Segments.Add(layoutSegment);
			line.Width += width;
			if (line.Start < 0)
			{
				line.Start = start;
			}
			line.End = end;
		}

		private LineBuilder StartLine(ref LayoutPage page, ref float y, float x, bool bullet, float bulletX)
		{
			float defaultLineHeight = GetDefaultLineHeight();
			if (y + defaultLineHeight > page.BodyBottom)
			{
				float bodyLeft = page.BodyLeft;
				AdvanceFlow(ref page, ref y);
				float num = page.BodyLeft - bodyLeft;
				x += num;
				bulletX += num;
			}
			LineBuilder lineBuilder = new LineBuilder();
			lineBuilder.X = x;
			lineBuilder.Y = y;
			lineBuilder.Height = defaultLineHeight;
			lineBuilder.Bullet = bullet;
			lineBuilder.BulletX = bulletX;
			lineBuilder.Start = -1;
			return lineBuilder;
		}

		private void EnsureLineSpace(ref LayoutPage page, ref float y, float height, float pageW, float bodyLeft, float top, float right, float bottom)
		{
			if (y + height > page.BodyBottom)
			{
				AdvanceFlow(ref page, ref y);
			}
		}

		private void FinishLine(LineBuilder build, int paragraphIndex, Paragraph p, LayoutPage page, float pageW, float bodyLeft, float contentW, bool paragraphEnd, ref float y)
		{
			if (build == null)
			{
				return;
			}
			if (build.Height < 1f)
			{
				build.Height = GetDefaultLineHeight();
			}
			if (p.LineSpacing == 1)
			{
				build.Height *= 1.5f;
			}
			else if (p.LineSpacing >= 2)
			{
				build.Height *= 2f;
			}
			if (build.Width < 1f && build.Segments.Count == 0)
			{
				build.End = ((build.Start >= 0) ? build.Start : 0);
			}
			if (p.Alignment != 0 && !p.Bullet)
			{
				float num = Math.Max(0f, contentW - (float)p.LeftIndentTwips * (1f / 15f) * _zoom - (float)p.RightIndentTwips * (1f / 15f) * _zoom);
				float num2 = num - build.Width;
				if (num2 > 0f)
				{
					if (p.Alignment == ParagraphAlignment.Justify && !paragraphEnd)
					{
						float num3 = 0f;
						for (int i = 0; i < build.Segments.Count; i++)
						{
							if (build.Segments[i].IsSpace)
							{
								num3 += build.Segments[i].Width;
							}
						}
						if (num3 > 0.5f)
						{
							float num4 = 0f;
							for (int j = 0; j < build.Segments.Count; j++)
							{
								LayoutSegment layoutSegment = build.Segments[j];
								layoutSegment.X += num4;
								if (layoutSegment.IsSpace)
								{
									float num5 = num2 * layoutSegment.Width / num3;
									layoutSegment.Width += num5;
									num4 += num5;
								}
							}
							build.Width = num;
						}
					}
					else if (p.Alignment == ParagraphAlignment.Center || p.Alignment == ParagraphAlignment.Right)
					{
						float num6 = ((p.Alignment == ParagraphAlignment.Center) ? (num2 / 2f) : num2);
						for (int k = 0; k < build.Segments.Count; k++)
						{
							build.Segments[k].X += num6;
						}
					}
				}
			}
			LayoutLine layoutLine = new LayoutLine();
			layoutLine.Page = page.Index;
			layoutLine.Paragraph = paragraphIndex;
			layoutLine.Start = ((build.Start >= 0) ? build.Start : 0);
			layoutLine.End = build.End;
			layoutLine.X = build.X;
			layoutLine.Y = build.Y;
			layoutLine.Width = build.Width;
			layoutLine.Height = build.Height;
			layoutLine.First = paragraphEnd && page.Lines.Count == 0;
			layoutLine.Bullet = (p.Bullet || p.Numbered) && build.Bullet;
			layoutLine.Number = build.Number;
			layoutLine.BulletX = build.BulletX;
			for (int l = 0; l < build.Segments.Count; l++)
			{
				layoutLine.Segments.Add(build.Segments[l]);
			}
			page.Lines.Add(layoutLine);
			_lines.Add(layoutLine);
			y = build.Y + build.Height;
		}

		private RunFormat DisplayFormat(TextInline run)
		{
			RunFormat format = run == null || run.Format == null ? new RunFormat() : run.Format.Clone();
			if (run != null && _document != null && _document.ShowRevisions && run.Revision != RevisionKind.None)
			{
				if (run.Revision == RevisionKind.Insert)
				{
					format.Underline = true;
					format.ForeColor = Color.FromArgb(0, 128, 0);
				}
				else
				{
					format.Strikeout = true;
					format.ForeColor = Color.FromArgb(180, 0, 0);
				}
			}
			return format;
		}

		private List<Token> Tokenize(Paragraph p)
		{
			List<Token> list = new List<Token>();
			int num = 0;
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline textInline = p.Inlines[i] as TextInline;
				if (textInline != null)
				{
					string text = textInline.Text ?? string.Empty;
					int j = 0;
					while (j < text.Length)
					{
						char c = text[j];
						switch (c)
						{
						case '\n':
						{
							Token token = new Token();
							token.Newline = true;
							token.Start = num;
							token.End = num + 1;
							list.Add(token);
							num++;
							j++;
							continue;
						}
						case '\t':
						{
							Token token2 = new Token();
							token2.Tab = true;
							token2.Start = num;
							token2.End = num + 1;
							list.Add(token2);
							num++;
							j++;
							continue;
						}
						}
						if (char.IsWhiteSpace(c))
						{
							int num2 = j;
							for (; j < text.Length && char.IsWhiteSpace(text[j]) && text[j] != '\t' && text[j] != '\n'; j++)
							{
							}
							Token token3 = new Token();
							token3.Text = text.Substring(num2, j - num2);
							token3.Format = DisplayFormat(textInline);
							token3.Start = num;
							token3.End = num + token3.Text.Length;
							token3.IsSpace = true;
							list.Add(token3);
							num += token3.Text.Length;
						}
						else
						{
							int num3 = j;
							for (; j < text.Length && !char.IsWhiteSpace(text[j]); j++)
							{
							}
							Token token4 = new Token();
							token4.Text = text.Substring(num3, j - num3);
							token4.Format = DisplayFormat(textInline);
							token4.Start = num;
							token4.End = num + token4.Text.Length;
							list.Add(token4);
							num += token4.Text.Length;
						}
					}
				}
				ImageInline imageInline = p.Inlines[i] as ImageInline;
				if (imageInline != null)
				{
					Token token5 = new Token();
					token5.Image = imageInline;
					token5.Format = imageInline.Format;
					token5.Start = num;
					token5.End = num;
					list.Add(token5);
				}
				NoteInline noteInline = p.Inlines[i] as NoteInline;
				if (noteInline != null)
				{
					Token token6 = new Token();
					token6.Note = noteInline.Note;
					token6.Text = "[" + ((noteInline.Note == null) ? 1 : noteInline.Note.Id) + "]";
					token6.Format = noteInline.Format;
					token6.Start = num;
					token6.End = num;
					list.Add(token6);
				}
				CommentInline commentInline = p.Inlines[i] as CommentInline;
				if (commentInline != null)
				{
					Token token7 = new Token();
					token7.Comment = commentInline.Comment;
					token7.Text = "[C]";
					token7.Format = commentInline.Format;
					token7.Start = num;
					token7.End = num;
					list.Add(token7);
				}
			}
			return list;
		}

		private Font GetFont(RunFormat format)
		{
			return GetFont(format, null);
		}

		private Font GetFont(RunFormat format, string text)
		{
			float num = (float)format.SizeHalfPoints / 2f * 96f / 72f * _zoom;
			if (format.Script != 0)
			{
				num *= 0.66f;
			}
			FontStyle fontStyle = FontStyle.Regular;
			if (format.Bold)
			{
				fontStyle |= FontStyle.Bold;
			}
			if (format.Italic)
			{
				fontStyle |= FontStyle.Italic;
			}
			if (format.Underline)
			{
				fontStyle |= FontStyle.Underline;
			}
			if (format.Strikeout)
			{
				fontStyle |= FontStyle.Strikeout;
			}
			string family = FontResolver.Resolve(format.FontFamily, text);
			return _fonts.Get(family, num, fontStyle);
		}

		private Graphics MeasurementGraphics()
		{
			if (_layoutGraphics != null) return _layoutGraphics;
			if (_measureGraphics == null) _measureGraphics = CreateGraphics();
			return _measureGraphics;
		}

		private float Measure(string text, Font font)
		{
			if (string.IsNullOrEmpty(text)) return 0f;
			Graphics graphics = MeasurementGraphics();
			if (graphics == null) return 0f;
			try
			{
				return graphics.MeasureString(text, font, int.MaxValue, MeasureFormat).Width;
			}
			catch
			{
				if (_measureGraphics != null)
				{
					_measureGraphics.Dispose();
					_measureGraphics = null;
				}
				return 0f;
			}
		}

		private float GetDefaultLineHeight()
		{
			// This is only the initial height of a line.  It must not follow
			// _typingFormat: after changing one fragment to 30 pt the typing
			// format would make every paragraph use a 30 pt line box and the
			// whole document would look reformatted.
			RunFormat format = new RunFormat("Calibri", 22);
			Font font = GetFont(format);
			if (_layoutGraphics == null)
			{
				return 18f;
			}
			return Math.Max(8f, font.GetHeight(_layoutGraphics));
		}

		private static float TabWidth(float current, Paragraph p, float zoom)
		{
			float num = 48f * zoom;
			if (p.TabStopsTwips != null && p.TabStopsTwips.Length > 0)
			{
				float num2 = -1f;
				for (int i = 0; i < p.TabStopsTwips.Length; i++)
				{
					float num3 = (float)p.TabStopsTwips[i] * (1f / 15f) * zoom;
					if (num3 > current + 0.5f && (num2 < 0f || num3 < num2))
					{
						num2 = num3;
					}
				}
				if (num2 > 0f)
				{
					return Math.Max(4f, num2 - current);
				}
			}
			float num4 = (float)(Math.Floor(current / num) + 1.0) * num;
			return Math.Max(4f, num4 - current);
		}

		private void AddTable(int paragraphIndex, Paragraph p, ref LayoutPage page, ref float y, float pageW, float bodyLeft, float top, float right, float bottom, float contentW)
		{
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length == 0)
			{
				return;
			}
			int num = 0;
			for (int i = 0; i < p.TableCells.Length; i++)
			{
				if (p.TableCells[i] != null && p.TableCells[i].Length > num)
				{
					num = p.TableCells[i].Length;
				}
			}
			if (num == 0)
			{
				return;
			}
			int num2 = p.TableCells.Length;
			float num3 = contentW / (float)num;
			float num4 = Math.Max(24f, GetDefaultLineHeight() + 8f);
			float[] rowHeights = new float[num2];
			for (int r = 0; r < num2; r++) rowHeights[r] = num4;
			for (int r = 0; r < num2; r++)
			{
				TableCell[] row = p.TableCells[r];
				if (row == null) continue;
				for (int c = 0; c < row.Length; c++)
				{
					TableCell cell = row[c];
					if (cell == null || cell.Covered || cell.Format == null) continue;
					Font cellFont = GetFont(cell.Format);
					float lineHeight = _layoutGraphics != null ? cellFont.GetHeight(_layoutGraphics) : num4 - 8f;
					rowHeights[r] = Math.Max(rowHeights[r], lineHeight + 8f);
				}
			}
			float[] rowTops = new float[num2];
			float num5 = 0f;
			for (int r = 0; r < num2; r++)
			{
				rowTops[r] = num5;
				num5 += rowHeights[r];
			}
			if (y + num5 > page.BodyBottom)
			{
				AdvanceFlow(ref page, ref y);
			}
			LayoutTable layoutTable = new LayoutTable();
			layoutTable.Page = page.Index;
			layoutTable.Paragraph = paragraphIndex;
			layoutTable.X = page.BodyLeft;
			layoutTable.Y = y;
			layoutTable.Width = contentW;
			layoutTable.Height = num5;
			layoutTable.RowHeight = num4;
			for (int j = 0; j < num2; j++)
			{
				LayoutTableRow layoutTableRow = new LayoutTableRow();
				layoutTableRow.Y = y + rowTops[j];
				layoutTableRow.Height = rowHeights[j];
				layoutTable.Rows.Add(layoutTableRow);
			}
			for (int k = 0; k < num2; k++)
			{
				TableCell[] array = p.TableCells[k] ?? new TableCell[0];
				for (int l = 0; l < array.Length; l++)
				{
					TableCell tableCell = array[l];
					if (tableCell != null && !tableCell.Covered)
					{
						int num6 = Math.Max(1, Math.Min(tableCell.RowSpan, num2 - k));
						int num7 = Math.Max(1, Math.Min(tableCell.ColSpan, num - l));
						LayoutTableCell layoutTableCell = new LayoutTableCell();
						layoutTableCell.Cell = tableCell;
						layoutTableCell.Row = k;
						layoutTableCell.Column = l;
						layoutTableCell.X = layoutTable.X + (float)l * num3;
						layoutTableCell.Y = layoutTable.Y + rowTops[k];
						layoutTableCell.Width = num3 * (float)num7;
						float cellHeight = 0f;
						for (int r = k; r < k + num6; r++) cellHeight += rowHeights[r];
						layoutTableCell.Height = cellHeight;
						layoutTable.Cells.Add(layoutTableCell);
					}
				}
			}
			page.Tables.Add(layoutTable);
			_tables.Add(layoutTable);
			y += num5 + 6f;
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			EnsureHandleLayout();
			if (_cellEditor != null && _cellEditor.Visible)
			{
				PositionCellEditor();
			}
			Graphics graphics = e.Graphics;
			graphics.SmoothingMode = SmoothingMode.AntiAlias;
			graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
			graphics.Clear(BackColor);
			int viewedPage = ViewedPage;
			if (viewedPage != _lastViewedPage)
			{
				_lastViewedPage = viewedPage;
				if (this.ViewChanged != null)
				{
					this.ViewChanged(this, EventArgs.Empty);
				}
			}
			int num = base.AutoScrollPosition.X;
			int num2 = base.AutoScrollPosition.Y;
			graphics.TranslateTransform(num, num2);
			using (Brush brush2 = new SolidBrush(Color.White))
			{
				using (Brush brush = new SolidBrush(Color.FromArgb(45, 30, 40, 50)))
				{
					using (Pen pen = new Pen(Color.FromArgb(185, 192, 201)))
					{
						for (int i = 0; i < _pages.Count; i++)
						{
							LayoutPage layoutPage = _pages[i];
							if (!(layoutPage.Top + layoutPage.Height + 4f < (float)(-num2)) && !(layoutPage.Top > (float)(-num2 + base.ClientSize.Height + 4)))
							{
								graphics.FillRectangle(brush, layoutPage.Left + 3f, layoutPage.Top + 4f, layoutPage.Width, layoutPage.Height);
								graphics.FillRectangle(brush2, layoutPage.Left, layoutPage.Top, layoutPage.Width, layoutPage.Height);
								graphics.DrawRectangle(pen, layoutPage.Left, layoutPage.Top, layoutPage.Width, layoutPage.Height);
								Region clip = graphics.Clip;
								graphics.SetClip(new RectangleF(layoutPage.Left, layoutPage.Top, layoutPage.Width, layoutPage.Height));
								DrawPage(graphics, layoutPage);
								graphics.Clip = clip;
							}
						}
						DrawImageHandles(graphics);
						DrawFloatingHandles(graphics);
					}
				}
			}
		}

		private void DrawRuler(Graphics g)
		{
			float rulerHeight = 22f;
			float pxPerCm = 96f / 2.54f * _zoom;
			LayoutPage page = null;
			if (_pages.Count > 0)
				page = _pages[Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage))];
			float pageLeft = 18f;
			float pageWidth = Math.Max(1f, base.ClientSize.Width - 36f);
			if (page != null)
			{
				pageLeft = page.Left + base.AutoScrollPosition.X;
				pageWidth = page.Width;
			}
			RectangleF ruler = new RectangleF(0f, 0f, base.ClientSize.Width, rulerHeight);
			using (Brush background = new SolidBrush(Color.FromArgb(241, 244, 248)))
			using (Brush margin = new SolidBrush(Color.FromArgb(226, 222, 210)))
			using (Pen edge = new Pen(Color.FromArgb(160, 155, 140)))
			using (Pen tick = new Pen(Color.FromArgb(95, 90, 80), 1f))
			using (Brush text = new SolidBrush(Color.FromArgb(70, 65, 55)))
			using (Font font = new Font("Segoe UI", 6.5f))
			{
				g.FillRectangle(background, ruler);
				if (page != null)
				{
					float leftMargin = page.BodyLeft - page.Left;
					float rightMargin = page.Left + page.Width - page.BodyRight;
					g.FillRectangle(margin, pageLeft, 0f, Math.Max(0, leftMargin), rulerHeight);
					g.FillRectangle(margin, pageLeft + pageWidth - Math.Max(0, rightMargin), 0f,
						Math.Max(0, rightMargin), rulerHeight);
					g.DrawRectangle(edge, pageLeft, 0f, pageWidth, rulerHeight);
					string widthLabel = (pageWidth / pxPerCm).ToString("0.0", CultureInfo.InvariantCulture) + " cm";
					SizeF measured = g.MeasureString(widthLabel, font);
					float labelLeft = pageLeft + pageWidth - measured.Width - 5f;
					int centimeters = (int)Math.Floor(pageWidth / pxPerCm);
					for (int i = 0; i <= centimeters; i++)
					{
						float x = pageLeft + i * pxPerCm;
						bool major = i % 5 == 0;
						float length = major ? 9f : 5f;
						g.DrawLine(tick, x, rulerHeight - length, x, rulerHeight);
						if (i < centimeters && x + 12f < labelLeft)
							g.DrawString(i.ToString(), font, text, x + 1f, 1f);
					}
					g.FillRectangle(background, labelLeft, 0f, pageLeft + pageWidth - labelLeft, rulerHeight);
					g.DrawRectangle(edge, labelLeft, 0f, pageLeft + pageWidth - labelLeft, rulerHeight);
					g.DrawString(widthLabel, font, text, labelLeft + 2f, 1f);
				}
				g.DrawLine(edge, 0f, rulerHeight - 1f, base.ClientSize.Width, rulerHeight - 1f);
			}
		}

		private void DrawPage(Graphics g, LayoutPage page)
		{
			DrawHeaderFooter(g, page, true);
			for (int i = 0; i < page.Floatings.Count; i++)
			{
				if (page.Floatings[i].Object.BehindText || page.Floatings[i].Object.Wrap == TextWrapMode.Behind)
				{
					DrawFloating(g, page.Floatings[i]);
				}
			}
			DrawSelection(g, page);
			DrawSpellingHighlight(g, page);
			for (int j = 0; j < page.Lines.Count; j++)
			{
				DrawLine(g, page.Lines[j]);
			}
			for (int k = 0; k < page.Tables.Count; k++)
			{
				DrawTable(g, page.Tables[k]);
			}
			for (int l = 0; l < page.Floatings.Count; l++)
			{
				if (!page.Floatings[l].Object.BehindText && page.Floatings[l].Object.Wrap != TextWrapMode.Behind)
				{
					DrawFloating(g, page.Floatings[l]);
				}
			}
			DrawFootnotes(g, page);
			DrawTableSelection(g, page);
			DrawCaret(g, page);
		}

		private void DrawFootnotes(Graphics g, LayoutPage page)
		{
			List<DocumentNote> list = new List<DocumentNote>();
			for (int i = 0; i < page.Lines.Count; i++)
			{
				LayoutLine layoutLine = page.Lines[i];
				for (int j = 0; j < layoutLine.Segments.Count; j++)
				{
					DocumentNote note = layoutLine.Segments[j].Note;
					if (note != null && !list.Contains(note))
					{
						list.Add(note);
					}
				}
			}
			if (list.Count != 0)
			{
				RunFormat format = new RunFormat("Calibri", 16);
				Font font = GetFont(format);
				float num = font.GetHeight(g) + 2f;
				float num2 = (float)list.Count * num + 6f;
				float num3 = page.BodyBottom - num2;
				using (Pen pen = new Pen(Color.FromArgb(120, 128, 138)))
				{
					g.DrawLine(pen, page.FullBodyLeft, num3 - 3f, page.FullBodyLeft + Math.Min(150f, page.FullBodyRight - page.FullBodyLeft), num3 - 3f);
				}
				for (int k = 0; k < list.Count; k++)
				{
					string s = "[" + list[k].Id + "] " + list[k].Text;
					g.DrawString(s, font, Brushes.Black, new RectangleF(page.FullBodyLeft, num3 + (float)k * num, page.FullBodyRight - page.FullBodyLeft, num), MeasureFormat);
				}
			}
		}

		public void DrawPageForPrint(Graphics g, int pageIndex, RectangleF target)
		{
			EnsureLayout();
			if (pageIndex < 0 || pageIndex >= _pages.Count)
			{
				return;
			}
			LayoutPage layoutPage = _pages[pageIndex];
			float val = target.Width / layoutPage.Width;
			float val2 = target.Height / layoutPage.Height;
			float num = Math.Min(val, val2);
			GraphicsState gstate = g.Save();
			try
			{
				g.TranslateTransform(target.X, target.Y);
				g.ScaleTransform(num, num);
				g.TranslateTransform(0f - layoutPage.Left, 0f - layoutPage.Top);
				g.SetClip(new RectangleF(layoutPage.Left, layoutPage.Top, layoutPage.Width, layoutPage.Height));
				DrawHeaderFooter(g, layoutPage, false);
				for (int i = 0; i < layoutPage.Floatings.Count; i++)
				{
					if (layoutPage.Floatings[i].Object.BehindText || layoutPage.Floatings[i].Object.Wrap == TextWrapMode.Behind)
					{
						DrawFloating(g, layoutPage.Floatings[i]);
					}
				}
				for (int j = 0; j < layoutPage.Lines.Count; j++)
				{
					DrawLine(g, layoutPage.Lines[j]);
				}
				for (int k = 0; k < layoutPage.Tables.Count; k++)
				{
					DrawTable(g, layoutPage.Tables[k]);
				}
				for (int l = 0; l < layoutPage.Floatings.Count; l++)
				{
					if (!layoutPage.Floatings[l].Object.BehindText && layoutPage.Floatings[l].Object.Wrap != TextWrapMode.Behind)
					{
						DrawFloating(g, layoutPage.Floatings[l]);
					}
				}
				DrawFootnotes(g, layoutPage);
			}
			finally
			{
				g.Restore(gstate);
			}
		}

		private void DrawHeaderFooter(Graphics g, LayoutPage page, bool design)
		{
			DocumentSection section = page.Section ?? _document.SectionForParagraph(0);
			string headerText = section != null ? section.HeaderText : _document.HeaderText;
			string footerText = section != null ? section.FooterText : _document.FooterText;
			bool firstPage = false;
			bool useFirst = false;
			bool useEven = false;
			if (section != null)
			{
				for (int i = 0; i < _pages.Count; i++)
				{
					if (object.ReferenceEquals(_pages[i].Section, section))
					{
						firstPage = _pages[i].Index == page.Index;
						break;
					}
				}
				useFirst = section.TitlePage && firstPage;
				useEven = section.DifferentOddEven && ((page.Index + 1) % 2) == 0;
				if (useFirst && !string.IsNullOrEmpty(section.FirstPageHeaderText))
					headerText = section.FirstPageHeaderText;
				if (useFirst && !string.IsNullOrEmpty(section.FirstPageFooterText))
					footerText = section.FirstPageFooterText;
				if (useEven)
				{
					if (!string.IsNullOrEmpty(section.EvenPageHeaderText)) headerText = section.EvenPageHeaderText;
					if (!string.IsNullOrEmpty(section.EvenPageFooterText)) footerText = section.EvenPageFooterText;
				}
			}
			if (section != null && section == _document.Sections[0])
			{
				if (string.IsNullOrEmpty(headerText)) headerText = _document.HeaderText;
				if (string.IsNullOrEmpty(footerText)) footerText = _document.FooterText;
			}
			RunFormat headerFormat = section != null && section.HeaderFormat != null
				? section.HeaderFormat : new RunFormat("Calibri", 18);
			RunFormat footerFormat = section != null && section.FooterFormat != null
				? section.FooterFormat : new RunFormat("Calibri", 18);
			ParagraphAlignment headerAlignment = section != null ? section.HeaderAlignment : ParagraphAlignment.Left;
			ParagraphAlignment footerAlignment = section != null ? section.FooterAlignment : ParagraphAlignment.Left;
			System.Collections.Generic.List<ImageInline> headerImages = section != null
				? section.HeaderImages : new System.Collections.Generic.List<ImageInline>();
			System.Collections.Generic.List<ImageInline> footerImages = section != null
				? section.FooterImages : new System.Collections.Generic.List<ImageInline>();
			if (section != null && useFirst)
			{
				headerFormat = section.FirstPageHeaderFormat ?? headerFormat;
				footerFormat = section.FirstPageFooterFormat ?? footerFormat;
				headerAlignment = section.FirstPageHeaderAlignment;
				footerAlignment = section.FirstPageFooterAlignment;
				headerImages = section.FirstPageHeaderImages;
				footerImages = section.FirstPageFooterImages;
			}
			else if (section != null && useEven)
			{
				headerFormat = section.EvenPageHeaderFormat ?? headerFormat;
				footerFormat = section.EvenPageFooterFormat ?? footerFormat;
				headerAlignment = section.EvenPageHeaderAlignment;
				footerAlignment = section.EvenPageFooterAlignment;
				headerImages = section.EvenPageHeaderImages;
				footerImages = section.EvenPageFooterImages;
			}
			int headerDistance = section != null ? section.HeaderDistanceTwips : _document.HeaderDistanceTwips;
			int footerDistance = section != null ? section.FooterDistanceTwips : _document.FooterDistanceTwips;
			float bodyLeft = page.FullBodyLeft;
			float bodyWidth = page.FullBodyRight - page.FullBodyLeft;
			float headerTop = page.Top + (float)headerDistance * (1f / 15f) * _zoom;
			float footerBottom = page.Top + page.Height - (float)footerDistance * (1f / 15f) * _zoom;
			bool hasHeader = !string.IsNullOrEmpty(headerText)
				|| (headerImages != null && headerImages.Count > 0);
			bool hasFooter = !string.IsNullOrEmpty(footerText)
				|| (footerImages != null && footerImages.Count > 0);

			if (design && hasHeader)
			{
				float bandBottom = Math.Min(page.BodyTop - 2f, headerTop + 30f);
				using (Brush band = new SolidBrush(Color.FromArgb(247, 250, 253)))
				using (Pen line = new Pen(Color.FromArgb(185, 198, 214), 1f))
				{
					line.DashStyle = DashStyle.Dash;
					g.FillRectangle(band, bodyLeft - 3f, headerTop - 4f, bodyWidth + 6f,
						Math.Max(18f, bandBottom - headerTop + 5f));
					g.DrawLine(line, bodyLeft - 3f, bandBottom, bodyLeft + bodyWidth + 3f, bandBottom);
				}
			}
			if (design && hasFooter)
			{
				float bandTop = Math.Max(page.BodyBottom + 2f, footerBottom - 24f);
				using (Brush band = new SolidBrush(Color.FromArgb(247, 250, 253)))
				using (Pen line = new Pen(Color.FromArgb(185, 198, 214), 1f))
				{
					line.DashStyle = DashStyle.Dash;
					g.FillRectangle(band, bodyLeft - 3f, bandTop, bodyWidth + 6f,
						Math.Max(18f, footerBottom - bandTop + 6f));
					g.DrawLine(line, bodyLeft - 3f, bandTop, bodyLeft + bodyWidth + 3f, bandTop);
				}
			}

			if (section != null)
			{
				float imageX = bodyLeft;
				for (int i = 0; i < headerImages.Count; i++)
				{
					ImageInline imageInline = headerImages[i];
					float imageWidth = Math.Max(12f, (float)imageInline.WidthTwips * (1f / 15f) * _zoom);
					float imageHeight = Math.Max(8f, (float)imageInline.HeightTwips * (1f / 15f) * _zoom);
					try
					{
						using (Image image = ImageUtil.LoadImage(imageInline.Data, imageInline.Extension,
							imageInline.WidthTwips, imageInline.HeightTwips))
						{
							g.DrawImage(image, imageX, headerTop, imageWidth, imageHeight);
						}
					}
					catch
					{
					}
					imageX += imageWidth + 6f;
				}
				imageX = bodyLeft;
				for (int i = 0; i < footerImages.Count; i++)
				{
					ImageInline imageInline = footerImages[i];
					float imageWidth = Math.Max(12f, (float)imageInline.WidthTwips * (1f / 15f) * _zoom);
					float imageHeight = Math.Max(8f, (float)imageInline.HeightTwips * (1f / 15f) * _zoom);
					try
					{
						using (Image image = ImageUtil.LoadImage(imageInline.Data, imageInline.Extension,
							imageInline.WidthTwips, imageInline.HeightTwips))
						{
							g.DrawImage(image, imageX, footerBottom - imageHeight, imageWidth, imageHeight);
						}
					}
					catch
					{
					}
					imageX += imageWidth + 6f;
				}
			}

			Font headerFont = GetFont(headerFormat);
			float headerLineHeight = Math.Max(8f, headerFont.GetHeight(g));
			if (!string.IsNullOrEmpty(headerText))
			{
				string[] headerLines = Normalize(headerText).Split('\n');
				Color color = headerFormat.ForeColor ?? (design ? Color.FromArgb(70, 82, 98) : Color.Black);
				using (Brush brush = new SolidBrush(color))
				using (StringFormat format = HeaderStringFormat(headerAlignment))
				{
					float y = headerTop;
					for (int i = 0; i < headerLines.Length; i++)
					{
						if (i > 0 && y + headerLineHeight > page.BodyTop - 2f) break;
						g.DrawString(Expand(headerLines[i], page.Index + 1), headerFont, brush,
							new RectangleF(bodyLeft, y, bodyWidth, headerLineHeight + 2f), format);
						y += headerLineHeight;
					}
				}
			}

			Font footerFont = GetFont(footerFormat);
			float footerLineHeight = Math.Max(8f, footerFont.GetHeight(g));
			if (!string.IsNullOrEmpty(footerText))
			{
				string[] footerLines = Normalize(footerText).Split('\n');
				Color color = footerFormat.ForeColor ?? (design ? Color.FromArgb(70, 82, 98) : Color.Black);
				float top = footerBottom - footerLines.Length * footerLineHeight;
				using (Brush brush = new SolidBrush(color))
				using (StringFormat format = HeaderStringFormat(footerAlignment))
				{
					for (int i = 0; i < footerLines.Length; i++)
					{
						g.DrawString(Expand(footerLines[i], page.Index + 1), footerFont, brush,
							new RectangleF(bodyLeft, top + i * footerLineHeight, bodyWidth, footerLineHeight + 2f), format);
					}
				}
			}
			else
			{
				string fallback = "Página " + (page.Index + 1) + " de " + _pages.Count;
				Font fallbackFont = GetFont(new RunFormat("Calibri", 18));
				SizeF size = g.MeasureString(fallback, fallbackFont);
				g.DrawString(fallback, fallbackFont, Brushes.Gray,
					page.Left + page.Width - (page.Left + page.Width - page.BodyRight) - size.Width,
					page.Top + page.Height - (float)_document.MarginBottomTwips * (1f / 15f) * _zoom + 4f);
			}
		}

		private static StringFormat HeaderStringFormat(ParagraphAlignment alignment)
		{
			StringFormat format = (StringFormat)StringFormat.GenericTypographic.Clone();
			format.Alignment = alignment == ParagraphAlignment.Center ? StringAlignment.Center
				: alignment == ParagraphAlignment.Right ? StringAlignment.Far : StringAlignment.Near;
			format.LineAlignment = StringAlignment.Near;
			format.Trimming = StringTrimming.None;
			return format;
		}

		private string Expand(string text, int pageNumber)
		{
			return (text ?? string.Empty).Replace("{PAGE}", pageNumber.ToString()).Replace("{NUMPAGES}", Math.Max(1, _pages.Count).ToString());
		}

		private static string Normalize(string text)
		{
			return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
		}

		private void DrawFloating(Graphics g, LayoutFloating floating)
		{
			FloatingObject obj = floating.Object;
			int rotation = obj == null ? 0 : ((obj.Rotation % 360) + 360) % 360;
			if (rotation == 0)
			{
				DrawFloatingCore(g, floating);
				return;
			}
			GraphicsState state = g.Save();
			try
			{
				float centerX = floating.X + floating.Width / 2f;
				float centerY = floating.Y + floating.Height / 2f;
				g.TranslateTransform(centerX, centerY);
				g.RotateTransform(rotation);
				g.TranslateTransform(-centerX, -centerY);
				DrawFloatingCore(g, floating);
			}
			finally
			{
				g.Restore(state);
			}
		}

		private void DrawFloatingCore(Graphics g, LayoutFloating floating)
		{
			FloatingObject @object = floating.Object;
			RectangleF rect = new RectangleF(floating.X, floating.Y, floating.Width, floating.Height);
			if (@object.Kind == FloatingKind.Image && @object.Image != null)
			{
				try
				{
					using (Image image = ImageUtil.LoadImage(@object.Image.Data, @object.Image.Extension, @object.Image.WidthTwips, @object.Image.HeightTwips))
					{
						g.DrawImage(image, rect);
						return;
					}
				}
				catch
				{
					return;
				}
			}
			if (@object.Kind == FloatingKind.Embedded)
			{
				if (@object.Image != null)
				{
					try
					{
						using (Image image2 = ImageUtil.LoadImage(@object.Image.Data, @object.Image.Extension, @object.Image.WidthTwips, @object.Image.HeightTwips))
						{
							g.DrawImage(image2, rect);
							return;
						}
					}
					catch
					{
						return;
					}
				}
				using (Brush brush = new SolidBrush(Color.FromArgb(250, 250, 250)))
				{
					using (Pen pen = new Pen(Color.FromArgb(150, 160, 170)))
					{
						g.FillRectangle(brush, rect);
						g.DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);
						return;
					}
				}
			}
			Color color = @object.Fill ?? ((@object.Kind == FloatingKind.TextBox) ? Color.FromArgb(250, 250, 250) : Color.FromArgb(235, 242, 250));
			using (Brush brush2 = new SolidBrush(color))
			{
				using (Pen pen2 = new Pen(@object.BorderColor ?? Color.FromArgb(90, 105, 125), Math.Max(1, @object.BorderSize)))
				{
					if (@object.Kind == FloatingKind.Shape)
					{
						using (GraphicsPath path = CreateShapePath(@object.ShapeKind, rect))
						{
							g.FillPath(brush2, path);
							g.DrawPath(pen2, path);
						}
					}
					else
					{
						g.FillRectangle(brush2, rect);
						g.DrawRectangle(pen2, rect.X, rect.Y, rect.Width, rect.Height);
					}
				}
			}
			if (!string.IsNullOrEmpty(@object.Text))
			{
				RunFormat textFormat = @object.Format ?? new RunFormat("Calibri", 20);
				Font font = GetFont(textFormat);
				using (Brush textBrush = new SolidBrush(textFormat.ForeColor ?? Color.Black))
				{
					g.DrawString(layoutRectangle: new RectangleF(rect.X + 6f, rect.Y + 4f, Math.Max(1f, rect.Width - 12f), Math.Max(1f, rect.Height - 8f)), s: @object.Text, font: font, brush: textBrush, format: MeasureFormat);
				}
			}
		}

		private static GraphicsPath CreateShapePath(FloatingShapeKind kind, RectangleF rect)
		{
			return ShapeGeometry.Create(kind, rect);
		}

		private void DrawLine(Graphics g, LayoutLine line)
		{
			if (!string.IsNullOrEmpty(line.Number))
			{
				RunFormat format = new RunFormat("Calibri", 22);
				g.DrawString(line.Number, GetFont(format), Brushes.Black, line.BulletX + 1f, line.Y + 1f, MeasureFormat);
			}
			else if (line.Bullet)
			{
				// The list marker is document furniture, not part of the run's
				// font size.  Use the default 11 pt height so increasing one
				// item to 30 pt does not turn its bullet into a huge dot.
				float baseHeight = GetFont(new RunFormat("Calibri", 22)).GetHeight(g);
				float num = Math.Max(4.5f, baseHeight * 0.26f);
				float num2 = line.BulletX + 5f;
				float num3 = line.Y + Math.Max(1f, (baseHeight - num) / 2f);
				using (Brush brush = new SolidBrush(Color.Black))
				{
					g.FillEllipse(brush, num2, num3, num, num);
				}
			}
			for (int i = 0; i < line.Segments.Count; i++)
			{
				LayoutSegment layoutSegment = line.Segments[i];
				if (layoutSegment.Image != null)
				{
					try
					{
						using (Image image = ImageUtil.LoadImage(layoutSegment.Image.Data, layoutSegment.Image.Extension, layoutSegment.Image.WidthTwips, layoutSegment.Image.HeightTwips))
						{
							g.DrawImage(image, layoutSegment.X, layoutSegment.Y, layoutSegment.Width, layoutSegment.Height);
						}
					}
					catch
					{
					}
					continue;
				}
				if (layoutSegment.Note != null)
				{
					RunFormat runFormat = new RunFormat("Calibri", 15);
					runFormat.ForeColor = Color.FromArgb(45, 95, 160);
					using (Brush brush2 = new SolidBrush(runFormat.ForeColor.Value))
					{
						g.DrawString(layoutSegment.Text, GetFont(runFormat), brush2, layoutSegment.X, layoutSegment.Y - 2f, MeasureFormat);
					}
					continue;
				}
				if (layoutSegment.Comment != null)
				{
					using (Brush brush3 = new SolidBrush(Color.FromArgb(255, 242, 160)))
					{
						g.FillRectangle(brush3, layoutSegment.X - 1f, layoutSegment.Y - 1f, layoutSegment.Width + 2f, layoutSegment.Height + 1f);
					}
				}
				if (string.IsNullOrEmpty(layoutSegment.Text))
				{
					continue;
				}
				bool isLink = !string.IsNullOrEmpty(layoutSegment.Format.LinkUrl);
				Color color = isLink ? Color.FromArgb(0, 82, 204)
					: (layoutSegment.Format.ForeColor ?? Color.Black);
				using (Brush brush5 = new SolidBrush(color))
				{
					RectangleF rect = new RectangleF(layoutSegment.X, layoutSegment.Y, layoutSegment.Width + 2f, layoutSegment.Height + 2f);
					if (layoutSegment.Format.BackColor.HasValue)
					{
						using (Brush brush4 = new SolidBrush(layoutSegment.Format.BackColor.Value))
						{
							g.FillRectangle(brush4, rect);
						}
					}
					g.DrawString(layoutSegment.Text, layoutSegment.Font, brush5, layoutSegment.X, layoutSegment.Y, MeasureFormat);
					if (layoutSegment.Format.Underline || isLink)
					{
						using (Pen pen = new Pen(color))
						{
							g.DrawLine(pen, layoutSegment.X, layoutSegment.Y + layoutSegment.Height + 1f, layoutSegment.X + layoutSegment.Width, layoutSegment.Y + layoutSegment.Height + 1f);
						}
					}
					if (layoutSegment.Format.Strikeout)
					{
						using (Pen pen2 = new Pen(color))
						{
							g.DrawLine(pen2, layoutSegment.X, layoutSegment.Y + layoutSegment.Height / 2f, layoutSegment.X + layoutSegment.Width, layoutSegment.Y + layoutSegment.Height / 2f);
						}
					}
				}
			}
			DrawSpellUnderlines(g, line);
		}

		private void DrawTable(Graphics g, LayoutTable table)
		{
			Font font = GetFont(new RunFormat("Calibri", 21));
			using (Pen pen = new Pen(Color.FromArgb(125, 132, 140)))
			{
				using (StringFormat stringFormat = new StringFormat())
				{
					stringFormat.Trimming = StringTrimming.EllipsisCharacter;
					stringFormat.FormatFlags = StringFormatFlags.LineLimit;
					for (int i = 0; i < table.Cells.Count; i++)
					{
						LayoutTableCell layoutTableCell = table.Cells[i];
						RectangleF rect = new RectangleF(layoutTableCell.X + 1f, layoutTableCell.Y + 1f, Math.Max(1f, layoutTableCell.Width - 2f), Math.Max(1f, layoutTableCell.Height - 2f));
						if (layoutTableCell.Cell != null && layoutTableCell.Cell.Shading.HasValue)
						{
							using (Brush brush = new SolidBrush(layoutTableCell.Cell.Shading.Value))
							{
								g.FillRectangle(brush, rect);
							}
						}
						Color color = ((layoutTableCell.Cell != null && layoutTableCell.Cell.BorderColor.HasValue) ? layoutTableCell.Cell.BorderColor.Value : pen.Color);
						int num = ((layoutTableCell.Cell == null || layoutTableCell.Cell.BorderSize <= 0) ? 1 : layoutTableCell.Cell.BorderSize);
						using (Pen pen2 = new Pen(color, num))
						{
							g.DrawRectangle(pen2, rect.X, rect.Y, rect.Width, rect.Height);
						}
						if (layoutTableCell.Cell != null && !string.IsNullOrEmpty(layoutTableCell.Cell.Text))
						{
							RectangleF layoutRectangle = new RectangleF(layoutTableCell.X + 4f, layoutTableCell.Y + 3f, Math.Max(1f, layoutTableCell.Width - 8f), Math.Max(1f, layoutTableCell.Height - 5f));
							RunFormat cellFormat = layoutTableCell.Cell.Format;
							if (cellFormat == null)
							{
								cellFormat = new RunFormat("Calibri", 21)
								{
									Bold = layoutTableCell.Cell.Bold,
									Italic = layoutTableCell.Cell.Italic
								};
							}
							Font font2 = GetFont(cellFormat);
							bool cellLink = !string.IsNullOrEmpty(layoutTableCell.Cell.LinkUrl);
							Color cellColor = cellLink ? Color.FromArgb(0, 82, 204)
								: (cellFormat.ForeColor ?? Color.Black);
							using (Brush cellText = new SolidBrush(cellColor))
							{
								g.DrawString(layoutTableCell.Cell.Text, font2, cellText, layoutRectangle, stringFormat);
								if (cellLink)
								{
									SizeF measured = g.MeasureString(layoutTableCell.Cell.Text, font2);
									float lineY = layoutRectangle.Y + Math.Min(layoutRectangle.Height - 2f, font2.GetHeight());
									using (Pen linkPen = new Pen(cellColor))
										g.DrawLine(linkPen, layoutRectangle.X, lineY,
											layoutRectangle.X + Math.Min(layoutRectangle.Width, measured.Width), lineY);
								}
							}
						}
					}
				}
			}
		}

		private void DrawSpellingHighlight(Graphics g, LayoutPage page)
		{
			if (_spellingParagraph < 0 || _spellingLength <= 0) return;
			int end = _spellingStart + _spellingLength;
			using (Brush fill = new SolidBrush(Color.FromArgb(70, 255, 190, 0)))
			using (Pen border = new Pen(Color.FromArgb(235, 220, 110, 0), 1.5f))
			{
				for (int i = 0; i < page.Lines.Count; i++)
				{
					LayoutLine line = page.Lines[i];
					if (line.Paragraph != _spellingParagraph) continue;
					for (int j = 0; j < line.Segments.Count; j++)
					{
						LayoutSegment segment = line.Segments[j];
						if (string.IsNullOrEmpty(segment.Text)) continue;
						int from = Math.Max(segment.Start, _spellingStart);
						int to = Math.Min(segment.End, end);
						if (to <= from) continue;
						float left = segment.X;
						if (to > from && segment.End > segment.Start)
						{
							float perChar = segment.Width / (segment.End - segment.Start);
							left += (from - segment.Start) * perChar;
							float width = Math.Max(2f, (to - from) * perChar);
							RectangleF rect = new RectangleF(left, segment.Y - 1f, width, segment.Height + 2f);
							g.FillRectangle(fill, rect);
							g.DrawRectangle(border, rect.X, rect.Y, rect.Width, rect.Height);
						}
					}
				}
			}
		}

		private void DrawSelection(Graphics g, LayoutPage page)
		{
			if (!HasSelection)
			{
				return;
			}
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			Color selectionColor = Focused
				? Color.FromArgb(120, 173, 214, 255)
				: Color.FromArgb(205, 35, 120, 215);
			using (Brush brush = new SolidBrush(selectionColor))
			{
				for (int i = 0; i < page.Lines.Count; i++)
				{
					LayoutLine layoutLine = page.Lines[i];
					if (layoutLine.Paragraph < start.P || layoutLine.Paragraph > end.P)
					{
						continue;
					}
					int num = ((layoutLine.Paragraph == start.P) ? Math.Max(layoutLine.Start, start.Offset) : layoutLine.Start);
					int num2 = ((layoutLine.Paragraph == end.P) ? Math.Min(layoutLine.End, end.Offset) : layoutLine.End);
					if (num2 <= num && (layoutLine.Start != layoutLine.End || num != num2))
					{
						continue;
					}
					for (int j = 0; j < layoutLine.Segments.Count; j++)
					{
						LayoutSegment layoutSegment = layoutLine.Segments[j];
						if (layoutSegment.Image != null)
						{
							if (num <= layoutSegment.Start && num2 >= layoutSegment.Start)
							{
								g.FillRectangle(brush, layoutSegment.X, layoutSegment.Y, layoutSegment.Width, layoutSegment.Height);
							}
						}
						else
						{
							if (string.IsNullOrEmpty(layoutSegment.Text))
							{
								continue;
							}
							int num3 = Math.Max(num, layoutSegment.Start);
							int num4 = Math.Min(num2, layoutSegment.End);
							if (num4 > num3)
							{
								float num5 = XForOffset(layoutSegment, num3);
								float num6 = XForOffset(layoutSegment, num4);
								if (num6 > num5)
								{
									g.FillRectangle(brush, num5, layoutSegment.Y, num6 - num5, layoutSegment.Height);
								}
							}
						}
					}
				}
			}
		}

		private LayoutSegment FindImageSegment(ImageInline image)
		{
			for (int i = 0; i < _lines.Count; i++)
			{
				for (int j = 0; j < _lines[i].Segments.Count; j++)
				{
					if (_lines[i].Segments[j].Image == image)
					{
						return _lines[i].Segments[j];
					}
				}
			}
			return null;
		}

		private LayoutSegment ImageSegmentAt(Point client)
		{
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			for (int i = 0; i < _lines.Count; i++)
			{
				for (int j = 0; j < _lines[i].Segments.Count; j++)
				{
					LayoutSegment layoutSegment = _lines[i].Segments[j];
					if (layoutSegment.Image != null && (float)num >= layoutSegment.X && (float)num <= layoutSegment.X + layoutSegment.Width && (float)num2 >= layoutSegment.Y && (float)num2 <= layoutSegment.Y + layoutSegment.Height)
					{
						return layoutSegment;
					}
				}
			}
			return null;
		}

		private RectangleF[] ImageHandleRects(LayoutSegment seg)
		{
			float num = 7f;
			float num2 = seg.X;
			float num3 = seg.Y;
			float num4 = seg.Width;
			float num5 = seg.Height;
			return new RectangleF[8]
			{
				new RectangleF(num2 - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 / 2f - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 + num5 / 2f - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 + num4 / 2f - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 - num / 2f, num3 + num5 / 2f - num / 2f, num, num)
			};
		}

		private int ImageHandleAt(Point client)
		{
			if (_selectedImage == null)
			{
				return -1;
			}
			LayoutSegment layoutSegment = FindImageSegment(_selectedImage);
			if (layoutSegment == null)
			{
				return -1;
			}
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			RectangleF[] array = ImageHandleRects(layoutSegment);
			for (int i = 0; i < array.Length; i++)
			{
				if (array[i].Contains(num, num2))
				{
					return i;
				}
			}
			return -1;
		}

		private void DrawImageHandles(Graphics g)
		{
			if (_selectedImage == null)
			{
				return;
			}
			LayoutSegment layoutSegment = FindImageSegment(_selectedImage);
			if (layoutSegment == null)
			{
				return;
			}
			RectangleF[] array = ImageHandleRects(layoutSegment);
			using (Pen pen = new Pen(Color.FromArgb(45, 105, 180), 1.5f))
			{
				using (Brush brush = new SolidBrush(Color.White))
				{
					g.DrawRectangle(pen, layoutSegment.X, layoutSegment.Y, layoutSegment.Width, layoutSegment.Height);
					for (int i = 0; i < array.Length; i++)
					{
						g.FillRectangle(brush, array[i]);
						g.DrawRectangle(pen, array[i].X, array[i].Y, array[i].Width, array[i].Height);
					}
				}
			}
		}

		private int PageAtClient(Point client)
		{
			EnsureLayout();
			float y = (float)(client.Y - base.AutoScrollPosition.Y);
			for (int i = 0; i < _pages.Count; i++)
			{
				if (y >= _pages[i].Top && y <= _pages[i].Top + _pages[i].Height)
					return i;
			}
			return Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage));
		}

		private void ReparentFloatingToPage(FloatingObject target, int pageIndex)
		{
			if (target == null || _pages.Count == 0) return;
			EnsureLayout();
			pageIndex = Math.Max(0, Math.Min(_pages.Count - 1, pageIndex));
			LayoutFloating layout = FindFloating(target);
			if (layout == null) return;
			int currentPage = -1;
			for (int i = 0; i < _pages.Count && currentPage < 0; i++)
			{
				for (int j = 0; j < _pages[i].Floatings.Count; j++)
				{
					if (object.ReferenceEquals(_pages[i].Floatings[j], layout))
					{
						currentPage = i;
						break;
					}
				}
			}
			if (currentPage == pageIndex) return;
			int targetParagraph = -1;
			for (int i = 0; i < _lines.Count; i++)
			{
				if (_lines[i].Page == pageIndex)
				{
					targetParagraph = _lines[i].Paragraph;
					break;
				}
			}
			if (targetParagraph < 0)
			{
				for (int i = 0; i < _tables.Count; i++)
				{
					if (_tables[i].Page == pageIndex)
					{
						targetParagraph = _tables[i].Paragraph;
						break;
					}
				}
			}
			if (targetParagraph < 0 || targetParagraph >= _document.Paragraphs.Count) return;
			for (int i = 0; i < _document.Paragraphs.Count; i++)
				_document.Paragraphs[i].FloatingObjects.Remove(target);
			_document.Paragraphs[targetParagraph].FloatingObjects.Add(target);
			LayoutPage page = _pages[pageIndex];
			target.OffsetXTwips = (int)Math.Round((layout.X - page.Left) / _zoom * 15f);
			target.OffsetYTwips = (int)Math.Round((layout.Y - page.Top) / _zoom * 15f);
			InvalidateLayout(true);
		}

		private LayoutFloating FindFloating(FloatingObject target)
		{
			for (int i = 0; i < _pages.Count; i++)
			{
				LayoutPage layoutPage = _pages[i];
				for (int j = 0; j < layoutPage.Floatings.Count; j++)
				{
					if (object.ReferenceEquals(layoutPage.Floatings[j].Object, target))
					{
						return layoutPage.Floatings[j];
					}
				}
			}
			return null;
		}

		private LayoutFloating FloatingAt(Point client)
		{
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			for (int num3 = _pages.Count - 1; num3 >= 0; num3--)
			{
				LayoutPage layoutPage = _pages[num3];
				for (int num4 = layoutPage.Floatings.Count - 1; num4 >= 0; num4--)
				{
					LayoutFloating layoutFloating = layoutPage.Floatings[num4];
					if ((float)num >= layoutFloating.X && (float)num <= layoutFloating.X + layoutFloating.Width && (float)num2 >= layoutFloating.Y && (float)num2 <= layoutFloating.Y + layoutFloating.Height)
					{
						return layoutFloating;
					}
				}
			}
			return null;
		}

		private RectangleF[] FloatingHandleRects(LayoutFloating floating)
		{
			float num = 8f;
			float num2 = floating.X;
			float num3 = floating.Y;
			float num4 = floating.Width;
			float num5 = floating.Height;
			return new RectangleF[8]
			{
				new RectangleF(num2 - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 / 2f - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 + num5 / 2f - num / 2f, num, num),
				new RectangleF(num2 + num4 - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 + num4 / 2f - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 - num / 2f, num3 + num5 - num / 2f, num, num),
				new RectangleF(num2 - num / 2f, num3 + num5 / 2f - num / 2f, num, num)
			};
		}

		private int FloatingHandleAt(Point client)
		{
			if (_selectedFloating == null)
			{
				return -1;
			}
			LayoutFloating layoutFloating = FindFloating(_selectedFloating);
			if (layoutFloating == null)
			{
				return -1;
			}
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			RectangleF[] array = FloatingHandleRects(layoutFloating);
			for (int i = 0; i < array.Length; i++)
			{
				if (array[i].Contains(num, num2))
				{
					return i;
				}
			}
			return -1;
		}

		private void DrawFloatingHandles(Graphics g)
		{
			if (_selectedFloating == null)
			{
				return;
			}
			LayoutFloating layoutFloating = FindFloating(_selectedFloating);
			if (layoutFloating == null)
			{
				return;
			}
			RectangleF[] array = FloatingHandleRects(layoutFloating);
			using (Pen pen = new Pen(Color.FromArgb(38, 95, 170), 1.5f))
			{
				using (Brush brush = new SolidBrush(Color.White))
				{
					g.DrawRectangle(pen, layoutFloating.X, layoutFloating.Y, layoutFloating.Width, layoutFloating.Height);
					for (int i = 0; i < array.Length; i++)
					{
						g.FillRectangle(brush, array[i]);
						g.DrawRectangle(pen, array[i].X, array[i].Y, array[i].Width, array[i].Height);
					}
				}
			}
		}

		private void DeleteSelectedFloating()
		{
			if (_selectedFloating != null)
			{
				PushUndo();
				for (int i = 0; i < _document.Paragraphs.Count && !_document.Paragraphs[i].FloatingObjects.Remove(_selectedFloating); i++)
				{
				}
				_selectedFloating = null;
				InvalidateLayout(true);
				OnTextChanged();
				Invalidate();
			}
		}

		private void BeginFloatingMove(Point client)
		{
			if (_selectedFloating != null)
			{
				_movingFloating = true;
				_floatingUndoPushed = false;
				Cursor = Cursors.SizeAll;
				_floatingDragStart = client;
				_floatingStartX = _selectedFloating.OffsetXTwips;
				_floatingStartY = _selectedFloating.OffsetYTwips;
				LayoutFloating layout = FindFloating(_selectedFloating);
				if (layout != null)
				{
					float pageX = client.X - base.AutoScrollPosition.X;
					float pageY = client.Y - base.AutoScrollPosition.Y;
					_floatingGrabX = pageX - layout.X;
					_floatingGrabY = pageY - layout.Y;
				}
				else
				{
					_floatingGrabX = 0f;
					_floatingGrabY = 0f;
				}
				base.Capture = true;
			}
		}

		private void BeginFloatingResize(int handle, Point client)
		{
			if (_selectedFloating != null)
			{
				LayoutFloating layoutFloating = FindFloating(_selectedFloating);
				if (layoutFloating != null)
				{
					PushUndo();
					_resizingFloating = true;
					_floatingHandle = handle;
					Cursor = ResizeCursorForHandle(handle);
					_floatingDragStart = client;
					_floatingStartX = _selectedFloating.OffsetXTwips;
					_floatingStartY = _selectedFloating.OffsetYTwips;
					_floatingStartWidth = Math.Max(1, (int)Math.Round(layoutFloating.Width / _zoom * 15f));
					_floatingStartHeight = Math.Max(1, (int)Math.Round(layoutFloating.Height / _zoom * 15f));
					base.Capture = true;
				}
			}
		}

		private void UpdateFloatingMove(Point client)
		{
			if (_movingFloating && _selectedFloating != null)
			{
				int num = client.X - _floatingDragStart.X;
				int num2 = client.Y - _floatingDragStart.Y;
				if ((Math.Abs(num) > 2 || Math.Abs(num2) > 2) && !_floatingUndoPushed)
				{
					PushUndo();
					_floatingUndoPushed = true;
				}
				EnsureLayout();
				if (_pages.Count == 0) return;
				int targetPage = PageAtClient(client);
				ReparentFloatingToPage(_selectedFloating, targetPage);
				EnsureLayout();
				LayoutPage page = _pages[Math.Max(0, Math.Min(_pages.Count - 1, targetPage))];
				float pageX = client.X - base.AutoScrollPosition.X;
				float pageY = client.Y - base.AutoScrollPosition.Y;
				_selectedFloating.OffsetXTwips = (int)Math.Round((pageX - _floatingGrabX - page.Left) / _zoom * 15f);
				_selectedFloating.OffsetYTwips = (int)Math.Round((pageY - _floatingGrabY - page.Top) / _zoom * 15f);
				InvalidateLayout(true);
			}
		}

		private void UpdateFloatingResize(Point client)
		{
			if (!_resizingFloating || _selectedFloating == null)
			{
				return;
			}
			float num = client.X - _floatingDragStart.X;
			float num2 = client.Y - _floatingDragStart.Y;
			float num3 = (float)_floatingStartWidth + (float)Math.Round(num / _zoom * 15f);
			float num4 = (float)_floatingStartHeight + (float)Math.Round(num2 / _zoom * 15f);
			bool flag = _floatingHandle == 0 || _floatingHandle == 6 || _floatingHandle == 7;
			bool flag2 = _floatingHandle == 0 || _floatingHandle == 1 || _floatingHandle == 2;
			bool flag3 = _floatingHandle == 2 || _floatingHandle == 3 || _floatingHandle == 4;
			bool flag4 = _floatingHandle == 4 || _floatingHandle == 5 || _floatingHandle == 6;
			if (flag)
			{
				num3 = (float)_floatingStartWidth - (float)Math.Round(num / _zoom * 15f);
			}
			if (flag2)
			{
				num4 = (float)_floatingStartHeight - (float)Math.Round(num2 / _zoom * 15f);
			}
			if (!flag3 && !flag)
			{
				num3 = _floatingStartWidth;
			}
			if (!flag4 && !flag2)
			{
				num4 = _floatingStartHeight;
			}
			if (_selectedFloating.Kind == FloatingKind.Image || _selectedFloating.Kind == FloatingKind.Embedded)
			{
				float num5 = ((_floatingStartHeight > 0) ? ((float)_floatingStartWidth / (float)_floatingStartHeight) : 1f);
				if (flag || flag3)
				{
					num4 = num3 / Math.Max(0.01f, num5);
				}
				else
				{
					num3 = num4 * num5;
				}
			}
			num3 = Math.Max(240f, num3);
			num4 = Math.Max(160f, num4);
			if (flag)
			{
				int num6 = _floatingStartX + _floatingStartWidth;
				_selectedFloating.OffsetXTwips = num6 - (int)Math.Round(num3);
			}
			if (flag2)
			{
				int num7 = _floatingStartY + _floatingStartHeight;
				_selectedFloating.OffsetYTwips = num7 - (int)Math.Round(num4);
			}
			_selectedFloating.WidthTwips = (int)Math.Round(num3);
			_selectedFloating.HeightTwips = (int)Math.Round(num4);
			InvalidateLayout(true);
		}

		private void BeginImageResize(int handle, Point client)
		{
			LayoutSegment layoutSegment = FindImageSegment(_selectedImage);
			if (layoutSegment != null)
			{
				PushUndo();
				_resizingImage = true;
				_resizeHandle = handle;
				Cursor = ResizeCursorForHandle(handle);
				_resizeStart = client;
				_resizeStartWidth = layoutSegment.Width;
				_resizeStartHeight = layoutSegment.Height;
				base.Capture = true;
			}
		}

		private void UpdateImageResize(Point client)
		{
			if (_resizingImage && _selectedImage != null)
			{
				float num = client.X - _resizeStart.X;
				float num2 = client.Y - _resizeStart.Y;
				float num3 = _resizeStartWidth;
				float num4 = _resizeStartHeight;
				if (_resizeHandle == 0 || _resizeHandle == 6 || _resizeHandle == 7)
				{
					num3 -= num;
				}
				if (_resizeHandle == 2 || _resizeHandle == 3 || _resizeHandle == 4)
				{
					num3 += num;
				}
				if (_resizeHandle == 0 || _resizeHandle == 1 || _resizeHandle == 2)
				{
					num4 -= num2;
				}
				if (_resizeHandle == 4 || _resizeHandle == 5 || _resizeHandle == 6)
				{
					num4 += num2;
				}
				num3 = Math.Max(24f, num3);
				num4 = Math.Max(18f, num4);
				bool flag = _resizeHandle == 3 || _resizeHandle == 7;
				bool flag2 = _resizeHandle == 1 || _resizeHandle == 5;
				float num5;
				if (flag && !flag2)
				{
					num5 = num3 / Math.Max(1f, _resizeStartWidth);
				}
				else if (flag2 && !flag)
				{
					num5 = num4 / Math.Max(1f, _resizeStartHeight);
				}
				else
				{
					float num6 = num3 / Math.Max(1f, _resizeStartWidth);
					float num7 = num4 / Math.Max(1f, _resizeStartHeight);
					num5 = ((Math.Abs(num6 - 1f) >= Math.Abs(num7 - 1f)) ? num6 : num7);
				}
				num3 = Math.Max(24f, _resizeStartWidth * num5);
				num4 = Math.Max(18f, _resizeStartHeight * num5);
				float num8 = (float)GetImageMaxWidthTwips() / 15f * _zoom;
				if (num3 > num8)
				{
					num4 *= num8 / num3;
					num3 = num8;
				}
				_selectedImage.WidthTwips = (int)Math.Round(num3 / _zoom * 15f);
				_selectedImage.HeightTwips = (int)Math.Round(num4 / _zoom * 15f);
				InvalidateLayout(true);
			}
		}

		private void DrawCaret(Graphics g, LayoutPage page)
		{
			if (Focused && _caretOn)
			{
				RectangleF rectangleF = CaretRect();
				if (!rectangleF.IsEmpty && !(rectangleF.Top < page.Top) && !(rectangleF.Top > page.Top + page.Height))
				{
					g.FillRectangle(Brushes.Black, rectangleF.X, rectangleF.Y, 1.5f, rectangleF.Height);
				}
			}
		}

		private RectangleF CaretRect()
		{
			EnsureHandleLayout();
			LayoutLine layoutLine = FindCaretLine();
			if (layoutLine == null)
			{
				return RectangleF.Empty;
			}
			return new RectangleF(XForOffset(layoutLine, _caret.Offset), layoutLine.Y, 1.5f, Math.Max(10f, layoutLine.Height));
		}

		private LayoutLine FindCaretLine()
		{
			LayoutLine layoutLine = null;
			for (int i = 0; i < _lines.Count; i++)
			{
				LayoutLine layoutLine2 = _lines[i];
				if (layoutLine2.Paragraph == _caret.P)
				{
					if (_caret.Offset < layoutLine2.Start)
					{
						break;
					}
					layoutLine = layoutLine2;
					if (_caret.Offset <= layoutLine2.End)
					{
						break;
					}
				}
			}
			if (layoutLine != null && _caret.Offset == layoutLine.End && layoutLine.End > layoutLine.Start)
			{
				for (int j = 0; j < _lines.Count; j++)
				{
					LayoutLine layoutLine3 = _lines[j];
					if (layoutLine3.Paragraph == layoutLine.Paragraph && layoutLine3.Start == _caret.Offset && layoutLine3 != layoutLine && layoutLine3.Y > layoutLine.Y)
					{
						return layoutLine3;
					}
				}
			}
			return layoutLine;
		}

		private float XForOffset(LayoutLine line, int offset)
		{
			if (line == null)
			{
				return 0f;
			}
			if (offset <= line.Start)
			{
				if (line.Segments.Count <= 0)
				{
					return line.X;
				}
				return line.Segments[0].X;
			}
			for (int i = 0; i < line.Segments.Count; i++)
			{
				LayoutSegment layoutSegment = line.Segments[i];
				if (layoutSegment.Text != null)
				{
					if (offset <= layoutSegment.Start)
					{
						return layoutSegment.X;
					}
					if (offset <= layoutSegment.End)
					{
						return XForOffset(layoutSegment, offset);
					}
				}
				else if (layoutSegment.Image != null && offset <= layoutSegment.Start)
				{
					return layoutSegment.X;
				}
			}
			return line.X + line.Width;
		}

		private float XForOffset(LayoutSegment seg, int offset)
		{
			if (seg.Text == null)
			{
				return seg.X;
			}
			if (offset <= seg.Start)
			{
				return seg.X;
			}
			if (offset >= seg.End)
			{
				return seg.X + seg.Width;
			}
			int length = Math.Max(0, Math.Min(seg.Text.Length, offset - seg.Start));
			return seg.X + Measure(seg.Text.Substring(0, length), seg.Font);
		}

		private void NormalizeSelection(out CaretPos start, out CaretPos end)
		{
			if (_caret.P < _anchor.P || (_caret.P == _anchor.P && _caret.Offset < _anchor.Offset))
			{
				start = _caret;
				end = _anchor;
			}
			else
			{
				start = _anchor;
				end = _caret;
			}
		}

		private CaretPos PointToPosition(Point client)
		{
			EnsureHandleLayout();
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			LayoutLine layoutLine = null;
			float num3 = float.MaxValue;
			for (int i = 0; i < _lines.Count; i++)
			{
				LayoutLine layoutLine2 = _lines[i];
				float num4 = 0f;
				if ((float)num2 < layoutLine2.Y)
				{
					num4 = layoutLine2.Y - (float)num2;
				}
				else if ((float)num2 > layoutLine2.Y + layoutLine2.Height)
				{
					num4 = (float)num2 - (layoutLine2.Y + layoutLine2.Height);
				}
				if (num4 < num3)
				{
					num3 = num4;
					layoutLine = layoutLine2;
				}
			}
			if (layoutLine == null)
			{
				return new CaretPos(0, 0);
			}
			int val = layoutLine.Start;
			if ((float)num > layoutLine.X)
			{
				val = layoutLine.End;
				for (int j = 0; j < layoutLine.Segments.Count; j++)
				{
					LayoutSegment layoutSegment = layoutLine.Segments[j];
					if ((float)num < layoutSegment.X)
					{
						val = layoutSegment.Start;
						break;
					}
					if ((float)num <= layoutSegment.X + layoutSegment.Width)
					{
						val = ((layoutSegment.Image == null) ? ((!string.IsNullOrEmpty(layoutSegment.Text)) ? ClosestChar(layoutSegment, num) : layoutSegment.End) : layoutSegment.Start);
						break;
					}
				}
			}
			Paragraph p = _document.Paragraphs[layoutLine.Paragraph];
			val = Math.Max(0, Math.Min(GetText(p).Length, val));
			return new CaretPos(layoutLine.Paragraph, val);
		}

		private int ClosestChar(LayoutSegment seg, float x)
		{
			if (seg == null || string.IsNullOrEmpty(seg.Text) || seg.Font == null)
				return 0;
			int result = seg.Start;

			// Layout disposes its measuring Graphics.  Hit testing runs after
			// layout, so measure with a short-lived Graphics instead of using
			// Measure(), which returns 0 when _layoutGraphics is null.
			Graphics graphics = _layoutGraphics;
			bool temporary = graphics == null;
			bool measured = false;
			try
			{
				if (graphics == null) graphics = CreateGraphics();
				float num = Math.Abs(x - seg.X);
				for (int i = 1; i <= seg.Text.Length; i++)
				{
					float prefix = graphics.MeasureString(
						seg.Text.Substring(0, i), seg.Font, int.MaxValue, MeasureFormat).Width;
					if (prefix > 0f) measured = true;
					float num2 = seg.X + prefix;
					float num3 = Math.Abs(x - num2);
					if (num3 < num)
					{
						num = num3;
						result = seg.Start + i;
					}
				}
			}
			catch
			{
				measured = false;
			}
			finally
			{
				if (temporary && graphics != null) graphics.Dispose();
			}
			if (!measured && seg.Width > 0f)
			{
				// If the measuring surface is unavailable or reports zero
				// widths, map the click proportionally instead of snapping to
				// the start or end of the word.
				float ratio = (x - seg.X) / seg.Width;
				ratio = Math.Max(0f, Math.Min(1f, ratio));
				result = seg.Start + (int)Math.Round(ratio * seg.Text.Length);
			}
			return result;
		}

		private void SetCaret(CaretPos pos, bool extend)
		{
			EnsureParagraphs();
			// Moving the document caret leaves table-cell and floating-object
			// contexts; otherwise a later format command could be applied to a
			// stale table or shape instead of the paragraph under the caret.
			_tableSelectionParagraph = -1;
			_selectedFloating = null;
			_selectedImage = null;
			if (pos.P < 0)
			{
				pos = new CaretPos(0, 0);
			}
			if (pos.P >= _document.Paragraphs.Count)
			{
				pos = new CaretPos(_document.Paragraphs.Count - 1, GetText(_document.Paragraphs[_document.Paragraphs.Count - 1]).Length);
			}
			int length = GetText(_document.Paragraphs[pos.P]).Length;
			if (pos.Offset < 0)
			{
				pos.Offset = 0;
			}
			if (pos.Offset > length)
			{
				pos.Offset = length;
			}
			_caret = pos;
			if (!extend)
			{
				_anchor = new CaretPos(pos.P, pos.Offset);
			}
			if (!HasSelection)
			{
				_typingFormat = FormatAt(_caret);
			}
			_caretOn = true;
			_blinkTimer.Stop();
			_blinkTimer.Start();
			_currentPage = PageForPosition(pos);
			EnsureCaretVisible();
			Invalidate();
			OnSelectionChanged();
		}

		private void MoveCaret(CaretPos pos, bool extend)
		{
			SetCaret(pos, extend);
		}

		private int PageForPosition(CaretPos pos)
		{
			EnsureHandleLayout();
			for (int i = 0; i < _lines.Count; i++)
			{
				if (_lines[i].Paragraph == pos.P && pos.Offset >= _lines[i].Start && pos.Offset <= _lines[i].End)
				{
					return _lines[i].Page;
				}
			}
			for (int j = 0; j < _lines.Count; j++)
			{
				if (_lines[j].Paragraph == pos.P)
				{
					return _lines[j].Page;
				}
			}
			return 0;
		}

		private void EnsureCaretVisible()
		{
			if (!base.IsHandleCreated)
			{
				return;
			}
			RectangleF rectangleF = CaretRect();
			if (!rectangleF.IsEmpty)
			{
				int num = -base.AutoScrollPosition.Y;
				int num2 = base.ClientSize.Height;
				int num3 = (int)rectangleF.Top;
				int num4 = num;
				if (num3 < num + 24)
				{
					num4 = Math.Max(0, num3 - 24);
				}
				else if ((float)num3 + rectangleF.Height > (float)(num + num2 - 24))
				{
					num4 = Math.Max(0, (int)((float)num3 + rectangleF.Height - (float)num2 + 24f));
				}
				if (num4 != num)
				{
					base.AutoScrollPosition = new Point(-base.AutoScrollPosition.X, num4);
				}
			}
		}

		private void InvalidateCaret()
		{
			if (Focused)
			{
				RectangleF rectangleF = CaretRect();
				if (rectangleF.IsEmpty)
				{
					Invalidate();
				}
				else
				{
					Invalidate(new Rectangle((int)rectangleF.X - 2 - base.AutoScrollPosition.X, (int)rectangleF.Y - 2 - base.AutoScrollPosition.Y, 8, (int)rectangleF.Height + 4));
				}
			}
		}

		protected override void OnGotFocus(EventArgs e)
		{
			base.OnGotFocus(e);
			_caretOn = true;
			_blinkTimer.Start();
			Invalidate();
		}

		protected override void OnLostFocus(EventArgs e)
		{
			base.OnLostFocus(e);
			_blinkTimer.Stop();
			Invalidate();
		}

		protected override bool IsInputKey(Keys keyData)
		{
			Keys keys = keyData & Keys.KeyCode;
			if (keys == Keys.Left || keys == Keys.Right || keys == Keys.Up || keys == Keys.Down || keys == Keys.Home || keys == Keys.End || keys == Keys.Prior || keys == Keys.Next || keys == Keys.Tab || keys == Keys.Return || keys == Keys.Back || keys == Keys.Delete)
			{
				return true;
			}
			return base.IsInputKey(keyData);
		}

		private string LinkAtClient(Point client)
		{
			EnsureLayout();
			// Layout coordinates are document coordinates.  Convert the mouse
			// position to that space, otherwise links fail after scrolling.
			float x = client.X - base.AutoScrollPosition.X;
			float y = client.Y - base.AutoScrollPosition.Y;
			for (int i = 0; i < _lines.Count; i++)
			{
				LayoutLine line = _lines[i];
				for (int j = 0; j < line.Segments.Count; j++)
				{
					LayoutSegment segment = line.Segments[j];
					if (segment.Format == null || string.IsNullOrEmpty(segment.Format.LinkUrl)) continue;
					RectangleF rect = new RectangleF(segment.X, segment.Y, segment.Width, segment.Height);
					if (rect.Contains(x, y)) return segment.Format.LinkUrl;
				}
			}
			for (int i = 0; i < _tables.Count; i++)
			{
				for (int j = 0; j < _tables[i].Cells.Count; j++)
				{
					LayoutTableCell cell = _tables[i].Cells[j];
					if (cell.Cell == null || string.IsNullOrEmpty(cell.Cell.LinkUrl)) continue;
					if (x >= cell.X && x <= cell.X + cell.Width
						&& y >= cell.Y && y <= cell.Y + cell.Height)
						return cell.Cell.LinkUrl;
				}
			}
			return null;
		}

		private bool OpenLinkWithConfirmation(string url)
		{
			string target = (url ?? string.Empty).Trim();
			Uri uri;
			if (!Uri.TryCreate(target, UriKind.Absolute, out uri)
				|| (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps
					&& uri.Scheme != Uri.UriSchemeMailto))
			{
				MessageBox.Show(FindForm(),
					"El enlace no tiene un destino http, https o mailto válido:\n\n" + target,
					"Abrir hipervínculo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}
			DialogResult result = MessageBox.Show(FindForm(),
				"Destino del enlace:\n\n" + uri.AbsoluteUri
				+ "\n\n¿Deseas abrirlo con el programa predeterminado?\n\n"
				+ "Sí = abrir   ·   No = omitir",
				"Abrir hipervínculo", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2);
			if (result != DialogResult.Yes) return false;
			try
			{
				IntPtr shellResult = ShellExecute(FindForm() == null ? IntPtr.Zero : FindForm().Handle, "open", uri.AbsoluteUri, null, null, 1);
				if (shellResult.ToInt64() > 32L) return true;
				Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
				return true;
			}
			catch (Exception ex)
			{
				MessageBox.Show(FindForm(), "No se pudo abrir el enlace.\n\n" + ex.Message,
					"Abrir hipervínculo", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}
		protected override void OnMouseDown(MouseEventArgs e)
		{
			base.OnMouseDown(e);
			if (e.Button != MouseButtons.Left)
			{
				return;
			}
			Focus();
			string link = LinkAtClient(e.Location);
			if (!string.IsNullOrEmpty(link) && OpenLinkWithConfirmation(link))
				return;
			int num = FloatingHandleAt(e.Location);
			if (num >= 0)
			{
				BeginFloatingResize(num, e.Location);
				return;
			}
			LayoutFloating layoutFloating = FloatingAt(e.Location);
			if (layoutFloating != null)
			{
				_selectedFloating = layoutFloating.Object;
				_selectedImage = null;
				_tableSelectionParagraph = -1;
				BeginFloatingMove(e.Location);
				Invalidate();
				return;
			}
			_selectedFloating = null;
			int num2 = ImageHandleAt(e.Location);
			if (num2 >= 0)
			{
				BeginImageResize(num2, e.Location);
				return;
			}
			LayoutSegment layoutSegment = ImageSegmentAt(e.Location);
			if (layoutSegment != null)
			{
				_selectedImage = layoutSegment.Image;
				_tableSelectionParagraph = -1;
				Invalidate();
				return;
			}
			_selectedImage = null;
			LayoutTableCell layoutTableCell = TableCellAt(e.Location);
			if (layoutTableCell != null)
			{
				_tableClickScroll = base.AutoScrollPosition;
				CommitCellEditor(false);
				_tableSelectionParagraph = FindTableParagraph(layoutTableCell);
				_tableSelectionStartRow = (_tableSelectionEndRow = layoutTableCell.Row);
				_tableSelectionStartColumn = (_tableSelectionEndColumn = layoutTableCell.Column);
				_tableDragging = true;
				_tableMoved = false;
				base.Capture = true;
				Invalidate();
				return;
			}
			_tableSelectionParagraph = -1;
			CommitCellEditor(false);
			CaretPos pos = PointToPosition(e.Location);
			if ((Control.ModifierKeys & Keys.Shift) != 0)
			{
				SetCaret(pos, true);
			}
			else
			{
				SetCaret(pos, false);
				_selecting = true;
				base.Capture = true;
			}
			if (e.Clicks == 2)
			{
				SelectWordAt(pos);
			}
		}

		protected override void OnMouseDoubleClick(MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Left)
			{
				LayoutFloating layoutFloating = FloatingAt(e.Location);
				if (layoutFloating != null && layoutFloating.Object.Kind == FloatingKind.Shape)
				{
					_selectedFloating = layoutFloating.Object;
					Invalidate();
					EditSelectedShape();
					return;
				}
				if (layoutFloating != null && layoutFloating.Object.Kind == FloatingKind.TextBox)
				{
					_selectedFloating = layoutFloating.Object;
					EditSelectedTextBox();
					return;
				}
				LayoutTableCell layoutTableCell = TableCellAt(e.Location);
				if (layoutTableCell != null)
				{
					StartCellEditor(layoutTableCell, e.Location);
					return;
				}
			}
			base.OnMouseDoubleClick(e);
		}

		protected override void OnMouseMove(MouseEventArgs e)
		{
			base.OnMouseMove(e);
			if (_resizingFloating && e.Button == MouseButtons.Left)
			{
				UpdateFloatingResize(e.Location);
			}
			else if (_movingFloating && e.Button == MouseButtons.Left)
			{
				UpdateFloatingMove(e.Location);
			}
			else if (_resizingImage && e.Button == MouseButtons.Left)
			{
				UpdateImageResize(e.Location);
			}
			else if (_tableDragging && e.Button == MouseButtons.Left)
			{
				LayoutTableCell layoutTableCell = TableCellAt(e.Location);
				if (layoutTableCell != null)
				{
					if (layoutTableCell.Row != _tableSelectionEndRow || layoutTableCell.Column != _tableSelectionEndColumn)
					{
						_tableMoved = true;
					}
					_tableSelectionEndRow = layoutTableCell.Row;
					_tableSelectionEndColumn = layoutTableCell.Column;
					Invalidate();
				}
			}
			else if (_selecting && e.Button == MouseButtons.Left)
			{
				if (e.Y < 20)
				{
					base.AutoScrollPosition = new Point(-base.AutoScrollPosition.X, Math.Max(0, -base.AutoScrollPosition.Y - 24));
				}
				else if (e.Y > base.ClientSize.Height - 20)
				{
					base.AutoScrollPosition = new Point(-base.AutoScrollPosition.X, -base.AutoScrollPosition.Y + 24);
				}
				SetCaret(PointToPosition(e.Location), true);
			}
			if (e.Button == MouseButtons.None && !_movingFloating && !_resizingFloating
				&& !_resizingImage && !_selecting && !_tableDragging)
			{
				UpdateEditorCursor(e.Location);
			}
		}

		private static Cursor ResizeCursorForHandle(int handle)
		{
			switch (handle)
			{
				case 0:
				case 4: return Cursors.SizeNWSE;
				case 1:
				case 5: return Cursors.SizeNS;
				case 2:
				case 6: return Cursors.SizeNESW;
				case 3:
				case 7: return Cursors.SizeWE;
				default: return Cursors.Default;
			}
		}

		private void UpdateEditorCursor(Point client)
		{
			int floatingHandle = FloatingHandleAt(client);
			if (floatingHandle >= 0)
			{
				Cursor = ResizeCursorForHandle(floatingHandle);
				return;
			}
			if (FloatingAt(client) != null)
			{
				Cursor = Cursors.SizeAll;
				return;
			}
			int imageHandle = ImageHandleAt(client);
			if (imageHandle >= 0)
			{
				Cursor = ResizeCursorForHandle(imageHandle);
				return;
			}
			if (ImageSegmentAt(client) != null)
			{
				Cursor = Cursors.SizeAll;
				return;
			}
			Cursor = string.IsNullOrEmpty(LinkAtClient(client)) ? Cursors.IBeam : Cursors.Hand;
		}

		protected override void OnMouseUp(MouseEventArgs e)
		{
			if (_resizingFloating || _movingFloating)
			{
				bool flag = _resizingFloating || _floatingUndoPushed;
				if (_movingFloating && _selectedFloating != null)
					ReparentFloatingToPage(_selectedFloating, PageAtClient(e.Location));
				_resizingFloating = false;
				_movingFloating = false;
				_floatingHandle = -1;
				_floatingUndoPushed = false;
				base.Capture = false;
				if (flag)
				{
					OnTextChanged();
				}
				Invalidate();
				return;
			}
			if (_resizingImage)
			{
				_resizingImage = false;
				_resizeHandle = -1;
				base.Capture = false;
				OnTextChanged();
				Invalidate();
				return;
			}
			if (_tableDragging)
			{
				_tableDragging = false;
				base.Capture = false;
				if (!_tableMoved)
				{
					LayoutTableCell layoutTableCell = TableCellAt(e.Location);
					if (layoutTableCell != null)
					{
						StartCellEditor(layoutTableCell, e.Location);
						if (_tableClickScroll.X != 0 || _tableClickScroll.Y != 0)
						{
							base.AutoScrollPosition = new Point(-_tableClickScroll.X, -_tableClickScroll.Y);
							PositionCellEditor();
						}
					}
				}
				Invalidate();
				return;
			}
			if (e.Button == MouseButtons.Right && _spellEnabled)
			{
				Focus();
				CaretPos caretPos = PointToPosition(e.Location);
				SpellError spellError = FindSpellError(caretPos.P, caretPos.Offset);
				if (spellError != null)
				{
					SelectRange(spellError.Paragraph, spellError.Start, spellError.Length);
					ShowSpellMenu(spellError, e.Location);
					return;
				}
			}
			base.OnMouseUp(e);
			_selecting = false;
			base.Capture = false;
		}

		protected override void OnMouseWheel(MouseEventArgs e)
		{
			if ((Control.ModifierKeys & Keys.Control) != 0)
			{
				Zoom = _zoom + ((e.Delta > 0) ? 0.1f : (-0.1f));
				HandledMouseEventArgs handledMouseEventArgs = e as HandledMouseEventArgs;
				if (handledMouseEventArgs != null)
				{
					handledMouseEventArgs.Handled = true;
				}
			}
			else
			{
				base.OnMouseWheel(e);
				PositionCellEditor();
				Invalidate();
			}
		}

		protected override void OnScroll(ScrollEventArgs se)
		{
			base.OnScroll(se);
			PositionCellEditor();
			// ScrollableControl can blit the previous frame when the viewport
			// moves.  Repaint the whole surface so fixed page furniture (the
			// ruler) and clipped images/tables never leave stale pixels behind.
			Invalidate();
		}

		private void SelectWordAt(CaretPos pos)
		{
			string text = GetText(_document.Paragraphs[pos.P]);
			int num = pos.Offset;
			while (num > 0 && IsWordChar(text[num - 1]))
			{
				num--;
			}
			int i;
			for (i = pos.Offset; i < text.Length && IsWordChar(text[i]); i++)
			{
			}
			_anchor = new CaretPos(pos.P, num);
			_caret = new CaretPos(pos.P, i);
			Invalidate();
			OnSelectionChanged();
		}

		private static bool IsWordChar(char c)
		{
			if (!char.IsLetterOrDigit(c))
			{
				return c == '_';
			}
			return true;
		}

		private static string NormalizeText(string text)
		{
			if (string.IsNullOrEmpty(text))
			{
				return string.Empty;
			}
			string text2 = text.Normalize(NormalizationForm.FormC);
			if (text2.Length > 0 && text2[0] == '\ufeff')
			{
				text2 = text2.Substring(1);
			}
			return text2;
		}

		private static void NormalizeDocumentText(DocumentModel document)
		{
			if (document == null)
			{
				return;
			}
			document.HeaderText = NormalizeText(document.HeaderText);
			document.FooterText = NormalizeText(document.FooterText);
			for (int i = 0; i < document.Paragraphs.Count; i++)
			{
				Paragraph paragraph = document.Paragraphs[i];
				for (int j = 0; j < paragraph.Inlines.Count; j++)
				{
					TextInline textInline = paragraph.Inlines[j] as TextInline;
					if (textInline != null)
					{
						textInline.Text = NormalizeText(textInline.Text);
					}
				}
				if (paragraph.TableRows != null)
				{
					for (int k = 0; k < paragraph.TableRows.Length; k++)
					{
						for (int l = 0; l < paragraph.TableRows[k].Length; l++)
						{
							paragraph.TableRows[k][l] = NormalizeText(paragraph.TableRows[k][l]);
						}
					}
				}
				if (paragraph.TableCells == null)
				{
					continue;
				}
				paragraph.EnsureTableCells();
				for (int m = 0; m < paragraph.TableCells.Length; m++)
				{
					for (int n = 0; n < paragraph.TableCells[m].Length; n++)
					{
						if (paragraph.TableCells[m][n] != null)
						{
							paragraph.TableCells[m][n].Text = NormalizeText(paragraph.TableCells[m][n].Text);
						}
					}
				}
			}
		}

		private static string GetText(Paragraph p)
		{
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline textInline = p.Inlines[i] as TextInline;
				if (textInline != null)
				{
					stringBuilder.Append(textInline.Text);
				}
			}
			return stringBuilder.ToString();
		}

		private RunFormat FormatAt(CaretPos pos)
		{
			Paragraph paragraph = _document.Paragraphs[pos.P];
			int num = 0;
			RunFormat runFormat = null;
			RunFormat runFormat2 = null;
			for (int i = 0; i < paragraph.Inlines.Count; i++)
			{
				TextInline textInline = paragraph.Inlines[i] as TextInline;
				if (textInline != null)
				{
					int length = textInline.Text.Length;
					if (pos.Offset > num && pos.Offset <= num + length)
					{
						return textInline.Format.Clone();
					}
					if (pos.Offset == num)
					{
						runFormat2 = textInline.Format;
					}
					num += length;
					if (pos.Offset == num)
					{
						runFormat = textInline.Format;
					}
				}
			}
			if (runFormat != null)
			{
				return runFormat.Clone();
			}
			if (runFormat2 != null)
			{
				return runFormat2.Clone();
			}
			return _typingFormat.Clone();
		}

		private void PushUndo()
		{
			_undo.Push(CloneDocument(_document));
			_redo.Clear();
			if (_undo.Count > 100)
			{
				DocumentModel[] array = _undo.ToArray();
				_undo.Clear();
				for (int num = Math.Min(99, array.Length - 1); num >= 0; num--)
				{
					_undo.Push(array[num]);
				}
			}
		}

		private void AfterEdit()
		{
			EnsureParagraphs();
			int num = Math.Min(_caret.P, _document.Paragraphs.Count - 1);
			int length = GetText(_document.Paragraphs[num]).Length;
			_caret = new CaretPos(num, Math.Min(_caret.Offset, length));
			_anchor = new CaretPos(_caret.P, _caret.Offset);
			_typingFormat = FormatAt(_caret);
			InvalidateLayout(true);
			EnsureCaretVisible();
			OnTextChanged();
			OnSelectionChanged();
		}

		private void MarkSelectionDeleted(CaretPos start, CaretPos end)
		{
			Paragraph first = _document.Paragraphs[start.P];
			if (start.P == end.P)
			{
				DocumentRevisions.MarkRange(first, start.Offset, end.Offset);
				return;
			}
			DocumentRevisions.MarkRange(first, start.Offset, GetText(first).Length);
			for (int i = start.P + 1; i < end.P; i++)
			{
				Paragraph middle = _document.Paragraphs[i];
				DocumentRevisions.MarkRange(middle, 0, GetText(middle).Length);
			}
			Paragraph last = _document.Paragraphs[end.P];
			DocumentRevisions.MarkRange(last, 0, end.Offset);
		}
		private void DeleteSelectionInternal()
		{
			if (!HasSelection)
			{
				return;
			}
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			if (_document.TrackChanges)
			{
				MarkSelectionDeleted(start, end);
				_caret = start;
				_anchor = start;
				return;
			}
			if (start.P == end.P)
			{
				RemoveText(_document.Paragraphs[start.P], start.Offset, end.Offset - start.Offset);
			}
			else
			{
				Paragraph paragraph = _document.Paragraphs[start.P];
				Paragraph paragraph2 = _document.Paragraphs[end.P];
				RemoveText(paragraph, start.Offset, GetText(paragraph).Length - start.Offset);
				RemoveText(paragraph2, 0, end.Offset);
				paragraph.Inlines.AddRange(paragraph2.Inlines);
				for (int num = end.P; num > start.P; num--)
				{
					_document.Paragraphs.RemoveAt(num);
				}
			}
			_caret = new CaretPos(start.P, start.Offset);
			_anchor = new CaretPos(_caret.P, _caret.Offset);
		}

		private void RemoveText(Paragraph p, int start, int length)
		{
			if (length <= 0)
			{
				return;
			}
			if (_document.TrackChanges)
			{
				DocumentRevisions.MarkRange(p, start, start + length);
				return;
			}
			int num = 0;
			List<Inline> list = new List<Inline>();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline textInline = p.Inlines[i] as TextInline;
				if (textInline == null)
				{
					list.Add(p.Inlines[i]);
					continue;
				}
				int length2 = textInline.Text.Length;
				if (num + length2 <= start || num >= start + length)
				{
					list.Add(textInline);
				}
				else
				{
					int num2 = Math.Max(0, start - num);
					int num3 = Math.Min(length2, start + length - num);
					if (num2 > 0)
					{
						TextInline item = new TextInline(textInline.Text.Substring(0, num2), textInline.Format.Clone());
						list.Add(item);
					}
					if (num3 < length2)
					{
						TextInline item2 = new TextInline(textInline.Text.Substring(num3), textInline.Format.Clone());
						list.Add(item2);
					}
				}
				num += length2;
			}
			p.Inlines.Clear();
			p.Inlines.AddRange(list);
			MergeRuns(p);
		}

		private static void MergeRuns(Paragraph p)
		{
			List<Inline> list = new List<Inline>();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline textInline = p.Inlines[i] as TextInline;
				if (textInline != null)
				{
					if (textInline.Text.Length == 0)
					{
						continue;
					}
					if (list.Count > 0)
					{
						TextInline textInline2 = list[list.Count - 1] as TextInline;
						if (textInline2 != null && CanMergeRuns(textInline2, textInline))
						{
							textInline2.Text += textInline.Text;
							continue;
						}
					}
				}
				list.Add(p.Inlines[i]);
			}
			p.Inlines.Clear();
			p.Inlines.AddRange(list);
		}

		private void InsertTextAtCaret(string text)
		{
			if (string.IsNullOrEmpty(text))
			{
				if (HasSelection)
				{
					PushUndo();
					DeleteSelectionInternal();
					AfterEdit();
				}
				return;
			}
			PushUndo();
			if (HasSelection)
			{
				DeleteSelectionInternal();
			}
			string text2 = NormalizeText(text).Replace("\r\n", "\n").Replace('\r', '\n');
			for (int i = 0; i < text2.Length; i++)
			{
				char c = text2[i];
				if (c == '\n')
				{
					SplitParagraphAt(_caret.P, _caret.Offset);
					_caret = new CaretPos(_caret.P + 1, 0);
					_anchor = new CaretPos(_caret.P, 0);
				}
				else
				{
					InsertRaw(_caret.P, _caret.Offset, c.ToString());
					_caret.Offset++;
					_anchor = new CaretPos(_caret.P, _caret.Offset);
				}
			}
			AfterEdit();
		}

		private void InsertRaw(int paragraphIndex, int offset, string text)
		{
			InsertRawFormatted(paragraphIndex, offset, text, _typingFormat);
		}

		private static bool CanMergeRuns(TextInline a, TextInline b)
		{
			return a != null && b != null
				&& a.Revision == b.Revision
				&& a.Field == b.Field
				&& string.Equals(a.FieldInstruction, b.FieldInstruction, StringComparison.Ordinal)
				&& string.Equals(a.FieldBookmark, b.FieldBookmark, StringComparison.Ordinal)
				&& a.Format.SameAs(b.Format);
		}

		private TextInline NewTypedRun(string text, RunFormat format)
		{
			TextInline run = new TextInline(text, format == null ? new RunFormat() : format.Clone());
			if (_document != null && _document.TrackChanges) DocumentRevisions.MarkInsert(run);
			return run;
		}
		private void InsertRawFormatted(int paragraphIndex, int offset, string text, RunFormat sourceFormat)
		{
			RunFormat format = sourceFormat == null ? _typingFormat.Clone() : sourceFormat.Clone();
			Paragraph paragraph = _document.Paragraphs[paragraphIndex];
			int num = 0;
			for (int i = 0; i < paragraph.Inlines.Count; i++)
			{
				TextInline textInline = paragraph.Inlines[i] as TextInline;
				if (textInline == null)
				{
					if (offset <= num)
					{
						paragraph.Inlines.Insert(i, NewTypedRun(text, format));
						MergeRuns(paragraph);
						return;
					}
					continue;
				}
				int length = textInline.Text.Length;
				if (offset < num + length)
				{
					int num2 = Math.Max(0, offset - num);
					if (num2 == 0)
					{
						paragraph.Inlines.Insert(i, NewTypedRun(text, format));
					}
					else
					{
						paragraph.Inlines.Insert(i + 1, new TextInline(textInline.Text.Substring(num2), textInline.Format.Clone()));
						textInline.Text = textInline.Text.Substring(0, num2);
						paragraph.Inlines.Insert(i + 1, NewTypedRun(text, format));
					}
					MergeRuns(paragraph);
					return;
				}
				num += length;
			}
			paragraph.Inlines.Add(NewTypedRun(text, format));
			MergeRuns(paragraph);
		}

		private void SplitParagraphAt(int paragraphIndex, int offset)
		{
			Paragraph paragraph = _document.Paragraphs[paragraphIndex];
			List<Inline> list = new List<Inline>();
			List<Inline> list2 = new List<Inline>();
			int num = 0;
			for (int i = 0; i < paragraph.Inlines.Count; i++)
			{
				TextInline textInline = paragraph.Inlines[i] as TextInline;
				if (textInline == null)
				{
					list.Add(paragraph.Inlines[i]);
					continue;
				}
				int length = textInline.Text.Length;
				if (offset >= num + length)
				{
					list.Add(textInline);
				}
				else if (offset <= num)
				{
					list2.Add(textInline);
				}
				else
				{
					int num2 = offset - num;
					list.Add(new TextInline(textInline.Text.Substring(0, num2), textInline.Format.Clone()));
					list2.Add(new TextInline(textInline.Text.Substring(num2), textInline.Format.Clone()));
				}
				num += length;
			}
			paragraph.Inlines.Clear();
			paragraph.Inlines.AddRange(list);
			Paragraph paragraph2 = CloneParagraph(paragraph);
			paragraph2.Inlines.Clear();
			paragraph2.FloatingObjects.Clear();
			paragraph2.Inlines.AddRange(list2);
			paragraph2.PageBreakBefore = false;
			_document.Paragraphs.Insert(paragraphIndex + 1, paragraph2);
		}

		private void MergeParagraphWithNext(int paragraphIndex)
		{
			if (paragraphIndex >= 0 && paragraphIndex + 1 < _document.Paragraphs.Count)
			{
				Paragraph paragraph = _document.Paragraphs[paragraphIndex];
				Paragraph paragraph2 = _document.Paragraphs[paragraphIndex + 1];
				paragraph.Inlines.AddRange(paragraph2.Inlines);
				paragraph.FloatingObjects.AddRange(paragraph2.FloatingObjects);
				paragraph2.FloatingObjects.Clear();
				_document.Paragraphs.RemoveAt(paragraphIndex + 1);
				MergeRuns(paragraph);
			}
		}

		private CaretPos PrevPosition(CaretPos pos)
		{
			if (pos.Offset > 0)
			{
				return new CaretPos(pos.P, pos.Offset - 1);
			}
			if (pos.P > 0)
			{
				return new CaretPos(pos.P - 1, GetText(_document.Paragraphs[pos.P - 1]).Length);
			}
			return pos;
		}

		private CaretPos NextPosition(CaretPos pos)
		{
			int length = GetText(_document.Paragraphs[pos.P]).Length;
			if (pos.Offset < length)
			{
				return new CaretPos(pos.P, pos.Offset + 1);
			}
			if (pos.P + 1 < _document.Paragraphs.Count)
			{
				return new CaretPos(pos.P + 1, 0);
			}
			return pos;
		}

		private void Backspace()
		{
			if (HasSelection)
			{
				InsertTextAtCaret(string.Empty);
			}
			else if (_caret.Offset > 0)
			{
				PushUndo();
				RemoveText(_document.Paragraphs[_caret.P], _caret.Offset - 1, 1);
				_caret = new CaretPos(_caret.P, _caret.Offset - 1);
				_anchor = new CaretPos(_caret.P, _caret.Offset);
				AfterEdit();
			}
			else if (_caret.P > 0)
			{
				PushUndo();
				int length = GetText(_document.Paragraphs[_caret.P - 1]).Length;
				MergeParagraphWithNext(_caret.P - 1);
				_caret = new CaretPos(_caret.P - 1, length);
				_anchor = new CaretPos(_caret.P, _caret.Offset);
				AfterEdit();
			}
		}

		private void DeleteForward()
		{
			if (HasSelection)
			{
				InsertTextAtCaret(string.Empty);
				return;
			}
			int length = GetText(_document.Paragraphs[_caret.P]).Length;
			if (_caret.Offset < length)
			{
				PushUndo();
				RemoveText(_document.Paragraphs[_caret.P], _caret.Offset, 1);
				if (_document.TrackChanges)
					_caret = new CaretPos(_caret.P, Math.Min(length, _caret.Offset + 1));
				AfterEdit();
			}
			else if (_caret.P + 1 < _document.Paragraphs.Count)
			{
				PushUndo();
				MergeParagraphWithNext(_caret.P);
				AfterEdit();
			}
		}

		private void AddParagraphAfterCaret()
		{
			PushUndo();
			if (HasSelection)
			{
				DeleteSelectionInternal();
			}
			SplitParagraphAt(_caret.P, _caret.Offset);
			_caret = new CaretPos(_caret.P + 1, 0);
			_anchor = new CaretPos(_caret.P, 0);
			AfterEdit();
		}

		private void MoveHorizontal(int direction, bool extend)
		{
			if (HasSelection && !extend)
			{
				CaretPos start;
				CaretPos end;
				NormalizeSelection(out start, out end);
				SetCaret((direction < 0) ? start : end, false);
			}
			else
			{
				SetCaret((direction < 0) ? PrevPosition(_caret) : NextPosition(_caret), extend);
			}
		}

		private void MoveVertical(int direction, bool extend)
		{
			LayoutLine layoutLine = FindCaretLine();
			if (layoutLine != null)
			{
				if (_preferredX < 0)
				{
					_preferredX = (int)XForOffset(layoutLine, _caret.Offset);
				}
				int num = _lines.IndexOf(layoutLine);
				int num2 = num + direction;
				if (num2 >= 0 && num2 < _lines.Count)
				{
					LayoutLine layoutLine2 = _lines[num2];
					int offset = OffsetAtX(layoutLine2, _preferredX);
					SetCaret(new CaretPos(layoutLine2.Paragraph, offset), extend);
				}
			}
		}

		private int OffsetAtX(LayoutLine line, float x)
		{
			if (x <= line.X)
			{
				return line.Start;
			}
			for (int i = 0; i < line.Segments.Count; i++)
			{
				LayoutSegment layoutSegment = line.Segments[i];
				if (layoutSegment.Image != null)
				{
					if (x <= layoutSegment.X + layoutSegment.Width / 2f)
					{
						return layoutSegment.Start;
					}
				}
				else if (!string.IsNullOrEmpty(layoutSegment.Text))
				{
					if (x <= layoutSegment.X)
					{
						return layoutSegment.Start;
					}
					if (x <= layoutSegment.X + layoutSegment.Width)
					{
						return ClosestChar(layoutSegment, x);
					}
				}
			}
			return line.End;
		}

		private void MoveHome(bool extend)
		{
			LayoutLine layoutLine = FindCaretLine();
			SetCaret(new CaretPos(_caret.P, (layoutLine != null) ? layoutLine.Start : 0), extend);
		}

		private void MoveEnd(bool extend)
		{
			LayoutLine layoutLine = FindCaretLine();
			SetCaret(new CaretPos(_caret.P, (layoutLine == null) ? GetText(_document.Paragraphs[_caret.P]).Length : layoutLine.End), extend);
		}

		private void MoveWord(int direction, bool extend)
		{
			string text = GetText(_document.Paragraphs[_caret.P]);
			int i = _caret.Offset;
			if (direction < 0)
			{
				while (i > 0 && !IsWordChar(text[i - 1]))
				{
					i--;
				}
				while (i > 0 && IsWordChar(text[i - 1]))
				{
					i--;
				}
			}
			else
			{
				for (; i < text.Length && IsWordChar(text[i]); i++)
				{
				}
				for (; i < text.Length && !IsWordChar(text[i]); i++)
				{
				}
			}
			SetCaret(new CaretPos(_caret.P, i), extend);
		}

		protected override void OnKeyPress(KeyPressEventArgs e)
		{
			base.OnKeyPress(e);
			if (!char.IsControl(e.KeyChar))
			{
				_selectedFloating = null;
				InsertTextAtCaret(e.KeyChar.ToString());
				e.Handled = true;
			}
		}

		protected override void OnKeyDown(KeyEventArgs e)
		{
			base.OnKeyDown(e);
			bool shift = e.Shift;
			bool control = e.Control;
			_caretOn = true;
			_blinkTimer.Stop();
			_blinkTimer.Start();
			if (_selectedFloating != null && e.KeyCode == Keys.Escape)
			{
				_selectedFloating = null;
				Invalidate();
				e.Handled = true;
				return;
			}
			if (_selectedFloating != null && (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back))
			{
				DeleteSelectedFloating();
				e.Handled = true;
				return;
			}
			if (!IsEditingTableCell && _tableSelectionParagraph >= 0
				&& !shift && !control && !e.Alt && IsArrowKey(e.KeyCode))
			{
				int rowDelta = e.KeyCode == Keys.Up ? -1 : (e.KeyCode == Keys.Down ? 1 : 0);
				int columnDelta = e.KeyCode == Keys.Left ? -1 : (e.KeyCode == Keys.Right ? 1 : 0);
				if (MoveTableSelection(rowDelta, columnDelta))
				{
					e.Handled = true;
					return;
				}
				ExitTable(rowDelta < 0 || columnDelta < 0 ? -1 : 1);
				e.Handled = true;
				return;
			}
			switch (e.KeyCode)
			{
			case Keys.Left:
				if (control)
				{
					MoveWord(-1, shift);
				}
				else
				{
					MoveHorizontal(-1, shift);
				}
				_preferredX = -1;
				e.Handled = true;
				break;
			case Keys.Right:
				if (control)
				{
					MoveWord(1, shift);
				}
				else
				{
					MoveHorizontal(1, shift);
				}
				_preferredX = -1;
				e.Handled = true;
				break;
			case Keys.Up:
				MoveVertical(-1, shift);
				e.Handled = true;
				break;
			case Keys.Down:
				MoveVertical(1, shift);
				e.Handled = true;
				break;
			case Keys.Home:
				if (control)
				{
					SetCaret(new CaretPos(0, 0), shift);
				}
				else
				{
					MoveHome(shift);
				}
				e.Handled = true;
				break;
			case Keys.End:
				if (control)
				{
					SetCaret(new CaretPos(_document.Paragraphs.Count - 1, GetText(_document.Paragraphs[_document.Paragraphs.Count - 1]).Length), shift);
				}
				else
				{
					MoveEnd(shift);
				}
				e.Handled = true;
				break;
			case Keys.Back:
				Backspace();
				e.Handled = true;
				break;
			case Keys.Delete:
				DeleteForward();
				e.Handled = true;
				break;
			case Keys.Return:
				AddParagraphAfterCaret();
				e.Handled = true;
				break;
			case Keys.Tab:
				InsertTextAtCaret("\t");
				e.Handled = true;
				break;
			case Keys.Prior:
				MovePage(-1, shift);
				e.Handled = true;
				break;
			case Keys.Next:
				MovePage(1, shift);
				e.Handled = true;
				break;
			case Keys.A:
				if (control)
				{
					_anchor = new CaretPos(0, 0);
					_caret = new CaretPos(_document.Paragraphs.Count - 1, GetText(_document.Paragraphs[_document.Paragraphs.Count - 1]).Length);
					Invalidate();
					OnSelectionChanged();
					e.Handled = true;
				}
				break;
			case Keys.C:
				if (control)
				{
					Copy();
					e.Handled = true;
				}
				break;
			case Keys.X:
				if (control)
				{
					Cut();
					e.Handled = true;
				}
				break;
			case Keys.V:
				if (control)
				{
					if (shift) PastePlainText();
					else Paste();
					e.Handled = true;
				}
				break;
			case Keys.Z:
				if (control)
				{
					Undo();
					e.Handled = true;
				}
				break;
			case Keys.Y:
				if (control)
				{
					Redo();
					e.Handled = true;
				}
				break;
			case Keys.B:
				if (control)
				{
					ToggleBold();
					e.Handled = true;
				}
				break;
			case Keys.I:
				if (control)
				{
					ToggleItalic();
					e.Handled = true;
				}
				break;
			case Keys.U:
				if (control)
				{
					ToggleUnderline();
					e.Handled = true;
				}
				break;
			}
		}

		private void MovePage(int direction, bool extend)
		{
			int num = PageForPosition(_caret) + direction;
			if (num < 0)
			{
				num = 0;
			}
			if (num >= _pages.Count)
			{
				num = _pages.Count - 1;
			}
			if (num >= 0)
			{
				LayoutPage layoutPage = _pages[num];
				int offset = 0;
				if (layoutPage.Lines.Count > 0)
				{
					offset = layoutPage.Lines[0].Start;
				}
				SetCaret(new CaretPos((layoutPage.Lines.Count > 0) ? layoutPage.Lines[0].Paragraph : 0, offset), extend);
			}
		}

		public void RecheckSpelling()
		{
			ScheduleSpellCheck();
		}

		public string[] GetSpellParagraphs()
		{
			EnsureParagraphs();
			string[] array = new string[_document.Paragraphs.Count];
			for (int i = 0; i < array.Length; i++)
			{
				array[i] = GetText(_document.Paragraphs[i]);
			}
			return array;
		}

		public void ReplaceSpelling(int paragraph, int start, int length, string replacement)
		{
			if (paragraph >= 0 && paragraph < _document.Paragraphs.Count)
			{
				int length2 = GetText(_document.Paragraphs[paragraph]).Length;
				start = Math.Max(0, Math.Min(length2, start));
				length = Math.Max(0, Math.Min(length2 - start, length));
				PushUndo();
				_anchor = new CaretPos(paragraph, start);
				_caret = new CaretPos(paragraph, start + length);
				DeleteSelectionInternal();
				_caret = new CaretPos(paragraph, start);
				string text = NormalizeText(replacement);
				for (int i = 0; i < text.Length; i++)
				{
					InsertRaw(paragraph, start + i, text[i].ToString());
				}
				_caret = new CaretPos(paragraph, start + text.Length);
				_anchor = new CaretPos(_caret.P, _caret.Offset);
				AfterEdit();
			}
		}

		public void InsertTable(string[][] rows)
		{
			if (rows == null || rows.Length == 0)
			{
				return;
			}
			PushUndo();
			if (HasSelection)
			{
				DeleteSelectionInternal();
			}
			int p = _caret.P;
			Paragraph paragraph = _document.Paragraphs[p];
			if (paragraph.TableRows == null && GetText(paragraph).Length == 0)
			{
				paragraph.TableRows = CloneRows(rows);
				paragraph.TableCells = null;
				paragraph.Inlines.Clear();
				paragraph.TableWidthTwips = 9000;
				if (p + 1 >= _document.Paragraphs.Count)
				{
					_document.Paragraphs.Add(new Paragraph());
				}
				_caret = new CaretPos(p + 1, 0);
			}
			else
			{
				Paragraph paragraph2 = new Paragraph();
				paragraph2.TableRows = CloneRows(rows);
				paragraph2.TableCells = null;
				paragraph2.TableWidthTwips = 9000;
				_document.Paragraphs.Insert(p, paragraph2);
				_caret = new CaretPos(p + 1, 0);
			}
			_anchor = new CaretPos(_caret.P, 0);
			AfterEdit();
		}

		private int TableParagraphForEditing()
		{
			if (_tableSelectionParagraph >= 0 && _tableSelectionParagraph < _document.Paragraphs.Count
				&& _document.Paragraphs[_tableSelectionParagraph].TableRows != null)
				return _tableSelectionParagraph;
			int at = TableAtCaret;
			if (at >= 0) return at;
			if (_caret.P > 0 && _caret.P - 1 < _document.Paragraphs.Count
				&& _document.Paragraphs[_caret.P - 1].TableRows != null)
				return _caret.P - 1;
			return -1;
		}

		private static bool TableHasMergedCells(Paragraph p)
		{
			if (p == null || p.TableCells == null) return false;
			for (int r = 0; r < p.TableCells.Length; r++)
			{
				if (p.TableCells[r] == null) continue;
				for (int c = 0; c < p.TableCells[r].Length; c++)
				{
					TableCell cell = p.TableCells[r][c];
					if (cell != null && (cell.Covered || cell.RowSpan > 1 || cell.ColSpan > 1))
						return true;
				}
			}
			return false;
		}

		public bool AddTableColumn()
		{
			int index = TableParagraphForEditing();
			if (index < 0) return false;
			Paragraph p = _document.Paragraphs[index];
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length == 0) return false;
			PushUndo();
			for (int r = 0; r < p.TableCells.Length; r++)
			{
				TableCell[] old = p.TableCells[r] ?? new TableCell[0];
				TableCell[] expanded = new TableCell[old.Length + 1];
				for (int c = 0; c < old.Length; c++) expanded[c] = old[c];
				expanded[old.Length] = new TableCell();
				p.TableCells[r] = expanded;
			}
			p.SyncTableRows();
			_tableSelectionParagraph = -1;
			InvalidateLayout(true);
			OnTextChanged();
			return true;
		}

		public bool RemoveTableColumn()
		{
			int index = TableParagraphForEditing();
			if (index < 0) return false;
			Paragraph p = _document.Paragraphs[index];
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length == 0) return false;
			if (TableHasMergedCells(p)) return false;
			int columns = 0;
			for (int r = 0; r < p.TableCells.Length; r++)
				if (p.TableCells[r] != null && p.TableCells[r].Length > columns) columns = p.TableCells[r].Length;
			if (columns <= 1) return false;
			int removeAt = _tableSelectionParagraph >= 0
				? Math.Max(0, Math.Min(columns - 1, Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn)))
				: columns - 1;
			PushUndo();
			for (int r = 0; r < p.TableCells.Length; r++)
			{
				TableCell[] old = p.TableCells[r] ?? new TableCell[0];
				if (removeAt >= old.Length) continue;
				TableCell[] reduced = new TableCell[old.Length - 1];
				for (int c = 0, d = 0; c < old.Length; c++)
					if (c != removeAt) reduced[d++] = old[c];
				p.TableCells[r] = reduced;
			}
			p.SyncTableRows();
			_tableSelectionParagraph = -1;
			InvalidateLayout(true);
			OnTextChanged();
			return true;
		}

		public bool AddTableRow()
		{
			int index = TableParagraphForEditing();
			if (index < 0) return false;
			Paragraph p = _document.Paragraphs[index];
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length == 0) return false;
			int columns = 0;
			for (int r = 0; r < p.TableCells.Length; r++)
				if (p.TableCells[r] != null && p.TableCells[r].Length > columns) columns = p.TableCells[r].Length;
			if (columns == 0) return false;
			int insertAt = p.TableCells.Length;
			PushUndo();
			TableCell[][] rows = new TableCell[p.TableCells.Length + 1][];
			for (int r = 0, d = 0; r < rows.Length; r++)
			{
				if (r == insertAt)
				{
					TableCell[] newRow = new TableCell[columns];
					for (int c = 0; c < columns; c++) newRow[c] = new TableCell();
					rows[r] = newRow;
				}
				else
				{
					rows[r] = p.TableCells[d++];
				}
			}
			p.TableCells = rows;
			p.SyncTableRows();
			_tableSelectionParagraph = -1;
			InvalidateLayout(true);
			OnTextChanged();
			return true;
		}

		public bool RemoveTableRow()
		{
			int index = TableParagraphForEditing();
			if (index < 0) return false;
			Paragraph p = _document.Paragraphs[index];
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length <= 1) return false;
			if (TableHasMergedCells(p)) return false;
			int removeAt = _tableSelectionParagraph >= 0
				? Math.Max(0, Math.Min(p.TableCells.Length - 1,
					Math.Min(_tableSelectionStartRow, _tableSelectionEndRow)))
				: p.TableCells.Length - 1;
			PushUndo();
			TableCell[][] rows = new TableCell[p.TableCells.Length - 1][];
			for (int r = 0, d = 0; r < p.TableCells.Length; r++)
				if (r != removeAt) rows[d++] = p.TableCells[r];
			p.TableCells = rows;
			p.SyncTableRows();
			_tableSelectionParagraph = -1;
			InvalidateLayout(true);
			OnTextChanged();
			return true;
		}

		public void ReplaceTable(int paragraphIndex, string[][] rows)
		{
			if (paragraphIndex >= 0 && paragraphIndex < _document.Paragraphs.Count && rows != null)
			{
				PushUndo();
				_document.Paragraphs[paragraphIndex].TableRows = CloneRows(rows);
				_document.Paragraphs[paragraphIndex].TableCells = null;
				_document.Paragraphs[paragraphIndex].TableWidthTwips = 9000;
				InvalidateLayout(true);
				OnTextChanged();
			}
		}

		public void DeleteTable(int paragraphIndex)
		{
			if (paragraphIndex >= 0 && paragraphIndex < _document.Paragraphs.Count && _document.Paragraphs[paragraphIndex].TableRows != null)
			{
				PushUndo();
				_document.Paragraphs.RemoveAt(paragraphIndex);
				if (_document.Paragraphs.Count == 0)
				{
					_document.Paragraphs.Add(new Paragraph());
				}
				_caret = new CaretPos(Math.Min(paragraphIndex, _document.Paragraphs.Count - 1), 0);
				_anchor = new CaretPos(_caret.P, 0);
				AfterEdit();
			}
		}

		public string[][] GetTableRows(int paragraphIndex)
		{
			if (paragraphIndex < 0 || paragraphIndex >= _document.Paragraphs.Count)
			{
				return null;
			}
			string[][] tableRows = _document.Paragraphs[paragraphIndex].TableRows;
			if (tableRows != null)
			{
				return CloneRows(tableRows);
			}
			return null;
		}

		public int TableParagraphAt(Point client)
		{
			LayoutTableCell layoutTableCell = TableCellAt(client);
			if (layoutTableCell != null)
			{
				return FindTableParagraph(layoutTableCell);
			}
			return -1;
		}

		private int FindTableParagraph(LayoutTableCell cell)
		{
			for (int i = 0; i < _tables.Count; i++)
			{
				for (int j = 0; j < _tables[i].Cells.Count; j++)
				{
					if (_tables[i].Cells[j] == cell)
					{
						return _tables[i].Paragraph;
					}
				}
			}
			return -1;
		}

		private LayoutTableCell TableCellAt(Point client)
		{
			EnsureHandleLayout();
			int num = client.X - base.AutoScrollPosition.X;
			int num2 = client.Y - base.AutoScrollPosition.Y;
			for (int i = 0; i < _tables.Count; i++)
			{
				LayoutTable layoutTable = _tables[i];
				for (int j = 0; j < layoutTable.Cells.Count; j++)
				{
					LayoutTableCell layoutTableCell = layoutTable.Cells[j];
					if ((float)num >= layoutTableCell.X && (float)num <= layoutTableCell.X + layoutTableCell.Width && (float)num2 >= layoutTableCell.Y && (float)num2 <= layoutTableCell.Y + layoutTableCell.Height)
					{
						return layoutTableCell;
					}
				}
			}
			return null;
		}

		public void SelectTableCells(int paragraphIndex, int startRow, int startColumn, int endRow, int endColumn)
		{
			if (paragraphIndex < 0 || paragraphIndex >= _document.Paragraphs.Count) return;
			Paragraph p = _document.Paragraphs[paragraphIndex];
			if (p.TableRows == null) return;
			p.EnsureTableCells();
			if (p.TableCells == null || p.TableCells.Length == 0) return;
			startRow = Math.Max(0, Math.Min(p.TableCells.Length - 1, startRow));
			endRow = Math.Max(0, Math.Min(p.TableCells.Length - 1, endRow));
			int columns = p.TableCells[startRow] == null ? 0 : p.TableCells[startRow].Length;
			if (columns == 0) return;
			startColumn = Math.Max(0, Math.Min(columns - 1, startColumn));
			endColumn = Math.Max(0, Math.Min(columns - 1, endColumn));
			_tableSelectionParagraph = paragraphIndex;
			_tableSelectionStartRow = startRow;
			_tableSelectionStartColumn = startColumn;
			_tableSelectionEndRow = endRow;
			_tableSelectionEndColumn = endColumn;
			Invalidate();
		}

		public void MergeSelectedTableCells()
		{
			if (!CanMergeTableCells)
			{
				return;
			}
			int num = Math.Min(_tableSelectionStartRow, _tableSelectionEndRow);
			int num2 = Math.Max(_tableSelectionStartRow, _tableSelectionEndRow);
			int num3 = Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn);
			int num4 = Math.Max(_tableSelectionStartColumn, _tableSelectionEndColumn);
			Paragraph paragraph = _document.Paragraphs[_tableSelectionParagraph];
			paragraph.EnsureTableCells();
			TableCell tableCell = FindCellOwner(paragraph.TableCells, num, num3);
			if (tableCell == null)
			{
				return;
			}
			PushUndo();
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = num; i <= num2; i++)
			{
				if (paragraph.TableCells[i] == null)
				{
					continue;
				}
				for (int j = num3; j <= num4 && j < paragraph.TableCells[i].Length; j++)
				{
					TableCell tableCell2 = paragraph.TableCells[i][j];
					if (tableCell2 == null || tableCell2.Covered || tableCell2 == tableCell)
					{
						continue;
					}
					if (!string.IsNullOrEmpty(tableCell2.Text))
					{
						if (stringBuilder.Length > 0)
						{
							stringBuilder.Append(' ');
						}
						stringBuilder.Append(tableCell2.Text);
					}
					tableCell2.Text = string.Empty;
				}
			}
			tableCell.Text = (tableCell.Text ?? string.Empty).Trim();
			if (stringBuilder.Length > 0)
			{
				if (tableCell.Text.Length > 0)
				{
					tableCell.Text += " ";
				}
				tableCell.Text += stringBuilder.ToString();
			}
			tableCell.RowSpan = Math.Max(1, num2 - num + 1);
			tableCell.ColSpan = Math.Max(1, num4 - num3 + 1);
			for (int k = num; k <= num2; k++)
			{
				if (paragraph.TableCells[k] == null)
				{
					continue;
				}
				for (int l = num3; l <= num4 && l < paragraph.TableCells[k].Length; l++)
				{
					if (paragraph.TableCells[k][l] == null)
					{
						paragraph.TableCells[k][l] = new TableCell();
					}
					if (k != num || l != num3)
					{
						paragraph.TableCells[k][l].Covered = true;
						paragraph.TableCells[k][l].RowSpan = 1;
						paragraph.TableCells[k][l].ColSpan = 1;
					}
				}
			}
			paragraph.SyncTableRows();
			_tableSelectionParagraph = -1;
			InvalidateLayout(true);
			OnTextChanged();
		}

		public void SplitSelectedTableCell()
		{
			if (_tableSelectionParagraph < 0)
			{
				return;
			}
			Paragraph paragraph = _document.Paragraphs[_tableSelectionParagraph];
			paragraph.EnsureTableCells();
			int num = Math.Min(_tableSelectionStartRow, _tableSelectionEndRow);
			int num2 = Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn);
			TableCell tableCell = FindCellOwner(paragraph.TableCells, num, num2);
			if (tableCell == null || (tableCell.RowSpan <= 1 && tableCell.ColSpan <= 1))
			{
				return;
			}
			PushUndo();
			int rowSpan = tableCell.RowSpan;
			int colSpan = tableCell.ColSpan;
			tableCell.RowSpan = 1;
			tableCell.ColSpan = 1;
			for (int i = num; i < num + rowSpan && i < paragraph.TableCells.Length; i++)
			{
				if (paragraph.TableCells[i] == null)
				{
					continue;
				}
				for (int j = num2; j < num2 + colSpan && j < paragraph.TableCells[i].Length; j++)
				{
					if (paragraph.TableCells[i][j] == null)
					{
						paragraph.TableCells[i][j] = new TableCell();
					}
					paragraph.TableCells[i][j].Covered = false;
					paragraph.TableCells[i][j].RowSpan = 1;
					paragraph.TableCells[i][j].ColSpan = 1;
				}
			}
			paragraph.SyncTableRows();
			InvalidateLayout(true);
			OnTextChanged();
		}

		public void EditSelectedTableCell()
		{
			if (_tableSelectionParagraph < 0)
			{
				return;
			}
			LayoutTableCell layoutTableCell = null;
			for (int i = 0; i < _tables.Count; i++)
			{
				if (layoutTableCell != null)
				{
					break;
				}
				LayoutTable layoutTable = _tables[i];
				if (layoutTable.Paragraph != _tableSelectionParagraph)
				{
					continue;
				}
				for (int j = 0; j < layoutTable.Cells.Count; j++)
				{
					LayoutTableCell layoutTableCell2 = layoutTable.Cells[j];
					if (layoutTableCell2.Row == Math.Min(_tableSelectionStartRow, _tableSelectionEndRow) && layoutTableCell2.Column == Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn))
					{
						layoutTableCell = layoutTableCell2;
						break;
					}
				}
			}
			if (layoutTableCell != null)
			{
				StartCellEditor(layoutTableCell);
			}
		}

		private void StartCellEditor(LayoutTableCell cell)
		{
			StartCellEditor(cell, Point.Empty, false);
		}

		private void StartCellEditor(LayoutTableCell cell, Point click)
		{
			StartCellEditor(cell, click, true);
		}

		private void StartCellEditor(LayoutTableCell cell, Point click, bool positionCaret)
		{
			if (cell == null || cell.Cell == null)
			{
				return;
			}
			CommitCellEditor(false);
			if (_cellEditor == null)
			{
				_cellEditor = new TextBox();
				_cellEditor.Multiline = true;
				_cellEditor.AcceptsReturn = true;
				_cellEditor.WordWrap = true;
				_cellEditor.ScrollBars = ScrollBars.None;
				_cellEditor.BorderStyle = BorderStyle.None;
				_cellEditor.BackColor = Color.White;
				_cellEditor.ForeColor = Color.Black;
				_cellEditor.Font = new Font(FontResolver.Resolve("Calibri", "Español"), 11f);
				_cellEditor.TextChanged += delegate
				{
					if (_cellEdit != null && !_suppressCellLeave)
					{
						_cellEdit.Cell.Text = _cellEditor.Text;
						if (_cellEditParagraph >= 0 && _cellEditParagraph < _document.Paragraphs.Count)
						{
							_document.Paragraphs[_cellEditParagraph].SyncTableRows();
						}
						Invalidate();
					}
				};
				_cellEditor.KeyDown += CellEditorKeyDown;
				_cellEditor.Leave += delegate
				{
					if (!_suppressCellLeave && !_suppressCellCommit && !FocusIsToolbar())
					{
						CommitCellEditor(true);
					}
				};
				base.Controls.Add(_cellEditor);
			}
			_cellEdit = cell;
			_cellEditParagraph = FindTableParagraph(cell);
			_cellOriginalText = cell.Cell.Text ?? string.Empty;
			PushUndo();
			_suppressCellLeave = true;
			try
			{
				_cellEditor.Text = cell.Cell.Text ?? string.Empty;
				PositionCellEditor();
				_cellEditor.Visible = true;
				_cellEditor.BringToFront();
				_cellEditor.Focus();
				if (positionCaret) PositionCellEditorCaret(click);
				else _cellEditor.SelectionStart = _cellEditor.TextLength;
				UpdateCellEditorStyle();
			}
			finally
			{
				_suppressCellLeave = false;
			}
			_tableSelectionParagraph = _cellEditParagraph;
			_tableSelectionStartRow = (_tableSelectionEndRow = cell.Row);
			_tableSelectionStartColumn = (_tableSelectionEndColumn = cell.Column);
			Invalidate();
		}

		private void PositionCellEditorCaret(Point client)
		{
			if (_cellEditor == null) return;
			int x = client.X - _cellEditor.Left;
			int y = client.Y - _cellEditor.Top;
			int index = _cellEditor.GetCharIndexFromPosition(new Point(x, y));
			_cellEditor.SelectionStart = Math.Max(0, Math.Min(_cellEditor.TextLength, index));
			_cellEditor.SelectionLength = 0;
		}

		private bool FocusIsToolbar()
		{
			Form form = FindForm();
			Control active = form == null ? null : form.ActiveControl;
			for (Control c = active; c != null; c = c.Parent)
			{
				string type = c.GetType().FullName ?? string.Empty;
				if (c is ToolStrip || c is ToolStripDropDown || type.StartsWith("System.Windows.Forms.ToolStrip"))
					return true;
			}
			return active == null;
		}

		private void CellEditorKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				if (_cellEdit != null)
				{
					_cellEdit.Cell.Text = _cellOriginalText;
				}
				CommitCellEditor(true);
				e.Handled = true;
			}
			else if (e.KeyCode == Keys.Tab)
			{
				LayoutTableCell cellEdit = _cellEdit;
				if (cellEdit != null)
				{
					SelectNextCell(cellEdit);
				}
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
			else if (!e.Shift && !e.Control && !e.Alt && IsArrowKey(e.KeyCode)
				&& CellEditorAtBoundary(e.KeyCode))
			{
				int rowDelta = e.KeyCode == Keys.Up ? -1 : (e.KeyCode == Keys.Down ? 1 : 0);
				int columnDelta = e.KeyCode == Keys.Left ? -1 : (e.KeyCode == Keys.Right ? 1 : 0);
				if (!MoveCellFromEditor(rowDelta, columnDelta))
					ExitTable(rowDelta < 0 || columnDelta < 0 ? -1 : 1);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private static bool IsArrowKey(Keys key)
		{
			return key == Keys.Left || key == Keys.Right || key == Keys.Up || key == Keys.Down;
		}

		private bool CellEditorAtBoundary(Keys key)
		{
			if (_cellEditor == null || _cellEditor.SelectionLength != 0) return false;
			int start = _cellEditor.SelectionStart;
			if (key == Keys.Left) return start == 0;
			if (key == Keys.Right) return start >= _cellEditor.TextLength;
			int line = _cellEditor.GetLineFromCharIndex(start);
			if (key == Keys.Up) return line <= 0;
			return line >= Math.Max(0, _cellEditor.Lines.Length - 1);
		}

		private bool MoveCellFromEditor(int rowDelta, int columnDelta)
		{
			LayoutTableCell current = _cellEdit;
			if (current == null) return false;
			LayoutTable table = null;
			for (int i = 0; i < _tables.Count; i++)
			{
				if (_tables[i].Cells.Contains(current))
				{
					table = _tables[i];
					break;
				}
			}
			if (table == null) return false;
			int currentRowSpan = current.Cell == null ? 1 : Math.Max(1, current.Cell.RowSpan);
			int currentColSpan = current.Cell == null ? 1 : Math.Max(1, current.Cell.ColSpan);
			int columns = 0;
			for (int i = 0; i < table.Cells.Count; i++)
				if (table.Cells[i] != null) columns = Math.Max(columns, table.Cells[i].Column + 1);
			int targetRow = current.Row + (rowDelta < 0 ? -1 : (rowDelta > 0 ? currentRowSpan : 0));
			int targetColumn = current.Column + (columnDelta < 0 ? -1 : (columnDelta > 0 ? currentColSpan : 0));
			// Same row-to-row wrap as the non-editing table navigation.
			if (columnDelta > 0 && targetColumn >= columns)
			{
				targetRow += currentRowSpan;
				targetColumn = 0;
			}
			else if (columnDelta < 0 && targetColumn < 0)
			{
				targetRow -= 1;
				targetColumn = Math.Max(0, columns - 1);
			}
			LayoutTableCell target = null;
			for (int i = 0; i < table.Cells.Count; i++)
			{
				LayoutTableCell cell = table.Cells[i];
				int rowSpan = cell.Cell == null ? 1 : Math.Max(1, cell.Cell.RowSpan);
				int colSpan = cell.Cell == null ? 1 : Math.Max(1, cell.Cell.ColSpan);
				if (targetRow >= cell.Row && targetRow < cell.Row + rowSpan
					&& targetColumn >= cell.Column && targetColumn < cell.Column + colSpan)
				{
					target = cell;
					break;
				}
			}
			if (target == null || target == current) return false;
			// StartCellEditor commits the current cell without rebuilding the
			// layout, so the target reference stays valid for the next arrow key.
			StartCellEditor(target);
			return true;
		}

		private void ExitTable(int direction)
		{
			int paragraphIndex = IsEditingTableCell ? _cellEditParagraph : _tableSelectionParagraph;
			if (paragraphIndex < 0 || paragraphIndex >= _document.Paragraphs.Count) return;
			EnsureParagraphs();
			int target = direction < 0 ? paragraphIndex - 1 : paragraphIndex + 1;
			int offset;
			if (target < 0)
			{
				// A table at the very top has no paragraph above it.  Exit to
				// the paragraph below when one exists; never invent a blank
				// paragraph just to leave the table.
				target = paragraphIndex + 1;
				if (target >= _document.Paragraphs.Count) return;
				offset = 0;
			}
			else if (target >= _document.Paragraphs.Count)
			{
				target = paragraphIndex - 1;
				if (target < 0) return;
				offset = GetText(_document.Paragraphs[target]).Length;
			}
			else
			{
				offset = direction < 0 ? GetText(_document.Paragraphs[target]).Length : 0;
			}
			if (IsEditingTableCell) CommitCellEditor(true);
			_tableSelectionParagraph = -1;
			SetCaret(new CaretPos(target, offset), false);
			_anchor = _caret;
			Focus();
			Invalidate();
		}

		private bool MoveTableSelection(int rowDelta, int columnDelta)
		{
			if (_tableSelectionParagraph < 0 || _tableSelectionParagraph >= _document.Paragraphs.Count)
				return false;
			Paragraph table = _document.Paragraphs[_tableSelectionParagraph];
			if (table.TableRows == null) return false;
			table.EnsureTableCells();
			if (table.TableCells == null || table.TableCells.Length == 0) return false;
			int currentRow = Math.Min(_tableSelectionStartRow, _tableSelectionEndRow);
			int currentColumn = Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn);
			TableCell currentOwner = FindCellOwner(table.TableCells, currentRow, currentColumn);
			int rowSpan = currentOwner == null ? 1 : Math.Max(1, currentOwner.RowSpan);
			int colSpan = currentOwner == null ? 1 : Math.Max(1, currentOwner.ColSpan);
			int row = currentRow + (rowDelta < 0 ? -1 : (rowDelta > 0 ? rowSpan : 0));
			int column = currentColumn + (columnDelta < 0 ? -1 : (columnDelta > 0 ? colSpan : 0));
			int columns = 0;
			for (int r = 0; r < table.TableCells.Length; r++)
				if (table.TableCells[r] != null && table.TableCells[r].Length > columns)
					columns = table.TableCells[r].Length;
			// Right from the last cell of a row moves to the first cell of
			// the next row instead of leaving the table.  Left from the
			// first cell moves to the last cell of the previous row.
			if (columnDelta > 0 && column >= columns)
			{
				row += rowSpan;
				column = 0;
			}
			else if (columnDelta < 0 && column < 0)
			{
				row -= 1;
				column = Math.Max(0, columns - 1);
			}
			if (row < 0 || row >= table.TableCells.Length) return false;
			if (column < 0 || column >= columns) return false;
			TableCell owner = FindCellOwner(table.TableCells, row, column);
			if (owner == null) return false;
			for (int r = 0; r < table.TableCells.Length; r++)
			{
				if (table.TableCells[r] == null) continue;
				for (int c = 0; c < table.TableCells[r].Length; c++)
					if (object.ReferenceEquals(table.TableCells[r][c], owner))
					{
						SelectTableCells(_tableSelectionParagraph, r, c, r, c);
						return true;
					}
			}
			return false;
		}

		private void SelectNextCell(LayoutTableCell current)
		{
			for (int i = 0; i < _tables.Count; i++)
			{
				LayoutTable layoutTable = _tables[i];
				if (!layoutTable.Cells.Contains(current)) continue;
				for (int j = 0; j < layoutTable.Cells.Count; j++)
				{
					LayoutTableCell layoutTableCell = layoutTable.Cells[j];
					if (layoutTableCell.Cell != null && layoutTableCell.Cell.Covered) continue;
					if (layoutTableCell.Row == current.Row && layoutTableCell.Column > current.Column)
					{
						_tableSelectionParagraph = layoutTable.Paragraph;
						_tableSelectionStartRow = (_tableSelectionEndRow = layoutTableCell.Row);
						_tableSelectionStartColumn = (_tableSelectionEndColumn = layoutTableCell.Column);
						StartCellEditor(layoutTableCell);
						return;
					}
				}
				for (int k = 0; k < layoutTable.Cells.Count; k++)
				{
					LayoutTableCell layoutTableCell2 = layoutTable.Cells[k];
					if (layoutTableCell2.Row == current.Row + 1)
					{
						_tableSelectionParagraph = layoutTable.Paragraph;
						_tableSelectionStartRow = (_tableSelectionEndRow = layoutTableCell2.Row);
						_tableSelectionStartColumn = (_tableSelectionEndColumn = layoutTableCell2.Column);
						StartCellEditor(layoutTableCell2);
						return;
					}
				}
				return;
			}
		}

		private void PositionCellEditor()
		{
			if (_cellEditor != null && _cellEdit != null && _cellEditor.Visible)
			{
				Rectangle bounds = new Rectangle((int)Math.Round(_cellEdit.X + (float)base.AutoScrollPosition.X) + 1, (int)Math.Round(_cellEdit.Y + (float)base.AutoScrollPosition.Y) + 1, Math.Max(10, (int)Math.Round(_cellEdit.Width) - 2), Math.Max(10, (int)Math.Round(_cellEdit.Height) - 2));
				_cellEditor.Bounds = bounds;
			}
		}

		private void CommitCellEditor(bool rebuild)
		{
			if (_suppressCellCommit) return;
			if (_cellEditor == null || !_cellEditor.Visible)
			{
				return;
			}
			_suppressCellLeave = true;
			try
			{
				if (_cellEdit != null)
				{
					_cellEdit.Cell.Text = _cellEditor.Text;
					if (_cellEditParagraph >= 0 && _cellEditParagraph < _document.Paragraphs.Count)
					{
						_document.Paragraphs[_cellEditParagraph].SyncTableRows();
					}
				}
				_cellEditor.Visible = false;
				_cellEdit = null;
				_cellEditParagraph = -1;
				if (rebuild)
				{
					InvalidateLayout(true);
					OnTextChanged();
				}
				Focus();
			}
			finally
			{
				_suppressCellLeave = false;
			}
		}

		private void DrawTableSelection(Graphics g, LayoutPage page)
		{
			if (_tableSelectionParagraph < 0)
			{
				return;
			}
			int num = Math.Min(_tableSelectionStartRow, _tableSelectionEndRow);
			int num2 = Math.Max(_tableSelectionStartRow, _tableSelectionEndRow);
			int num3 = Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn);
			int num4 = Math.Max(_tableSelectionStartColumn, _tableSelectionEndColumn);
			using (Brush brush = new SolidBrush(Color.FromArgb(55, 65, 125, 205)))
			{
				using (Pen pen = new Pen(Color.FromArgb(180, 52, 105, 180), 1.5f))
				{
					for (int i = 0; i < page.Tables.Count; i++)
					{
						LayoutTable layoutTable = page.Tables[i];
						if (layoutTable.Paragraph != _tableSelectionParagraph)
						{
							continue;
						}
						for (int j = 0; j < layoutTable.Cells.Count; j++)
						{
							LayoutTableCell layoutTableCell = layoutTable.Cells[j];
							if (layoutTableCell.Row >= num && layoutTableCell.Row <= num2 && layoutTableCell.Column >= num3 && layoutTableCell.Column <= num4)
							{
								g.FillRectangle(brush, layoutTableCell.X + 1f, layoutTableCell.Y + 1f, layoutTableCell.Width - 2f, layoutTableCell.Height - 2f);
								g.DrawRectangle(pen, layoutTableCell.X + 1f, layoutTableCell.Y + 1f, layoutTableCell.Width - 3f, layoutTableCell.Height - 3f);
							}
						}
					}
				}
			}
		}

		private static TableCell FindCellOwner(TableCell[][] cells, int row, int column)
		{
			for (int i = 0; i < cells.Length; i++)
			{
				if (cells[i] == null)
				{
					continue;
				}
				for (int j = 0; j < cells[i].Length; j++)
				{
					TableCell tableCell = cells[i][j];
					if (tableCell != null && !tableCell.Covered && row >= i && row < i + Math.Max(1, tableCell.RowSpan) && column >= j && column < j + Math.Max(1, tableCell.ColSpan))
					{
						return tableCell;
					}
				}
			}
			return null;
		}

		private static string[][] CloneRows(string[][] rows)
		{
			string[][] array = new string[rows.Length][];
			for (int i = 0; i < rows.Length; i++)
			{
				array[i] = ((rows[i] == null) ? new string[0] : ((string[])rows[i].Clone()));
			}
			return array;
		}

		public int GetImageMaxWidthTwips()
		{
			if (_document == null)
			{
				return 0;
			}
			int num = Math.Max(1000, _document.PageWidthTwips);
			int val = Math.Max(1000, num - _document.MarginLeftTwips - _document.MarginRightTwips);
			int val2 = num * 2 / 3;
			return Math.Max(240, Math.Min(val, val2));
		}

		public void InsertImage(byte[] data, string extension, int pixelWidth, int pixelHeight)
		{
			InsertImage(data, extension, pixelWidth, pixelHeight, GetImageMaxWidthTwips());
		}

		public void InsertImage(byte[] data, string extension, int pixelWidth, int pixelHeight, int maxWidthTwips)
		{
			if (data != null && data.Length != 0 && !((DateTime.UtcNow - _lastImageInsertUtc).TotalMilliseconds < 200.0))
			{
				_lastImageInsertUtc = DateTime.UtcNow;
				if (pixelWidth < 1)
				{
					pixelWidth = 1;
				}
				if (pixelHeight < 1)
				{
					pixelHeight = 1;
				}
				int num = Math.Max(240, pixelWidth * 15);
				int num2 = Math.Max(180, pixelHeight * 15);
				int val = ((maxWidthTwips > 0) ? maxWidthTwips : GetImageMaxWidthTwips());
				int num3 = Math.Min(num, val);
				int heightTwips = (int)Math.Max(1L, (long)num2 * (long)num3 / Math.Max(1, num));
				PushUndo();
				if (HasSelection)
				{
					DeleteSelectionInternal();
				}
				int index = Math.Max(0, Math.Min(_caret.P, _document.Paragraphs.Count - 1));
				Paragraph paragraph = _document.Paragraphs[index];
				EnsureLayout();
				int page = Math.Max(0, Math.Min(_pages.Count - 1, PageForPosition(_caret)));
				Point offset = InsertionOffsetFor(page, index);
				ImageInline image = new ImageInline(data, extension, num3, heightTwips, _typingFormat.Clone());
				FloatingObject floatingObject = new FloatingObject();
				floatingObject.Kind = FloatingKind.Image;
				floatingObject.Image = image;
				floatingObject.WidthTwips = num3;
				floatingObject.HeightTwips = heightTwips;
				floatingObject.OffsetXTwips = offset.X;
				floatingObject.OffsetYTwips = offset.Y;
				floatingObject.Wrap = TextWrapMode.Square;
				floatingObject.BehindText = false;
				paragraph.FloatingObjects.Add(floatingObject);
				_selectedFloating = floatingObject;
				_selectedImage = null;
				Cursor = Cursors.SizeAll;
				AfterEdit();
				Invalidate();
			}
		}

		public bool ToggleSelectedImageWrap()
		{
			if (_selectedFloating != null && _selectedFloating.Kind == FloatingKind.Image)
			{
				if (_selectedFloating.Wrap == TextWrapMode.Square)
				{
					_selectedFloating.Wrap = TextWrapMode.TopAndBottom;
				}
				else if (_selectedFloating.Wrap == TextWrapMode.TopAndBottom)
				{
					_selectedFloating.Wrap = TextWrapMode.Behind;
				}
				else if (_selectedFloating.Wrap == TextWrapMode.Behind)
				{
					_selectedFloating.Wrap = TextWrapMode.InFront;
					_selectedFloating.BehindText = false;
				}
				else
				{
					_selectedFloating.Wrap = TextWrapMode.Square;
				}
				InvalidateLayout(true);
				return true;
			}
			if (_selectedImage == null)
			{
				return false;
			}
			Paragraph paragraph = null;
			int num = -1;
			for (int i = 0; i < _document.Paragraphs.Count; i++)
			{
				if (paragraph != null)
				{
					break;
				}
				Paragraph paragraph2 = _document.Paragraphs[i];
				for (int j = 0; j < paragraph2.Inlines.Count; j++)
				{
					if (object.ReferenceEquals(paragraph2.Inlines[j], _selectedImage))
					{
						paragraph = paragraph2;
						num = j;
						break;
					}
				}
			}
			if (paragraph == null || num < 0)
			{
				return false;
			}
			PushUndo();
			paragraph.Inlines.RemoveAt(num);
			FloatingObject floatingObject = new FloatingObject();
			floatingObject.Kind = FloatingKind.Image;
			floatingObject.Image = _selectedImage;
			floatingObject.WidthTwips = _selectedImage.WidthTwips;
			floatingObject.HeightTwips = _selectedImage.HeightTwips;
			floatingObject.Wrap = TextWrapMode.Square;
			floatingObject.OffsetXTwips = 0;
			floatingObject.OffsetYTwips = 0;
			paragraph.FloatingObjects.Add(floatingObject);
			_selectedImage = null;
			AfterEdit();
			return true;
		}

		private int FirstParagraphOnPageIndex(int pageIndex)
		{
			EnsureLayout();
			int best = -1;
			for (int i = 0; i < _lines.Count; i++)
			{
				if (_lines[i].Page != pageIndex || _lines[i].Paragraph < 0) continue;
				if (best < 0 || _lines[i].Paragraph < best) best = _lines[i].Paragraph;
			}
			for (int i = 0; i < _tables.Count; i++)
			{
				if (_tables[i].Page != pageIndex || _tables[i].Paragraph < 0) continue;
				if (best < 0 || _tables[i].Paragraph < best) best = _tables[i].Paragraph;
			}
			return best;
		}

		private void ReassignSectionIndexes()
		{
			_document.EnsureSections();
			for (int i = 0; i < _document.Paragraphs.Count; i++)
			{
				int index = 0;
				for (int s = 0; s < _document.Sections.Count; s++)
				{
					if (_document.Sections[s].StartParagraph <= i) index = s;
					else break;
				}
				_document.Paragraphs[i].SectionIndex = index;
			}
		}

		private int SplitSectionAt(int paragraphIndex, DocumentSection template)
		{
			_document.EnsureSections();
			if (paragraphIndex <= 0 || paragraphIndex >= _document.Paragraphs.Count) return -1;
			int sectionIndex = 0;
			for (int i = 0; i < _document.Sections.Count; i++)
			{
				if (_document.Sections[i].StartParagraph <= paragraphIndex) sectionIndex = i;
				else break;
			}
			if (_document.Sections[sectionIndex].StartParagraph == paragraphIndex) return sectionIndex;
			DocumentSection copy = template.Clone();
			copy.StartParagraph = paragraphIndex;
			copy.BreakType = SectionBreakType.NextPage;
			_document.Sections.Insert(sectionIndex + 1, copy);
			ReassignSectionIndexes();
			return sectionIndex + 1;
		}

		public void SetCurrentSectionColumns(int count, int spacingTwips, bool separator)
		{
			_document.EnsureSections();
			EnsureLayout();
			if (_pages.Count == 0) return;
			int page = Math.Max(0, Math.Min(_pages.Count - 1, PageForPosition(_caret)));
			int start = FirstParagraphOnPageIndex(page);
			if (start < 0 || start >= _document.Paragraphs.Count) return;

			int sectionIndex = 0;
			for (int i = 0; i < _document.Sections.Count; i++)
			{
				if (_document.Sections[i].StartParagraph <= start) sectionIndex = i;
				else break;
			}
			DocumentSection original = _document.Sections[sectionIndex].Clone();
			int sectionEnd = (sectionIndex + 1 < _document.Sections.Count
				? _document.Sections[sectionIndex + 1].StartParagraph - 1
				: _document.Paragraphs.Count - 1);
			int pageBreak = -1;
			for (int i = start + 1; i <= sectionEnd && i < _document.Paragraphs.Count; i++)
			{
				if (_document.Paragraphs[i].PageBreakBefore)
				{
					pageBreak = i;
					break;
				}
			}

			PushUndo();
			int activeIndex = SplitSectionAt(start, original);
			if (activeIndex < 0) activeIndex = sectionIndex;
			DocumentSection active = _document.Sections[activeIndex];
			active.ColumnCount = Math.Max(1, Math.Min(4, count));
			active.ColumnSpacingTwips = Math.Max(120, spacingTwips);
			active.ColumnsSeparator = separator;

			// A new section after the next explicit page break restores the
			// previous column layout.  The setting therefore starts on the
			// page where it was invoked and does not leak to the whole document.
			if (pageBreak > start)
			{
				int restoreIndex = SplitSectionAt(pageBreak, original);
				if (restoreIndex >= 0)
				{
					DocumentSection restore = _document.Sections[restoreIndex];
					restore.ColumnCount = original.ColumnCount;
					restore.ColumnSpacingTwips = original.ColumnSpacingTwips;
					restore.ColumnsSeparator = original.ColumnsSeparator;
				}
			}
			_caret = new CaretPos(Math.Min(_caret.P, _document.Paragraphs.Count - 1),
				Math.Min(_caret.Offset, GetText(_document.Paragraphs[Math.Min(_caret.P, _document.Paragraphs.Count - 1)]).Length));
			_anchor = _caret;
			InvalidateLayout(true);
			OnTextChanged();
		}

		public void InsertSectionBreak()
		{
			_document.EnsureSections();
			DocumentSection documentSection = _document.SectionForParagraph(_caret.P);
			int num = _document.Sections.IndexOf(documentSection);
			if (num >= 0)
			{
				PushUndo();
				int num2 = _caret.P + 1;
				DocumentSection documentSection2 = documentSection.Clone();
				documentSection2.StartParagraph = num2;
				documentSection2.BreakType = SectionBreakType.NextPage;
				_document.Sections.Insert(num + 1, documentSection2);
				for (int i = num2; i < _document.Paragraphs.Count; i++)
				{
					_document.Paragraphs[i].SectionIndex = num + 1;
				}
				if (num2 >= _document.Paragraphs.Count)
				{
					Paragraph paragraph = new Paragraph();
					paragraph.SectionIndex = num + 1;
					_document.Paragraphs.Add(paragraph);
				}
				_caret = new CaretPos(num2, 0);
				_anchor = new CaretPos(num2, 0);
				InvalidateLayout(true);
				OnTextChanged();
			}
		}

		private int ParagraphOnViewedPage()
		{
			EnsureLayout();
			// Base the insertion on what is really visible, not only on the
			// caret/page counter: the caret often remains on page 1 while the
			// user scrolls to another page.
			float viewTop = -base.AutoScrollPosition.Y;
			float viewBottom = viewTop + Math.Max(1, ClientSize.Height);
			float center = viewTop + Math.Max(1, ClientSize.Height) / 2f;
			int best = -1;
			float bestDistance = float.MaxValue;
			for (int i = 0; i < _lines.Count; i++)
			{
				LayoutLine line = _lines[i];
				if (line.Paragraph < 0) continue;
				float lineCenter = line.Y + line.Height / 2f;
				if (lineCenter < viewTop || lineCenter > viewBottom) continue;
				float distance = Math.Abs(lineCenter - center);
				if (distance < bestDistance)
				{
					bestDistance = distance;
					best = line.Paragraph;
				}
			}
			if (best < 0)
			{
				for (int i = 0; i < _tables.Count; i++)
				{
					LayoutTable table = _tables[i];
					if (table.Paragraph < 0) continue;
					float tableCenter = table.Y + table.Height / 2f;
					if (tableCenter < viewTop || tableCenter > viewBottom) continue;
					float distance = Math.Abs(tableCenter - center);
					if (distance < bestDistance)
					{
						bestDistance = distance;
						best = table.Paragraph;
					}
				}
			}
			if (best >= 0) return Math.Min(_document.Paragraphs.Count - 1, best);
			return Math.Max(0, Math.Min(_caret.P, _document.Paragraphs.Count - 1));
		}

		private int ParagraphForFloatingInsertion(int page)
		{
			EnsureLayout();
			// If the caret is on the page being viewed, insert at the caret.
			// Only fall back to the page center when the user has scrolled to a
			// different page without moving the caret.
			if (PageForPosition(_caret) == page)
				return Math.Max(0, Math.Min(_caret.P, _document.Paragraphs.Count - 1));
			return ParagraphOnViewedPage();
		}

		private Point InsertionOffsetFor(int page)
		{
			return InsertionOffsetFor(page, Math.Max(0, Math.Min(_caret.P, _document.Paragraphs.Count - 1)));
		}

		private Point InsertionOffsetFor(int page, int paragraphIndex)
		{
			int count = 0;
			for (int i = 0; i < _document.Paragraphs.Count; i++)
				for (int j = 0; j < _document.Paragraphs[i].FloatingObjects.Count; j++)
					count++;
			DocumentSection section = _document.SectionForParagraph(paragraphIndex);
			int x = section.MarginLeftTwips + 720 + (count % 4) * 360;
			int y = section.MarginTopTwips + 720 + ((count / 4) % 4) * 360;
			EnsureLayout();
			if (page >= 0 && page < _pages.Count)
			{
				LayoutPage layoutPage = _pages[page];
				for (int i = 0; i < _lines.Count; i++)
				{
					LayoutLine line = _lines[i];
					if (line.Page != page || line.Paragraph != paragraphIndex) continue;
					float localY = (line.Y - layoutPage.Top) / Math.Max(0.1f, _zoom) * 15f;
					y = (int)Math.Max(section.MarginTopTwips, localY);
					break;
				}
			}
			return new Point(x, y);
		}

		public void InsertTextBox()
		{
			if (!((DateTime.UtcNow - _lastTextBoxInsertUtc).TotalMilliseconds < 300.0))
			{
				_lastTextBoxInsertUtc = DateTime.UtcNow;
				PushUndo();
				EnsureLayout();
				int page = Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage));
				int index = ParagraphForFloatingInsertion(page);
				Paragraph paragraph = _document.Paragraphs[index];
				FloatingObject floatingObject = new FloatingObject();
				floatingObject.Kind = FloatingKind.TextBox;
				floatingObject.Fill = Color.FromArgb(250, 250, 250);
				floatingObject.Text = "Cuadro de texto";
				floatingObject.WidthTwips = 2880;
				floatingObject.HeightTwips = 1080;
				Point offset = InsertionOffsetFor(page, index);
				floatingObject.OffsetXTwips = offset.X;
				floatingObject.OffsetYTwips = offset.Y;
				floatingObject.Wrap = TextWrapMode.Square;
				floatingObject.Format = _typingFormat.Clone();
				paragraph.FloatingObjects.Add(floatingObject);
				_selectedFloating = floatingObject;
				_selectedImage = null;
				Cursor = Cursors.SizeAll;
				AfterEdit();
				Invalidate();
				Focus();
			}
		}

		public void InsertShape()
		{
			InsertShape(FloatingShapeKind.Rectangle);
		}

		public FloatingObject SelectedShape
		{
			get
			{
				return _selectedFloating != null && _selectedFloating.Kind == FloatingKind.Shape
					? _selectedFloating : null;
			}
		}

		public bool EditSelectedTextBox()
		{
			FloatingObject box = _selectedFloating;
			if (box == null || box.Kind != FloatingKind.TextBox) return false;
			using (FloatingTextDialog dialog = new FloatingTextDialog(box.Text,
				box.Format ?? _typingFormat.Clone(), box.Fill))
			{
				if (dialog.ShowDialog(FindForm()) != DialogResult.OK) return true;
				PushUndo();
				box.Text = dialog.Value;
				box.Format = dialog.Format;
				box.Fill = dialog.FillColor;
				_selectedFloating = box;
				InvalidateLayout(true);
				OnTextChanged();
				Invalidate();
				return true;
			}
		}

		public bool RotateSelectedFloating(int degrees)
		{
			if (_selectedFloating == null) return false;
			PushUndo();
			_selectedFloating.Rotation = ((_selectedFloating.Rotation + degrees) % 360 + 360) % 360;
			InvalidateLayout(true);
			OnTextChanged();
			Invalidate();
			return true;
		}

		public bool BringSelectedFloatingToFront()
		{
			return MoveSelectedFloating(true);
		}

		public bool SendSelectedFloatingToBack()
		{
			return MoveSelectedFloating(false);
		}

		private bool MoveSelectedFloating(bool toFront)
		{
			FloatingObject selected = _selectedFloating;
			if (selected == null) return false;
			for (int p = 0; p < _document.Paragraphs.Count; p++)
			{
				List<FloatingObject> list = _document.Paragraphs[p].FloatingObjects;
				int index = list.IndexOf(selected);
				if (index < 0) continue;
				if ((toFront && index == list.Count - 1) || (!toFront && index == 0)) return true;
				PushUndo();
				list.RemoveAt(index);
				if (toFront) list.Add(selected);
				else list.Insert(0, selected);
				InvalidateLayout(true);
				OnTextChanged();
				Invalidate();
				return true;
			}
			return false;
		}

		public bool SetSelectedShapeFormat(Color fill, Color border, int borderSize)
		{
			FloatingObject shape = SelectedShape;
			if (shape == null) return false;
			PushUndo();
			shape.Fill = fill;
			shape.BorderColor = border;
			shape.BorderSize = Math.Max(1, Math.Min(16, borderSize));
			InvalidateLayout(true);
			OnTextChanged();
			Invalidate();
			return true;
		}

		public bool EditSelectedShape()
		{
			FloatingObject shape = SelectedShape;
			if (shape == null) return false;
			using (ShapeFormatDialog dialog = new ShapeFormatDialog(shape.ShapeKind,
				shape.Fill ?? Color.FromArgb(221, 235, 247),
				shape.BorderColor ?? Color.FromArgb(68, 114, 196),
				shape.BorderSize))
			{
				if (dialog.ShowDialog(FindForm()) != DialogResult.OK) return true;
				SetSelectedShapeFormat(dialog.FillColor, dialog.BorderColor, dialog.BorderSize);
				return true;
			}
		}

		public void InsertShape(FloatingShapeKind kind)
		{
			if (!((DateTime.UtcNow - _lastShapeInsertUtc).TotalMilliseconds < 200.0))
			{
				_lastShapeInsertUtc = DateTime.UtcNow;
				PushUndo();
				EnsureLayout();
				int page = Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage));
				int index = ParagraphForFloatingInsertion(page);
				Paragraph paragraph = _document.Paragraphs[index];
				FloatingObject floatingObject = new FloatingObject();
				floatingObject.Kind = FloatingKind.Shape;
				floatingObject.ShapeKind = kind;
				floatingObject.Text = string.Empty;
				floatingObject.WidthTwips = 2520;
				floatingObject.HeightTwips = 1260;
				if (kind == FloatingShapeKind.Heart)
				{
					floatingObject.WidthTwips = 2000;
					floatingObject.HeightTwips = 1600;
				}
				Point offset = InsertionOffsetFor(page, index);
				floatingObject.OffsetXTwips = offset.X;
				floatingObject.OffsetYTwips = offset.Y;
				floatingObject.Fill = Color.FromArgb(221, 235, 247);
				floatingObject.BorderColor = Color.FromArgb(68, 114, 196);
				floatingObject.BorderSize = 2;
				floatingObject.Wrap = TextWrapMode.Square;
				paragraph.FloatingObjects.Add(floatingObject);
				_selectedFloating = floatingObject;
				_selectedImage = null;
				Cursor = Cursors.SizeAll;
				AfterEdit();
				Invalidate();
				Focus();
			}
		}

		public void InsertText(string text)
		{
			InsertTextAtCaret(text);
		}

		public void GoToPreviousPage()
		{
			MovePage(-1, false);
		}

		public void GoToNextPage()
		{
			MovePage(1, false);
		}

		public void PlaceCaret(int paragraph, int offset)
		{
			SetCaret(new CaretPos(paragraph, offset), false);
		}

		public void SelectAll()
		{
			_anchor = new CaretPos(0, 0);
			_caret = new CaretPos(_document.Paragraphs.Count - 1, GetText(_document.Paragraphs[_document.Paragraphs.Count - 1]).Length);
			Invalidate();
			OnSelectionChanged();
		}

		private List<FormattedRun> BuildSelectedRuns()
		{
			List<FormattedRun> runs = new List<FormattedRun>();
			if (!HasSelection) return runs;
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int p = start.P; p <= end.P && p < _document.Paragraphs.Count; p++)
			{
				Paragraph paragraph = _document.Paragraphs[p];
				int localStart = p == start.P ? start.Offset : 0;
				int localEnd = p == end.P ? end.Offset : GetText(paragraph).Length;
				int cursor = 0;
				RunFormat lastFormat = new RunFormat("Calibri", 22);
				for (int i = 0; i < paragraph.Inlines.Count; i++)
				{
					TextInline text = paragraph.Inlines[i] as TextInline;
					if (text == null) continue;
					int length = text.Text.Length;
					int from = Math.Max(localStart, cursor);
					int to = Math.Min(localEnd, cursor + length);
					if (to > from)
					{
						runs.Add(new FormattedRun(text.Text.Substring(from - cursor, to - from),
							text.Format == null ? lastFormat.Clone() : text.Format.Clone()));
						if (text.Format != null) lastFormat = text.Format;
					}
					cursor += length;
				}
				if (p < end.P)
					runs.Add(new FormattedRun("\n", lastFormat.Clone()));
			}
			return runs;
		}

		private static string EncodeBase64(string value)
		{
			return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
		}

		private static string DecodeBase64(string value)
		{
			try { return Encoding.UTF8.GetString(Convert.FromBase64String(value)); }
			catch { return string.Empty; }
		}

		private static string SerializeRuns(List<FormattedRun> runs)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < runs.Count; i++)
			{
				FormattedRun run = runs[i];
				RunFormat f = run.Format ?? new RunFormat("Calibri", 22);
				sb.Append(EncodeBase64(f.FontFamily)).Append('\t')
					.Append(f.SizeHalfPoints).Append('\t').Append(f.Bold ? 1 : 0).Append('\t')
					.Append(f.Italic ? 1 : 0).Append('\t').Append(f.Underline ? 1 : 0).Append('\t')
					.Append(f.Strikeout ? 1 : 0).Append('\t')
					.Append(f.ForeColor.HasValue ? f.ForeColor.Value.ToArgb() : 0).Append('\t')
					.Append(f.BackColor.HasValue ? f.BackColor.Value.ToArgb() : 0).Append('\t')
					.Append(f.Script).Append('\t').Append(EncodeBase64(f.LinkUrl)).Append('\t')
					.Append(EncodeBase64(run.Text)).Append('\n');
			}
			return sb.ToString();
		}

		private static List<FormattedRun> DeserializeRuns(string data)
		{
			List<FormattedRun> runs = new List<FormattedRun>();
			if (string.IsNullOrEmpty(data)) return runs;
			string[] lines = data.Split('\n');
			for (int i = 0; i < lines.Length; i++)
			{
				string[] fields = lines[i].Split('\t');
				if (fields.Length < 11) continue;
				try
				{
					RunFormat f = new RunFormat("Calibri", 22);
					f.FontFamily = DecodeBase64(fields[0]);
					f.SizeHalfPoints = int.Parse(fields[1], CultureInfo.InvariantCulture);
					f.Bold = fields[2] == "1";
					f.Italic = fields[3] == "1";
					f.Underline = fields[4] == "1";
					f.Strikeout = fields[5] == "1";
					int fore = int.Parse(fields[6], CultureInfo.InvariantCulture);
					int back = int.Parse(fields[7], CultureInfo.InvariantCulture);
					f.ForeColor = fore == 0 ? (Color?)null : Color.FromArgb(fore);
					f.BackColor = back == 0 ? (Color?)null : Color.FromArgb(back);
					f.Script = int.Parse(fields[8], CultureInfo.InvariantCulture);
					f.LinkUrl = DecodeBase64(fields[9]);
					runs.Add(new FormattedRun(DecodeBase64(fields[10]), f));
				}
				catch
				{
				}
			}
			return runs;
		}

		private static string HtmlEncode(string text)
		{
			return (text ?? string.Empty).Replace("&", "&amp;").Replace("<", "&lt;")
				.Replace(">", "&gt;").Replace("\"", "&quot;");
		}

		private static string BuildClipboardHtml(List<FormattedRun> runs)
		{
			StringBuilder fragment = new StringBuilder();
			for (int i = 0; i < runs.Count; i++)
			{
				FormattedRun run = runs[i];
				RunFormat f = run.Format ?? new RunFormat("Calibri", 22);
				string[] lines = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
				for (int line = 0; line < lines.Length; line++)
				{
					if (line > 0) fragment.Append("<br>");
					if (lines[line].Length == 0) continue;
					fragment.Append("<span style=\"font-family:").Append(HtmlEncode(f.FontFamily))
						.Append(";font-size:").Append((f.SizeHalfPoints / 2f).ToString("0.#", CultureInfo.InvariantCulture)).Append("pt");
					if (f.Bold) fragment.Append(";font-weight:bold");
					if (f.Italic) fragment.Append(";font-style:italic");
					if (f.Underline) fragment.Append(";text-decoration:underline");
					if (f.ForeColor.HasValue)
						fragment.Append(";color:#").Append(f.ForeColor.Value.R.ToString("X2"))
							.Append(f.ForeColor.Value.G.ToString("X2")).Append(f.ForeColor.Value.B.ToString("X2"));
					fragment.Append("\">").Append(HtmlEncode(lines[line])).Append("</span>");
				}
			}
			string body = "<html><body><!--StartFragment-->" + fragment + "<!--EndFragment--></body></html>";
			string header = "Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n";
			int startHtml = header.Length;
			int endHtml = startHtml + Encoding.UTF8.GetByteCount(body);
			int startFragment = startHtml + Encoding.UTF8.GetByteCount("<html><body><!--StartFragment-->");
			int endFragment = startHtml + Encoding.UTF8.GetByteCount("<html><body><!--StartFragment-->" + fragment);
			return string.Format(CultureInfo.InvariantCulture,
				"Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n{4}",
				startHtml, endHtml, startFragment, endFragment, body);
		}

		private string BuildClipboardRtf(List<FormattedRun> runs)
		{
			try
			{
				DocumentModel model = new DocumentModel();
				Paragraph paragraph = new Paragraph();
				model.Paragraphs.Add(paragraph);
				for (int i = 0; i < runs.Count; i++)
				{
					string text = runs[i].Text ?? string.Empty;
					string[] parts = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
					for (int part = 0; part < parts.Length; part++)
					{
						if (part > 0)
						{
							paragraph = new Paragraph();
							model.Paragraphs.Add(paragraph);
						}
						if (parts[part].Length > 0)
							paragraph.Inlines.Add(new TextInline(parts[part],
								runs[i].Format == null ? new RunFormat("Calibri", 22) : runs[i].Format.Clone()));
					}
				}
				return RtfWriter.Write(model, false);
			}
			catch
			{
				return null;
			}
		}

		private void InsertFormattedRuns(List<FormattedRun> runs)
		{
			if (runs == null || runs.Count == 0) return;
			PushUndo();
			if (HasSelection) DeleteSelectionInternal();
			for (int i = 0; i < runs.Count; i++)
			{
				FormattedRun run = runs[i];
				RunFormat format = run.Format ?? new RunFormat("Calibri", 22);
				string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
				int start = 0;
				for (int index = 0; index <= text.Length; index++)
				{
					if (index == text.Length || text[index] == '\n')
					{
						if (index > start)
						{
							string piece = text.Substring(start, index - start);
							InsertRawFormatted(_caret.P, _caret.Offset, piece, format);
							_caret.Offset += piece.Length;
							_anchor = new CaretPos(_caret.P, _caret.Offset);
						}
						if (index < text.Length)
						{
							SplitParagraphAt(_caret.P, _caret.Offset);
							_caret = new CaretPos(_caret.P + 1, 0);
							_anchor = new CaretPos(_caret.P, 0);
						}
						start = index + 1;
					}
				}
			}
			AfterEdit();
			OnSelectionChanged();
		}

		private static bool LooksLikeMojibake(string text)
		{
			for (int i = 0; i < text.Length; i++)
			{
				char c = text[i];
				char next = i + 1 < text.Length ? text[i + 1] : '\0';
				if ((c == '\u00C3' || c == '\u00C2') && next >= '\u0080' && next <= '\u00BF')
					return true;
				if (c == '\u00E2' && next == '\u20AC') return true;
			}
			return false;
		}

		private static int MojibakeScore(string text)
		{
			int score = 0;
			for (int i = 0; i < text.Length; i++)
			{
				if (text[i] == '\u00C3' || text[i] == '\u00C2' || text[i] == '\u00E2') score++;
				if (text[i] == '\uFFFD') score += 10;
			}
			return score;
		}

		private static string RepairMojibake(string text)
		{
			if (string.IsNullOrEmpty(text)) return text;
			try
			{
				// Some browsers place UTF-8 bytes on the clipboard and Windows
				// exposes them as Windows-1252.  A double decode inserts an
				// extra U+00C2 before the continuation byte (for example
				// "tambiÃÂ©n").  Remove that bridge before repairing the rest.
				text = Regex.Replace(text, "\u00C3\u00C2(?=[\u00C3\u0080-\u00BF])", "\u00C3");
				text = Regex.Replace(text, "\u00C3\u00C3(?=[\u0080-\u00BF])", "\u00C3");
				Encoding source = Encoding.GetEncoding(1252);
				for (int pass = 0; pass < 4 && LooksLikeMojibake(text); pass++)
				{
					string candidate = Encoding.UTF8.GetString(source.GetBytes(text));
					if (candidate.IndexOf('\uFFFD') >= 0) break;
					if (MojibakeScore(candidate) >= MojibakeScore(text)) break;
					text = candidate;
				}
			}
			catch
			{
			}
			return text;
		}

		private static string DecodeHtmlText(string text)
		{
			if (string.IsNullOrEmpty(text)) return string.Empty;
			text = Regex.Replace(text, @"&#x([0-9a-fA-F]+);", delegate(Match m)
			{
				int value;
				return int.TryParse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value)
					? char.ConvertFromUtf32(value) : string.Empty;
			});
			text = Regex.Replace(text, @"&#([0-9]+);", delegate(Match m)
			{
				int value;
				return int.TryParse(m.Groups[1].Value, out value) ? char.ConvertFromUtf32(value) : string.Empty;
			});
			return text.Replace("&nbsp;", " ").Replace("&amp;", "&").Replace("&lt;", "<")
				.Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'");
		}

		private static string ExtractHtmlFragment(string html)
		{
			if (string.IsNullOrEmpty(html)) return string.Empty;
			int start = html.IndexOf("<!--StartFragment-->", StringComparison.OrdinalIgnoreCase);
			if (start >= 0)
			{
				start += "<!--StartFragment-->".Length;
				int end = html.IndexOf("<!--EndFragment-->", start, StringComparison.OrdinalIgnoreCase);
				return end > start ? html.Substring(start, end - start) : html.Substring(start);
			}
			int body = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
			if (body >= 0)
			{
				int close = html.IndexOf('>', body);
				if (close >= 0) return html.Substring(close + 1);
			}
			return html;
		}

		private static List<FormattedRun> ParseHtmlRuns(string html)
		{
			string source = RepairMojibake(ExtractHtmlFragment(html));
			if (string.IsNullOrEmpty(source)) return null;
			source = Regex.Replace(source, @"<(script|style)[^>]*>.*?</\1>", string.Empty,
				RegexOptions.IgnoreCase | RegexOptions.Singleline);
			List<FormattedRun> runs = new List<FormattedRun>();
			Stack<KeyValuePair<string, RunFormat>> states = new Stack<KeyValuePair<string, RunFormat>>();
			RunFormat current = new RunFormat("Calibri", 22);
			bool hasText = false;
			foreach (Match match in Regex.Matches(source, @"<[^>]+>|[^<]+"))
			{
				string token = match.Value;
				if (token.StartsWith("<", StringComparison.Ordinal))
				{
					string lower = token.ToLowerInvariant();
					string tag = Regex.Match(lower, @"^</?\s*([a-z0-9]+)").Groups[1].Value;
					bool closing = lower.StartsWith("</", StringComparison.Ordinal);
					if (tag == "br" || tag == "hr")
					{
						runs.Add(new FormattedRun("\n", current.Clone()));
						continue;
					}
					if (closing)
					{
						if (states.Count > 0)
						{
							KeyValuePair<string, RunFormat> state = states.Pop();
							current = state.Value;
						}
						if (tag == "p" || tag == "div" || tag == "li" || tag == "tr" || tag == "h1"
							|| tag == "h2" || tag == "h3" || tag == "h4" || tag == "h5" || tag == "h6")
							runs.Add(new FormattedRun("\n", current.Clone()));
						continue;
					}
					if (tag == "img" || tag == "meta" || tag == "link" || tag == "input") continue;
					states.Push(new KeyValuePair<string, RunFormat>(tag, current.Clone()));
					ApplyHtmlTag(tag, token, ref current);
					continue;
				}
				string text = DecodeHtmlText(token);
				text = Regex.Replace(text, @"\s+", " ");
				if (text.Length == 0) continue;
				if (runs.Count > 0 && runs[runs.Count - 1].Format != null && runs[runs.Count - 1].Format.SameAs(current))
					runs[runs.Count - 1].Text += text;
				else
					runs.Add(new FormattedRun(text, current.Clone()));
				hasText = true;
			}
			if (!hasText) return null;
			// Trim leading/trailing structural newlines.
			while (runs.Count > 0 && runs[0].Text == "\n") runs.RemoveAt(0);
			while (runs.Count > 0 && runs[runs.Count - 1].Text == "\n") runs.RemoveAt(runs.Count - 1);
			return runs;
		}

		private static void ApplyHtmlTag(string tag, string token, ref RunFormat current)
		{
			if (tag == "a")
			{
				Match href = Regex.Match(token, @"href\s*=\s*[""']([^""']+)", RegexOptions.IgnoreCase);
				if (href.Success) current.LinkUrl = DecodeHtmlText(href.Groups[1].Value);
			}
			else if (tag == "b" || tag == "strong") current.Bold = true;
			else if (tag == "i" || tag == "em") current.Italic = true;
			else if (tag == "u") current.Underline = true;
			else if (tag == "s" || tag == "strike" || tag == "del") current.Strikeout = true;
			else if (tag == "h1" || tag == "h2" || tag == "h3" || tag == "h4" || tag == "h5" || tag == "h6")
			{
				int level = tag[1] - '0';
				current.Bold = true;
				current.SizeHalfPoints = Math.Max(22, 48 - (level - 1) * 6);
			}
			else if (tag == "font")
			{
				Match size = Regex.Match(token, @"size\s*=\s*[""']?([1-7])", RegexOptions.IgnoreCase);
				if (size.Success)
				{
					int[] sizes = new int[] { 0, 16, 20, 24, 28, 36, 48, 72 };
					current.SizeHalfPoints = sizes[int.Parse(size.Groups[1].Value)];
				}
				Match face = Regex.Match(token, @"face\s*=\s*[""']([^""']+)", RegexOptions.IgnoreCase);
				if (face.Success) current.FontFamily = face.Groups[1].Value;
				Match color = Regex.Match(token, @"color\s*=\s*[""']#?([0-9a-fA-F]{6})", RegexOptions.IgnoreCase);
				if (color.Success) current.ForeColor = ColorFromHtml(color.Groups[1].Value);
			}
			Match style = Regex.Match(token, @"style\s*=\s*[""']([^""']*)", RegexOptions.IgnoreCase);
			if (!style.Success) return;
			string[] declarations = style.Groups[1].Value.Split(';');
			for (int i = 0; i < declarations.Length; i++)
			{
				int colon = declarations[i].IndexOf(':');
				if (colon <= 0) continue;
				string key = declarations[i].Substring(0, colon).Trim().ToLowerInvariant();
				string value = declarations[i].Substring(colon + 1).Trim().ToLowerInvariant();
				if (key == "font-weight" && (value == "bold" || value == "700" || value == "800" || value == "900"))
					current.Bold = true;
				else if (key == "font-style" && value == "italic") current.Italic = true;
				else if (key == "text-decoration" && value.IndexOf("underline", StringComparison.Ordinal) >= 0)
					current.Underline = true;
				else if (key == "font-family")
				{
					string family = value.Split(',')[0].Trim().Trim('"', '\'');
					if (family.Length > 0) current.FontFamily = family;
				}
				else if (key == "font-size")
				{
					Match sizeMatch = Regex.Match(value, @"([0-9.]+)\s*(px|pt|em)?");
					if (sizeMatch.Success)
					{
						float number;
						if (float.TryParse(sizeMatch.Groups[1].Value, NumberStyles.Float,
							CultureInfo.InvariantCulture, out number))
						{
							float points = number;
							string unit = sizeMatch.Groups[2].Value;
							if (unit == "px") points = number * 0.75f;
							else if (unit == "em") points = number * (current.SizeHalfPoints / 2f);
							current.SizeHalfPoints = Math.Max(2, (int)Math.Round(points * 2f));
						}
					}
				}
				else if (key == "color")
				{
					Color? color = ParseHtmlColor(value);
					if (color.HasValue) current.ForeColor = color;
				}
			}
		}

		private static Color? ParseHtmlColor(string value)
		{
			if (string.IsNullOrEmpty(value)) return null;
			Match rgb = Regex.Match(value, @"rgba?\(([^)]+)\)", RegexOptions.IgnoreCase);
			if (rgb.Success)
			{
				string[] parts = rgb.Groups[1].Value.Split(',');
				if (parts.Length >= 3)
				{
					int r, g, b;
					if (int.TryParse(parts[0].Trim(), out r) && int.TryParse(parts[1].Trim(), out g)
						&& int.TryParse(parts[2].Trim(), out b))
						return Color.FromArgb(Math.Max(0, Math.Min(255, r)),
							Math.Max(0, Math.Min(255, g)), Math.Max(0, Math.Min(255, b)));
				}
			}
			Match hex = Regex.Match(value, @"#?([0-9a-fA-F]{6})");
			if (hex.Success)
			{
				try { return ColorFromHtml(hex.Groups[1].Value); }
				catch { }
			}
			return null;
		}

		private static Color ColorFromHtml(string hex)
		{
			return Color.FromArgb(Convert.ToInt32(hex.Substring(0, 2), 16),
				Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16));
		}

		private static List<FormattedRun> ParseRtfRuns(string rtf)
		{
			if (string.IsNullOrEmpty(rtf)) return null;
			try
			{
				DocumentModel model = RtfReader.Read(rtf);
				if (model == null || model.Paragraphs.Count == 0) return null;
				List<FormattedRun> runs = new List<FormattedRun>();
				for (int p = 0; p < model.Paragraphs.Count; p++)
				{
					Paragraph paragraph = model.Paragraphs[p];
					for (int i = 0; i < paragraph.Inlines.Count; i++)
					{
						TextInline text = paragraph.Inlines[i] as TextInline;
						if (text != null && text.Text.Length > 0)
							runs.Add(new FormattedRun(text.Text, text.Format == null
								? new RunFormat("Calibri", 22) : text.Format.Clone()));
					}
					if (p + 1 < model.Paragraphs.Count)
						runs.Add(new FormattedRun("\n", new RunFormat("Calibri", 22)));
				}
				return runs.Count > 0 ? runs : null;
			}
			catch { return null; }
		}

		public void Copy()
		{
			if (IsEditingTableCell)
			{
				_cellEditor.Copy();
				return;
			}
			if (_selectedFloating != null && !HasSelection)
			{
				_clipboardFloating = _selectedFloating.Clone();
				_clipboardIsShape = true;
				try
				{
					DataObject shapeData = new DataObject();
					shapeData.SetData("MiniWord.Shape", "1");
					Clipboard.SetDataObject(shapeData, true);
				}
				catch
				{
				}
				return;
			}
			if (!HasSelection)
			{
				return;
			}
			try
			{
				List<FormattedRun> runs = BuildSelectedRuns();
				DataObject data = new DataObject();
				string text = SelectedText;
				data.SetData(DataFormats.UnicodeText, text);
				data.SetData(DataFormats.Text, text);
				data.SetData("MiniWord.Runs", SerializeRuns(runs));
				string html = BuildClipboardHtml(runs);
				if (!string.IsNullOrEmpty(html)) data.SetData(DataFormats.Html, html);
				string rtf = BuildClipboardRtf(runs);
				if (!string.IsNullOrEmpty(rtf)) data.SetData(DataFormats.Rtf, rtf);
				Clipboard.SetDataObject(data, true);
				_clipboardIsShape = false;
			}
			catch
			{
				try { Clipboard.SetText(SelectedText); } catch { }
			}
		}

		public void Cut()
		{
			if (IsEditingTableCell)
			{
				_cellEditor.Cut();
				return;
			}
			if (_selectedFloating != null && !HasSelection)
			{
				Copy();
				DeleteSelectedFloating();
				return;
			}
			if (HasSelection)
			{
				Copy();
				InsertTextAtCaret(string.Empty);
			}
		}

		public void PastePlainText()
		{
			if (IsEditingTableCell)
			{
				_cellEditor.Paste();
				return;
			}
			try
			{
				if (!Clipboard.ContainsText()) return;
				string text = RepairMojibake(Clipboard.GetText());
				InsertFormattedRuns(new List<FormattedRun>
				{
					new FormattedRun(text, new RunFormat("Calibri", 22))
				});
				_clipboardIsShape = false;
			}
			catch
			{
			}
		}

		private static bool HasShapeClipboard()
		{
			try { return Clipboard.ContainsData("MiniWord.Shape"); }
			catch { return false; }
		}

		private bool TryPasteClipboardImage()
		{
			try
			{
				if (!Clipboard.ContainsImage()) return false;
				using (Image source = Clipboard.GetImage())
				{
					if (source == null || source.Width <= 0 || source.Height <= 0) return false;
					if ((long)source.Width * source.Height > 50000000L) return false;
					using (Bitmap bitmap = new Bitmap(source.Width, source.Height,
						System.Drawing.Imaging.PixelFormat.Format32bppArgb))
					{
						using (Graphics graphics = Graphics.FromImage(bitmap))
						{
							graphics.Clear(Color.Transparent);
							graphics.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
						}
						using (MemoryStream stream = new MemoryStream())
						{
							bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
							byte[] data = stream.ToArray();
							if (data.Length == 0 || data.Length > 32 * 1024 * 1024) return false;
							InsertImage(data, "png", source.Width, source.Height);
							_clipboardIsShape = false;
							return true;
						}
					}
				}
			}
			catch
			{
				return false;
			}
		}

		public void Paste()
		{
			if (IsEditingTableCell)
			{
				_cellEditor.Paste();
				return;
			}
			if (_clipboardIsShape && _clipboardFloating != null && HasShapeClipboard())
			{
				PushUndo();
				EnsureLayout();
				int page = Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage));
				int index = ParagraphForFloatingInsertion(page);
				FloatingObject copy = _clipboardFloating.Clone();
				Point offset = InsertionOffsetFor(page, index);
				copy.OffsetXTwips = offset.X;
				copy.OffsetYTwips = offset.Y;
				_document.Paragraphs[index].FloatingObjects.Add(copy);
				_selectedFloating = copy;
				_selectedImage = null;
				AfterEdit();
				Invalidate();
				return;
			}
			if (TryPasteClipboardImage()) return;
			try
			{
				List<FormattedRun> runs = null;
				object custom = Clipboard.GetData("MiniWord.Runs");
				if (custom != null) runs = DeserializeRuns(custom.ToString());
				if (runs == null || runs.Count == 0)
				{
					object rtfData = Clipboard.GetData(DataFormats.Rtf);
					if (rtfData != null) runs = ParseRtfRuns(rtfData.ToString());
				}
				if (runs == null || runs.Count == 0)
				{
					object htmlData = Clipboard.GetData(DataFormats.Html);
					if (htmlData == null) htmlData = Clipboard.GetData("HTML Format");
					if (htmlData != null) runs = ParseHtmlRuns(htmlData.ToString());
				}
				if (runs != null && runs.Count > 0)
				{
					InsertFormattedRuns(runs);
					_clipboardIsShape = false;
					return;
				}
				if (Clipboard.ContainsText())
				{
					// Plain text has no source formatting: use the document
					// default, never the Arial 30 currently selected in the toolbar.
					string text = RepairMojibake(Clipboard.GetText());
					InsertFormattedRuns(new List<FormattedRun>
					{
						new FormattedRun(text, new RunFormat("Calibri", 22))
					});
				}
			}
			catch
			{
			}
		}

		public void Undo()
		{
			if (_undo.Count != 0)
			{
				_redo.Push(CloneDocument(_document));
				_document = _undo.Pop();
				EnsureParagraphs();
				_caret = new CaretPos(0, 0);
				_anchor = new CaretPos(0, 0);
				_typingFormat = GetDefaultFormat();
				InvalidateLayout(true);
				OnTextChanged();
				OnSelectionChanged();
			}
		}

		public void Redo()
		{
			if (_redo.Count != 0)
			{
				_undo.Push(CloneDocument(_document));
				_document = _redo.Pop();
				EnsureParagraphs();
				_caret = new CaretPos(0, 0);
				_anchor = new CaretPos(0, 0);
				_typingFormat = GetDefaultFormat();
				InvalidateLayout(true);
				OnTextChanged();
				OnSelectionChanged();
			}
		}

		private RunFormat GetDefaultFormat()
		{
			return new RunFormat("Arial", 22);
		}

		public bool IsEditingTableCell
		{
			get { return _cellEditor != null && _cellEditor.Visible && _cellEdit != null; }
		}

		public RunFormat CellEditorFormat
		{
			get
			{
				if (!IsEditingTableCell) return null;
				TableCell cell = _cellEdit.Cell;
				return cell.Format != null ? cell.Format.Clone()
					: new RunFormat("Calibri", 21) { Bold = cell.Bold, Italic = cell.Italic };
			}
		}

		public void SetCellEditorCommitSuppressed(bool suppress)
		{
			_suppressCellCommit = suppress;
		}

		public void FocusCellEditor()
		{
			if (IsEditingTableCell) _cellEditor.Focus();
		}

		private bool ApplyContextFormat(Action<RunFormat> change)
		{
			if (change == null) return false;
			if (IsEditingTableCell)
			{
				TableCell cell = _cellEdit.Cell;
				RunFormat format = cell.Format == null
					? new RunFormat("Calibri", 21) { Bold = cell.Bold, Italic = cell.Italic }
					: cell.Format.Clone();
				change(format);
				cell.Format = format;
				cell.Bold = format.Bold;
				cell.Italic = format.Italic;
				int row = _cellEdit.Row;
				int column = _cellEdit.Column;
				int tableParagraph = _cellEditParagraph;
				int selectionStart = _cellEditor.SelectionStart;
				int selectionLength = _cellEditor.SelectionLength;
				RefreshCellEditorLayout(tableParagraph, row, column, selectionStart, selectionLength);
				OnTextChanged();
				return true;
			}
			if (_tableSelectionParagraph >= 0 && _tableSelectionParagraph < _document.Paragraphs.Count)
			{
				Paragraph table = _document.Paragraphs[_tableSelectionParagraph];
				table.EnsureTableCells();
				if (table.TableCells != null)
				{
					int row1 = Math.Min(_tableSelectionStartRow, _tableSelectionEndRow);
					int row2 = Math.Max(_tableSelectionStartRow, _tableSelectionEndRow);
					int col1 = Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn);
					int col2 = Math.Max(_tableSelectionStartColumn, _tableSelectionEndColumn);
					bool changed = false;
					for (int r = Math.Max(0, row1); r <= row2 && r < table.TableCells.Length; r++)
					{
						if (table.TableCells[r] == null) continue;
						for (int c = Math.Max(0, col1); c <= col2 && c < table.TableCells[r].Length; c++)
						{
							TableCell cell = table.TableCells[r][c];
							if (cell == null || cell.Covered) continue;
							RunFormat format = cell.Format == null
								? new RunFormat("Calibri", 21) { Bold = cell.Bold, Italic = cell.Italic }
								: cell.Format.Clone();
							change(format);
							cell.Format = format;
							cell.Bold = format.Bold;
							cell.Italic = format.Italic;
							changed = true;
						}
					}
					if (changed)
					{
						InvalidateLayout(true);
						OnTextChanged();
						Invalidate();
						return true;
					}
				}
			}
			if (_selectedFloating != null &&
				(_selectedFloating.Kind == FloatingKind.TextBox || _selectedFloating.Kind == FloatingKind.Shape))
			{
				RunFormat format = _selectedFloating.Format == null
					? new RunFormat("Calibri", 20) : _selectedFloating.Format.Clone();
				change(format);
				_selectedFloating.Format = format;
				Invalidate();
				OnTextChanged();
				return true;
			}
			return false;
		}

		private void UpdateCellEditorStyle()
		{
			if (!IsEditingTableCell) return;
			int start = _cellEditor.SelectionStart;
			int length = _cellEditor.SelectionLength;
			RunFormat format = _cellEdit.Cell.Format;
			if (format == null) return;
			try
			{
				FontStyle style = FontStyle.Regular;
				if (format.Bold) style |= FontStyle.Bold;
				if (format.Italic) style |= FontStyle.Italic;
				if (format.Underline) style |= FontStyle.Underline;
				_cellEditor.Font = new Font(format.FontFamily,
					Math.Max(6f, Math.Min(96f, format.SizeHalfPoints / 2f * _zoom)), style);
				_cellEditor.ForeColor = format.ForeColor ?? Color.Black;
				_cellEditor.SelectionStart = start;
				_cellEditor.SelectionLength = length;
			}
			catch
			{
			}
		}

		private bool ApplySelectionRunChange(Action<RunFormat> change)
		{
			if (!HasSelection) return false;
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph p = _document.Paragraphs[i];
				int from = i == start.P ? start.Offset : 0;
				int to = i == end.P ? end.Offset : GetText(p).Length;
				ApplyRunFormatChange(p, from, to, change);
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
			return true;
		}

		private void ApplyRunFormatChange(Paragraph p, int start, int end, Action<RunFormat> change)
		{
			if (start >= end) return;
			int offset = 0;
			List<Inline> result = new List<Inline>();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline run = p.Inlines[i] as TextInline;
				if (run == null)
				{
					result.Add(p.Inlines[i]);
					continue;
				}
				int length = run.Text.Length;
				int localStart = Math.Max(0, start - offset);
				int localEnd = Math.Min(length, end - offset);
				if (localEnd <= localStart)
				{
					result.Add(run);
				}
				else
				{
					if (localStart > 0)
						result.Add(new TextInline(run.Text.Substring(0, localStart), run.Format.Clone()));
					RunFormat selected = run.Format.Clone();
					change(selected);
					result.Add(new TextInline(run.Text.Substring(localStart, localEnd - localStart), selected));
					if (localEnd < length)
						result.Add(new TextInline(run.Text.Substring(localEnd), run.Format.Clone()));
				}
				offset += length;
			}
			p.Inlines.Clear();
			p.Inlines.AddRange(result);
			MergeRuns(p);
		}

		private void RefreshCellEditorLayout(int tableParagraph, int row, int column,
			int selectionStart, int selectionLength)
		{
			if (_cellEditor == null || !_cellEditor.Visible) return;
			_suppressCellLeave = true;
			try
			{
				InvalidateLayout(true);
				bool found = false;
				for (int i = 0; i < _tables.Count && !found; i++)
				{
					// Stay in the same table.  Several tables can have the same
					// row/column coordinates; picking the first one would move
					// the caret to a different table (often on a previous page).
					if (_tables[i].Paragraph != tableParagraph) continue;
					for (int j = 0; j < _tables[i].Cells.Count; j++)
					{
						LayoutTableCell candidate = _tables[i].Cells[j];
						if (candidate.Row == row && candidate.Column == column)
						{
							_cellEdit = candidate;
							found = true;
							break;
						}
					}
				}
				PositionCellEditor();
				_cellEditor.Visible = true;
				_cellEditor.BringToFront();
				UpdateCellEditorStyle();
				int length = _cellEditor.TextLength;
				_cellEditor.SelectionStart = Math.Max(0, Math.Min(length, selectionStart));
				_cellEditor.SelectionLength = Math.Max(0, Math.Min(length - _cellEditor.SelectionStart, selectionLength));
			}
			finally
			{
				_suppressCellLeave = false;
			}
		}

		public void ToggleBold()
		{
			if (ApplyContextFormat(delegate(RunFormat f) { f.Bold = !f.Bold; })) return;
			if (ApplySelectionRunChange(delegate(RunFormat f) { f.Bold = !f.Bold; })) return;
			_typingFormat.Bold = !_typingFormat.Bold;
			OnSelectionChanged();
		}

		public void ToggleItalic()
		{
			if (ApplyContextFormat(delegate(RunFormat f) { f.Italic = !f.Italic; })) return;
			if (ApplySelectionRunChange(delegate(RunFormat f) { f.Italic = !f.Italic; })) return;
			_typingFormat.Italic = !_typingFormat.Italic;
			OnSelectionChanged();
		}

		public void ToggleUnderline()
		{
			if (ApplyContextFormat(delegate(RunFormat f) { f.Underline = !f.Underline; })) return;
			if (ApplySelectionRunChange(delegate(RunFormat f) { f.Underline = !f.Underline; })) return;
			_typingFormat.Underline = !_typingFormat.Underline;
			OnSelectionChanged();
		}

		public void SetFontFamily(string family)
		{
			if (string.IsNullOrEmpty(family)) return;
			if (ApplyContextFormat(delegate(RunFormat f) { f.FontFamily = family; })) return;
			if (ApplySelectionRunChange(delegate(RunFormat f) { f.FontFamily = family; })) return;
			_typingFormat.FontFamily = family;
			OnSelectionChanged();
		}

		public void SetFontSize(float points)
		{
			if (points < 4f) points = 4f;
			if (points > 200f) points = 200f;
			int halfPoints = Math.Max(2, (int)Math.Round(points * 2f));
			if (ApplyContextFormat(delegate(RunFormat f) { f.SizeHalfPoints = halfPoints; })) return;
			if (!HasSelection)
			{
				_typingFormat = _typingFormat.Clone();
				_typingFormat.SizeHalfPoints = halfPoints;
				OnSelectionChanged();
				return;
			}
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph p = _document.Paragraphs[i];
				int from = i == start.P ? start.Offset : 0;
				int to = i == end.P ? end.Offset : GetText(p).Length;
				ApplySizeToParagraph(p, from, to, halfPoints);
			}
			_typingFormat = _typingFormat.Clone();
			_typingFormat.SizeHalfPoints = halfPoints;
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		private void ApplySizeToParagraph(Paragraph p, int start, int end, int halfPoints)
		{
			if (start >= end) return;
			int offset = 0;
			List<Inline> result = new List<Inline>();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline run = p.Inlines[i] as TextInline;
				if (run == null)
				{
					result.Add(p.Inlines[i]);
					continue;
				}
				int length = run.Text.Length;
				int localStart = Math.Max(0, start - offset);
				int localEnd = Math.Min(length, end - offset);
				if (localEnd <= localStart)
				{
					result.Add(run);
				}
				else
				{
					if (localStart > 0)
						result.Add(new TextInline(run.Text.Substring(0, localStart), run.Format.Clone()));
					RunFormat selected = run.Format.Clone();
					selected.SizeHalfPoints = halfPoints;
					result.Add(new TextInline(run.Text.Substring(localStart, localEnd - localStart), selected));
					if (localEnd < length)
						result.Add(new TextInline(run.Text.Substring(localEnd), run.Format.Clone()));
				}
				offset += length;
			}
			p.Inlines.Clear();
			p.Inlines.AddRange(result);
			MergeRuns(p);
		}

		private string LinkAtPosition(CaretPos position)
		{
			if (position.P < 0 || position.P >= _document.Paragraphs.Count) return null;
			Paragraph paragraph = _document.Paragraphs[position.P];
			int offset = 0;
			for (int i = 0; i < paragraph.Inlines.Count; i++)
			{
				TextInline run = paragraph.Inlines[i] as TextInline;
				if (run == null) continue;
				int start = offset;
				int end = offset + run.Text.Length;
				if (run.Format != null && position.Offset >= start && position.Offset <= end
					&& !string.IsNullOrEmpty(run.Format.LinkUrl))
					return run.Format.LinkUrl;
				offset = end;
			}
			return null;
		}

		public string SelectedHyperlinkUrl
		{
			get
			{
				if (IsEditingTableCell && _cellEdit != null && _cellEdit.Cell != null)
					return _cellEdit.Cell.LinkUrl;
				if (_tableSelectionParagraph >= 0 && _tableSelectionParagraph < _document.Paragraphs.Count)
				{
					Paragraph table = _document.Paragraphs[_tableSelectionParagraph];
					table.EnsureTableCells();
					if (table.TableCells != null)
					{
						int r = Math.Max(0, Math.Min(table.TableCells.Length - 1,
							Math.Min(_tableSelectionStartRow, _tableSelectionEndRow)));
						if (r >= 0 && table.TableCells[r] != null)
						{
							int c = Math.Max(0, Math.Min(table.TableCells[r].Length - 1,
								Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn)));
							if (c >= 0 && table.TableCells[r][c] != null)
								return table.TableCells[r][c].LinkUrl;
						}
					}
				}
				if (!HasSelection) return LinkAtPosition(_caret);
				CaretPos start;
				CaretPos end;
				NormalizeSelection(out start, out end);
				string first = LinkAtPosition(start);
				if (!string.IsNullOrEmpty(first)) return first;
				return LinkAtPosition(new CaretPos(start.P, Math.Max(0, start.Offset + 1)));
			}
		}

		public bool ApplyHyperlink(string url)
		{
			string target = (url ?? string.Empty).Trim();
			if (target.Length == 0) return false;
			if (IsEditingTableCell && _cellEdit != null && _cellEdit.Cell != null)
			{
				PushUndo();
				_cellEdit.Cell.LinkUrl = target;
				InvalidateLayout(true);
				OnTextChanged();
				return true;
			}
			if (_tableSelectionParagraph >= 0 && _tableSelectionParagraph < _document.Paragraphs.Count)
			{
				Paragraph table = _document.Paragraphs[_tableSelectionParagraph];
				table.EnsureTableCells();
				if (table.TableCells != null)
				{
					PushUndo();
					int firstRow = Math.Max(0, Math.Min(_tableSelectionStartRow, _tableSelectionEndRow));
					int lastRow = Math.Min(table.TableCells.Length - 1,
						Math.Max(_tableSelectionStartRow, _tableSelectionEndRow));
					int firstColumn = Math.Max(0, Math.Min(_tableSelectionStartColumn, _tableSelectionEndColumn));
					int lastColumn = Math.Max(_tableSelectionStartColumn, _tableSelectionEndColumn);
					for (int r = firstRow; r <= lastRow; r++)
					{
						if (table.TableCells[r] == null) continue;
						for (int c = firstColumn; c <= lastColumn && c < table.TableCells[r].Length; c++)
							if (table.TableCells[r][c] != null) table.TableCells[r][c].LinkUrl = target;
					}
					InvalidateLayout(true);
					OnTextChanged();
					return true;
				}
			}
			if (HasSelection)
			{
				return ApplySelectionRunChange(delegate(RunFormat format)
				{
					format.LinkUrl = target;
				});
			}
			RunFormat linkFormat = _typingFormat.Clone();
			linkFormat.LinkUrl = target;
			InsertFormattedRuns(new List<FormattedRun>
			{
				new FormattedRun(target, linkFormat)
			});
			return true;
		}

		public void SetForeColor(Color color)
		{
			if (ApplyContextFormat(delegate(RunFormat f) { f.ForeColor = color; })) return;
			if (ApplySelectionRunChange(delegate(RunFormat f) { f.ForeColor = color; })) return;
			_typingFormat.ForeColor = color;
			OnSelectionChanged();
		}

		public void ApplySelectionFormat(RunFormat format)
		{
			if (format == null)
			{
				return;
			}
			if (!HasSelection)
			{
				_typingFormat = format.Clone();
				OnSelectionChanged();
				return;
			}
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph p = _document.Paragraphs[i];
				int start2 = ((i == start.P) ? start.Offset : 0);
				int end2 = ((i == end.P) ? end.Offset : GetText(p).Length);
				ApplyFormatToParagraph(p, start2, end2, format.Clone());
			}
			_typingFormat = format.Clone();
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		private void ApplyFormatToParagraph(Paragraph p, int start, int end, RunFormat format)
		{
			if (start >= end)
			{
				return;
			}
			int num = 0;
			List<Inline> list = new List<Inline>();
			for (int i = 0; i < p.Inlines.Count; i++)
			{
				TextInline textInline = p.Inlines[i] as TextInline;
				if (textInline == null)
				{
					list.Add(p.Inlines[i]);
					continue;
				}
				int length = textInline.Text.Length;
				int num2 = Math.Max(0, start - num);
				int num3 = Math.Min(length, end - num);
				if (num3 <= num2)
				{
					list.Add(textInline);
				}
				else
				{
					if (num2 > 0)
					{
						list.Add(new TextInline(textInline.Text.Substring(0, num2), textInline.Format.Clone()));
					}
					TextInline item = new TextInline(textInline.Text.Substring(num2, num3 - num2), format.Clone());
					list.Add(item);
					if (num3 < length)
					{
						list.Add(new TextInline(textInline.Text.Substring(num3), textInline.Format.Clone()));
					}
				}
				num += length;
			}
			p.Inlines.Clear();
			p.Inlines.AddRange(list);
			MergeRuns(p);
		}

		public void SetAlignment(ParagraphAlignment alignment)
		{
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				_document.Paragraphs[i].Alignment = alignment;
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		public void ApplyParagraphFormatting(ParagraphAlignment alignment, int leftIndent, int rightIndent, int firstLineIndent, int spaceBefore, int spaceAfter, int lineSpacing, bool pageBreak, bool noSpaceSameStyle, int outlineLevel, bool keepLines, bool keepNext, bool widowControl)
		{
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph paragraph = _document.Paragraphs[i];
				paragraph.Alignment = alignment;
				paragraph.LeftIndentTwips = Math.Max(0, leftIndent);
				paragraph.RightIndentTwips = Math.Max(0, rightIndent);
				paragraph.FirstLineIndentTwips = firstLineIndent;
				paragraph.SpaceBeforeTwips = Math.Max(0, spaceBefore);
				paragraph.SpaceAfterTwips = Math.Max(0, spaceAfter);
				paragraph.LineSpacing = Math.Max(0, Math.Min(2, lineSpacing));
				paragraph.PageBreakBefore = pageBreak;
				paragraph.NoSpaceBetweenSameStyle = noSpaceSameStyle;
				paragraph.OutlineLevel = Math.Max(0, Math.Min(8, outlineLevel));
				paragraph.KeepLinesTogether = keepLines;
				paragraph.KeepWithNext = keepNext;
				paragraph.WidowControl = widowControl;
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		private void ChangeIndent(int delta)
		{
			PushUndo();
			CaretPos start, end;
			NormalizeSelection(out start, out end);
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph p = _document.Paragraphs[i];
				p.LeftIndentTwips = Math.Max(0, p.LeftIndentTwips + delta);
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		public void IncreaseIndent()
		{
			ChangeIndent(360);
		}

		public void DecreaseIndent()
		{
			ChangeIndent(-360);
		}

		public void ToggleBullet()
		{
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			bool flag = !_document.Paragraphs[start.P].Bullet;
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph paragraph = _document.Paragraphs[i];
				paragraph.Bullet = flag;
				if (flag)
				{
					paragraph.Numbered = false;
					if (paragraph.LeftIndentTwips < 360)
					{
						paragraph.LeftIndentTwips = 360;
					}
					if (paragraph.FirstLineIndentTwips < 0)
					{
						paragraph.FirstLineIndentTwips = 0;
					}
				}
				else
				{
					paragraph.LeftIndentTwips = 0;
					paragraph.FirstLineIndentTwips = 0;
				}
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		public void ToggleNumberedList()
		{
			PushUndo();
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			bool flag = !_document.Paragraphs[start.P].Numbered;
			for (int i = start.P; i <= end.P; i++)
			{
				Paragraph paragraph = _document.Paragraphs[i];
				paragraph.Numbered = flag;
				if (flag)
				{
					paragraph.Bullet = false;
					if (paragraph.LeftIndentTwips < 360)
					{
						paragraph.LeftIndentTwips = 360;
					}
					paragraph.FirstLineIndentTwips = 0;
				}
				else
				{
					paragraph.LeftIndentTwips = 0;
					paragraph.FirstLineIndentTwips = 0;
				}
			}
			InvalidateLayout(true);
			OnTextChanged();
			OnSelectionChanged();
		}

		public void InsertPageBreak()
		{
			PushUndo();
			if (HasSelection)
			{
				DeleteSelectionInternal();
			}
			if (_caret.Offset > 0)
			{
				SplitParagraphAt(_caret.P, _caret.Offset);
				_caret = new CaretPos(_caret.P + 1, 0);
				_anchor = new CaretPos(_caret.P, 0);
				_document.Paragraphs[_caret.P].PageBreakBefore = true;
			}
			else
			{
				int p = _caret.P;
				Paragraph source = _document.Paragraphs[p];
				Paragraph paragraph = CloneParagraph(source);
				paragraph.Inlines.Clear();
				paragraph.FloatingObjects.Clear();
				paragraph.TableRows = null;
				paragraph.PageBreakBefore = false;
				paragraph.SpaceBeforeTwips = 0;
				paragraph.SpaceAfterTwips = 0;
				paragraph.Bullet = false;
				_document.Paragraphs.Insert(p, paragraph);
				_document.Paragraphs[p + 1].PageBreakBefore = true;
				_caret = new CaretPos(p + 1, 0);
				_anchor = new CaretPos(p + 1, 0);
			}
			AfterEdit();
		}

		public void ClearSpellingHighlight()
		{
			_spellingParagraph = -1;
			_spellingLength = 0;
			if (!IsDisposed) Invalidate();
		}

		public void SelectSpellingRange(int paragraph, int start, int length)
		{
			SelectRange(paragraph, start, length);
			_spellingParagraph = paragraph;
			_spellingStart = start;
			_spellingLength = length;
			EnsureLayout();
			int index = Math.Max(0, Math.Min(_document.Paragraphs.Count - 1, paragraph));
			float lineY = -1f;
			for (int i = 0; i < _lines.Count; i++)
			{
				if (_lines[i].Paragraph != index) continue;
				if (start >= _lines[i].Start && start <= _lines[i].End)
				{
					lineY = _lines[i].Y;
					break;
				}
			}
			if (lineY >= 0f)
			{
				float target = lineY - Math.Max(80f, ClientSize.Height * 0.62f);
				int y = Math.Max(0, (int)Math.Round(target));
				base.AutoScrollPosition = new Point(-base.AutoScrollPosition.X, y);
				Invalidate();
			}
		}

		public void SelectRange(int paragraph, int start, int length)
		{
			_spellingParagraph = -1;
			_spellingLength = 0;
			if (paragraph >= 0 && paragraph < _document.Paragraphs.Count)
			{
				int length2 = GetText(_document.Paragraphs[paragraph]).Length;
				start = Math.Max(0, Math.Min(length2, start));
				length = Math.Max(0, Math.Min(length2 - start, length));
				_anchor = new CaretPos(paragraph, start);
				_caret = new CaretPos(paragraph, start + length);
				_typingFormat = FormatAt(_caret);
				EnsureCaretVisible();
				Invalidate();
				OnSelectionChanged();
			}
		}

		public bool FindNext(string text, bool matchCase, bool forward)
		{
			if (string.IsNullOrEmpty(text))
			{
				return false;
			}
			CaretPos start;
			CaretPos end;
			NormalizeSelection(out start, out end);
			for (int i = 0; i < 2; i++)
			{
				int num = ((i == 0) ? _caret.P : 0);
				int val = ((i == 0) ? _caret.Offset : 0);
				for (int j = num; j < _document.Paragraphs.Count; j++)
				{
					string text2 = GetText(_document.Paragraphs[j]);
					int startIndex = ((j == num) ? Math.Min(val, text2.Length) : 0);
					int num2 = text2.IndexOf(text, startIndex, matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
					if (num2 >= 0)
					{
						SelectRange(j, num2, text.Length);
						return true;
					}
				}
			}
			return false;
		}

		public void ReplaceSelection(string replacement)
		{
			InsertTextAtCaret(replacement ?? string.Empty);
		}

		public void SetPageSize(int widthTwips, int heightTwips)
		{
			_document.PageWidthTwips = widthTwips;
			_document.PageHeightTwips = heightTwips;
			InvalidateLayout(true);
			OnTextChanged();
		}

		public void SetMargins(int left, int top, int right, int bottom)
		{
			_document.MarginLeftTwips = Math.Max(0, left);
			_document.MarginTopTwips = Math.Max(0, top);
			_document.MarginRightTwips = Math.Max(0, right);
			_document.MarginBottomTwips = Math.Max(0, bottom);
			InvalidateLayout(true);
			OnTextChanged();
		}

		public void GetRulerMetrics(out float pageLeft, out float pageWidth,
			out float leftMargin, out float rightMargin)
		{
			EnsureLayout();
			LayoutPage page = _pages.Count == 0 ? null
				: _pages[Math.Max(0, Math.Min(_pages.Count - 1, ViewedPage))];
			if (page == null)
			{
				pageLeft = 18f;
				pageWidth = Math.Max(1f, ClientSize.Width - 36f);
				leftMargin = 0f;
				rightMargin = 0f;
				return;
			}
			pageLeft = page.Left + base.AutoScrollPosition.X;
			pageWidth = page.Width;
			leftMargin = Math.Max(0f, page.BodyLeft - page.Left);
			rightMargin = Math.Max(0f, page.Left + page.Width - page.BodyRight);
		}

		public DocumentSection CurrentSection
		{
			get
			{
				EnsureParagraphs();
				return _document.SectionForParagraph(_caret.P);
			}
		}

		private static RunFormat HeaderFormatAt(RunFormat[] formats, int index)
		{
			if (formats == null || formats.Length <= index || formats[index] == null)
				return new RunFormat("Calibri", 18);
			return formats[index].Clone();
		}

		private static ParagraphAlignment HeaderAlignmentAt(ParagraphAlignment[] alignments, int index)
		{
			if (alignments == null || alignments.Length <= index) return ParagraphAlignment.Left;
			return alignments[index];
		}

		private static List<ImageInline> HeaderImagesAt(List<ImageInline>[] sets, int index)
		{
			List<ImageInline> result = new List<ImageInline>();
			if (sets == null || sets.Length <= index || sets[index] == null) return result;
			for (int i = 0; i < sets[index].Count; i++)
			{
				ImageInline image = sets[index][i];
				result.Add(new ImageInline(image.Data, image.Extension, image.WidthTwips,
					image.HeightTwips, image.Format.Clone()));
			}
			return result;
		}

		public void SetCurrentHeaderFooter(string header, string footer,
			string firstHeader, string firstFooter, string evenHeader, string evenFooter,
			bool differentFirstPage, bool differentOddEven,
			int headerDistance, int footerDistance,
			RunFormat[] headerFormats, RunFormat[] footerFormats,
			ParagraphAlignment[] headerAlignments, ParagraphAlignment[] footerAlignments,
			List<ImageInline>[] headerImages, List<ImageInline>[] footerImages)
		{
			DocumentSection section = CurrentSection;
			section.HeaderText = header ?? string.Empty;
			section.FooterText = footer ?? string.Empty;
			section.FirstPageHeaderText = firstHeader ?? string.Empty;
			section.FirstPageFooterText = firstFooter ?? string.Empty;
			section.EvenPageHeaderText = evenHeader ?? string.Empty;
			section.EvenPageFooterText = evenFooter ?? string.Empty;
			section.TitlePage = differentFirstPage;
			section.DifferentOddEven = differentOddEven;
			section.HeaderDistanceTwips = Math.Max(0, headerDistance);
			section.FooterDistanceTwips = Math.Max(0, footerDistance);
			section.HeaderFormat = HeaderFormatAt(headerFormats, 0);
			section.FooterFormat = HeaderFormatAt(footerFormats, 0);
			section.FirstPageHeaderFormat = HeaderFormatAt(headerFormats, 1);
			section.FirstPageFooterFormat = HeaderFormatAt(footerFormats, 1);
			section.EvenPageHeaderFormat = HeaderFormatAt(headerFormats, 2);
			section.EvenPageFooterFormat = HeaderFormatAt(footerFormats, 2);
			section.HeaderAlignment = HeaderAlignmentAt(headerAlignments, 0);
			section.FooterAlignment = HeaderAlignmentAt(footerAlignments, 0);
			section.FirstPageHeaderAlignment = HeaderAlignmentAt(headerAlignments, 1);
			section.FirstPageFooterAlignment = HeaderAlignmentAt(footerAlignments, 1);
			section.EvenPageHeaderAlignment = HeaderAlignmentAt(headerAlignments, 2);
			section.EvenPageFooterAlignment = HeaderAlignmentAt(footerAlignments, 2);
			section.HeaderImages = HeaderImagesAt(headerImages, 0);
			section.FooterImages = HeaderImagesAt(footerImages, 0);
			section.FirstPageHeaderImages = HeaderImagesAt(headerImages, 1);
			section.FirstPageFooterImages = HeaderImagesAt(footerImages, 1);
			section.EvenPageHeaderImages = HeaderImagesAt(headerImages, 2);
			section.EvenPageFooterImages = HeaderImagesAt(footerImages, 2);
			if (_document.Sections.Count > 0 && object.ReferenceEquals(_document.Sections[0], section))
			{
				_document.HeaderText = section.HeaderText;
				_document.FooterText = section.FooterText;
				_document.HeaderDistanceTwips = section.HeaderDistanceTwips;
				_document.FooterDistanceTwips = section.FooterDistanceTwips;
			}
			RefreshDocument();
		}

		public void RefreshDocument()
		{
			NormalizeDocumentText(_document);
			InvalidateLayout(true);
			OnTextChanged();
		}

		private static DocumentModel CloneDocument(DocumentModel source)
		{
			DocumentModel documentModel = new DocumentModel();
			documentModel.PageWidthTwips = source.PageWidthTwips;
			documentModel.PageHeightTwips = source.PageHeightTwips;
			documentModel.MarginLeftTwips = source.MarginLeftTwips;
			documentModel.MarginRightTwips = source.MarginRightTwips;
			documentModel.MarginTopTwips = source.MarginTopTwips;
			documentModel.MarginBottomTwips = source.MarginBottomTwips;
			documentModel.Title = source.Title;
			documentModel.Author = source.Author;
			documentModel.HeaderText = source.HeaderText;
			documentModel.FooterText = source.FooterText;
			documentModel.HeaderDistanceTwips = source.HeaderDistanceTwips;
			documentModel.FooterDistanceTwips = source.FooterDistanceTwips;
			source.EnsureSections();
			for (int i = 0; i < source.Sections.Count; i++)
			{
				documentModel.Sections.Add(source.Sections[i].Clone());
			}
			for (int j = 0; j < source.Footnotes.Count; j++)
			{
				documentModel.Footnotes.Add(source.Footnotes[j].Clone());
			}
			for (int k = 0; k < source.Endnotes.Count; k++)
			{
				documentModel.Endnotes.Add(source.Endnotes[k].Clone());
			}
			for (int l = 0; l < source.Comments.Count; l++)
			{
				documentModel.Comments.Add(source.Comments[l].Clone());
			}
			for (int n = 0; n < source.UnsupportedFeatures.Count; n++)
			{
				documentModel.UnsupportedFeatures.Add(source.UnsupportedFeatures[n]);
			}
			if (source.PreservedParts != null)
			{
				foreach (KeyValuePair<string, byte[]> part in source.PreservedParts)
					documentModel.PreservedParts[part.Key] = part.Value;
			}
			if (source.PreservedRelationships != null)
			{
				for (int i = 0; i < source.PreservedRelationships.Count; i++)
				{
					PreservedRelationship original = source.PreservedRelationships[i];
					PreservedRelationship copy = new PreservedRelationship();
					copy.Type = original.Type;
					copy.Target = original.Target;
					copy.TargetMode = original.TargetMode;
					documentModel.PreservedRelationships.Add(copy);
				}
			}
			if (source.PreservedContentTypes != null)
			{
				for (int i = 0; i < source.PreservedContentTypes.Count; i++)
				{
					PreservedContentType original = source.PreservedContentTypes[i];
					PreservedContentType copy = new PreservedContentType();
					copy.Extension = original.Extension;
					copy.PartName = original.PartName;
					copy.ContentType = original.ContentType;
					documentModel.PreservedContentTypes.Add(copy);
				}
			}
			for (int m = 0; m < source.Paragraphs.Count; m++)
			{
				documentModel.Paragraphs.Add(CloneParagraph(source.Paragraphs[m]));
			}
			return documentModel;
		}

		private static Paragraph CloneParagraph(Paragraph source)
		{
			Paragraph paragraph = new Paragraph();
			paragraph.PreservedXml = source.PreservedXml;
			paragraph.PreservedText = source.PreservedText;
			paragraph.PreservedFormatKey = source.PreservedFormatKey;
			paragraph.Alignment = source.Alignment;
			paragraph.LeftIndentTwips = source.LeftIndentTwips;
			paragraph.RightIndentTwips = source.RightIndentTwips;
			paragraph.FirstLineIndentTwips = source.FirstLineIndentTwips;
			paragraph.SpaceBeforeTwips = source.SpaceBeforeTwips;
			paragraph.SpaceAfterTwips = source.SpaceAfterTwips;
			paragraph.LineSpacing = source.LineSpacing;
			paragraph.Bullet = source.Bullet;
			paragraph.BulletLevel = source.BulletLevel;
			paragraph.Numbered = source.Numbered;
			paragraph.NumberLevel = source.NumberLevel;
			paragraph.PageBreakBefore = source.PageBreakBefore;
			paragraph.SectionIndex = source.SectionIndex;
			paragraph.OutlineLevel = source.OutlineLevel;
			paragraph.HasModeledField = source.HasModeledField;
			paragraph.IsTocEntry = source.IsTocEntry;
			paragraph.IsTocTitle = source.IsTocTitle;
			paragraph.TocLevel = source.TocLevel;
			paragraph.TocTargetParagraph = source.TocTargetParagraph;
			paragraph.TocPage = source.TocPage;
			paragraph.NoSpaceBetweenSameStyle = source.NoSpaceBetweenSameStyle;
			paragraph.KeepLinesTogether = source.KeepLinesTogether;
			paragraph.KeepWithNext = source.KeepWithNext;
			paragraph.WidowControl = source.WidowControl;
			paragraph.TabStopsTwips = ((source.TabStopsTwips == null) ? new int[0] : ((int[])source.TabStopsTwips.Clone()));
			paragraph.TableWidthTwips = source.TableWidthTwips;
			paragraph.TableStyleName = source.TableStyleName;
			paragraph.TableShading = source.TableShading;
			paragraph.TableBorderColor = source.TableBorderColor;
			paragraph.TableBorderSize = source.TableBorderSize;
			if (source.TableRows != null)
			{
				paragraph.TableRows = new string[source.TableRows.Length][];
				for (int i = 0; i < source.TableRows.Length; i++)
				{
					paragraph.TableRows[i] = ((source.TableRows[i] == null) ? new string[0] : ((string[])source.TableRows[i].Clone()));
				}
			}
			if (source.TableCells != null)
			{
				paragraph.TableCells = new TableCell[source.TableCells.Length][];
				for (int j = 0; j < source.TableCells.Length; j++)
				{
					paragraph.TableCells[j] = ((source.TableCells[j] == null) ? new TableCell[0] : new TableCell[source.TableCells[j].Length]);
					for (int k = 0; k < paragraph.TableCells[j].Length; k++)
					{
						paragraph.TableCells[j][k] = ((source.TableCells[j][k] == null) ? null : source.TableCells[j][k].Clone());
					}
				}
			}
			for (int l = 0; l < source.FloatingObjects.Count; l++)
			{
				paragraph.FloatingObjects.Add(source.FloatingObjects[l].Clone());
			}
			for (int m = 0; m < source.Inlines.Count; m++)
			{
				TextInline textInline = source.Inlines[m] as TextInline;
				if (textInline != null)
				{
					paragraph.Inlines.Add(textInline.Clone());
					continue;
				}
				ImageInline imageInline = source.Inlines[m] as ImageInline;
				if (imageInline != null)
				{
					paragraph.Inlines.Add(new ImageInline(imageInline.Data, imageInline.Extension, imageInline.WidthTwips, imageInline.HeightTwips, imageInline.Format.Clone()));
				}
				NoteInline noteInline = source.Inlines[m] as NoteInline;
				if (noteInline != null)
				{
					paragraph.Inlines.Add(new NoteInline(noteInline.Note, noteInline.Format.Clone()));
				}
				CommentInline commentInline = source.Inlines[m] as CommentInline;
				if (commentInline != null)
				{
					paragraph.Inlines.Add(new CommentInline(commentInline.Comment, commentInline.Format.Clone()));
				}
			}
			return paragraph;
		}

		private void ScheduleSpellCheck()
		{
			if (_spellEnabled && _document != null && _spellTimer != null)
			{
				_spellErrors.Clear();
				_spellTimer.Stop();
				_spellTimer.Start();
			}
		}

		private void RunSpellCheck()
		{
			if (!_spellEnabled || base.IsDisposed)
			{
				return;
			}
			if (_spellChecker == null)
			{
				_spellChecker = new SystemSpellChecker(_spellLanguage);
				_spellUnavailableAttempts = 0;
			}
			if (!_spellChecker.IsAvailable)
			{
				if (++_spellUnavailableAttempts < 6 && _spellTimer != null)
				{
					_spellTimer.Start();
				}
				return;
			}
			_spellUnavailableAttempts = 0;
			string[] array = new string[_document.Paragraphs.Count];
			for (int i = 0; i < _document.Paragraphs.Count; i++)
			{
				Paragraph paragraph = _document.Paragraphs[i];
				array[i] = ((paragraph.TableRows == null) ? GetText(paragraph) : string.Empty);
			}
			int version = ++_spellVersion;
			_spellChecker.Check(array, version, delegate(int resultVersion, List<SpellError> errors)
			{
				if (!base.IsDisposed && resultVersion == _spellVersion)
				{
					_spellErrors.Clear();
					if (errors != null)
					{
						for (int j = 0; j < errors.Count; j++)
						{
							_spellErrors.Add(errors[j]);
						}
					}
					Invalidate();
				}
			});
		}

		private SpellError FindSpellError(int paragraph, int offset)
		{
			for (int i = 0; i < _spellErrors.Count; i++)
			{
				SpellError spellError = _spellErrors[i];
				if (spellError.Paragraph == paragraph && offset >= spellError.Start && offset <= spellError.Start + spellError.Length)
				{
					return spellError;
				}
			}
			return null;
		}

		private void DrawSpellUnderlines(Graphics g, LayoutLine line)
		{
			if (!_spellEnabled || _spellErrors.Count == 0)
			{
				return;
			}
			for (int i = 0; i < _spellErrors.Count; i++)
			{
				SpellError spellError = _spellErrors[i];
				if (spellError.Paragraph != line.Paragraph)
				{
					continue;
				}
				for (int j = 0; j < line.Segments.Count; j++)
				{
					LayoutSegment layoutSegment = line.Segments[j];
					if (string.IsNullOrEmpty(layoutSegment.Text))
					{
						continue;
					}
					int num = Math.Max(layoutSegment.Start, spellError.Start);
					int num2 = Math.Min(layoutSegment.End, spellError.Start + spellError.Length);
					if (num2 <= num)
					{
						continue;
					}
					float num3 = XForOffset(layoutSegment, num);
					float num4 = XForOffset(layoutSegment, num2);
					if (num4 <= num3)
					{
						continue;
					}
					float num5 = layoutSegment.Y + layoutSegment.Height - 1f;
					using (Pen pen = new Pen(Color.FromArgb(190, 32, 38), 1f))
					{
						float num6 = num3;
						bool flag = true;
						while (num6 < num4)
						{
							float num7 = Math.Min(num6 + 3f, num4);
							g.DrawLine(pen, num6, flag ? (num5 - 1f) : (num5 + 1f), num7, flag ? (num5 + 1f) : (num5 - 1f));
							flag = !flag;
							num6 = num7;
						}
					}
				}
			}
		}

		private void ShowSpellMenu(SpellError error, Point location)
		{
			ContextMenuStrip menu = new ContextMenuStrip();
			ToolStripMenuItem toolStripMenuItem = new ToolStripMenuItem("Buscando sugerencias...");
			toolStripMenuItem.Enabled = false;
			menu.Items.Add(toolStripMenuItem);
			menu.Show(this, location);
			if (_spellChecker == null || !_spellChecker.IsAvailable)
			{
				toolStripMenuItem.Text = "Corrección ortográfica del sistema no disponible";
				return;
			}
			_spellChecker.Suggest(error.Word, delegate(List<string> suggestions)
			{
				if (!base.IsDisposed && !menu.IsDisposed)
				{
					menu.Items.Clear();
					if (suggestions == null || suggestions.Count == 0)
					{
						ToolStripMenuItem value = new ToolStripMenuItem("Sin sugerencias")
						{
							Enabled = false
						};
						menu.Items.Add(value);
					}
					else
					{
						for (int i = 0; i < suggestions.Count; i++)
						{
							string replacement = suggestions[i];
							ToolStripMenuItem toolStripMenuItem2 = new ToolStripMenuItem(replacement);
							toolStripMenuItem2.Font = new Font(Font, FontStyle.Bold);
							toolStripMenuItem2.Click += delegate
							{
								SelectRange(error.Paragraph, error.Start, error.Length);
								InsertText(replacement);
							};
							menu.Items.Add(toolStripMenuItem2);
						}
					}
					menu.Items.Add(new ToolStripSeparator());
					ToolStripMenuItem toolStripMenuItem3 = new ToolStripMenuItem("Omitir");
					toolStripMenuItem3.Click += delegate
					{
						if (_spellChecker != null)
						{
							_spellChecker.Ignore(error.Word);
						}
						_spellErrors.Remove(error);
						Invalidate();
					};
					menu.Items.Add(toolStripMenuItem3);
					ToolStripMenuItem toolStripMenuItem4 = new ToolStripMenuItem("Agregar al diccionario");
					toolStripMenuItem4.Click += delegate
					{
						if (_spellChecker != null)
						{
							_spellChecker.Add(error.Word);
						}
						_spellErrors.Remove(error);
						Invalidate();
					};
					menu.Items.Add(toolStripMenuItem4);
				}
			});
		}

		private void OnTextChanged()
		{
			ScheduleSpellCheck();
			if (this.TextChanged != null)
			{
				this.TextChanged(this, EventArgs.Empty);
			}
		}

		private void OnSelectionChanged()
		{
			if (this.SelectionChanged != null)
			{
				this.SelectionChanged(this, EventArgs.Empty);
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (_measureGraphics != null)
			{
				_measureGraphics.Dispose();
				_measureGraphics = null;
			}
			if (_spellTimer != null)
				{
					_spellTimer.Dispose();
				}
				if (_spellChecker != null)
				{
					_spellChecker.Dispose();
				}
				if (_blinkTimer != null)
				{
					_blinkTimer.Dispose();
				}
				_fonts.Dispose();
			}
			base.Dispose(disposing);
		}
	}
}
