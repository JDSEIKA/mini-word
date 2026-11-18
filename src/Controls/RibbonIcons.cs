using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace WordPadClone.Controls
{
    /// <summary>
    /// Small vector icons drawn with GDI+, so the ribbon has no external image
    /// dependency and remains sharp at any DPI.
    /// </summary>
    internal static class RibbonIcons
    {
        private static readonly Color Accent = Color.FromArgb(38, 78, 140);
        private static readonly Color Dark = Color.FromArgb(58, 68, 82);
        private static readonly Color Muted = Color.FromArgb(120, 130, 142);
        private static readonly Dictionary<string, Image> Cache = new Dictionary<string, Image>();

        public static Image Get(string name)
        {
            Image image;
            if (Cache.TryGetValue(name, out image)) return image;
            Bitmap bitmap = new Bitmap(32, 32);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                Draw(g, name ?? string.Empty);
            }
            Cache[name] = bitmap;
            return bitmap;
        }

        private static void Draw(Graphics g, string name)
        {
            using (Pen pen = new Pen(Accent, 2f))
            using (Pen dark = new Pen(Dark, 1.5f))
            using (Brush fill = new SolidBrush(Color.FromArgb(235, 242, 251)))
            using (Brush accent = new SolidBrush(Accent))
            {
                switch (name)
                {
                    case "new":
                        g.FillRectangle(fill, 7, 4, 18, 24);
                        g.DrawRectangle(pen, 7, 4, 18, 24);
                        g.DrawLine(pen, 16, 10, 16, 22);
                        g.DrawLine(pen, 11, 16, 21, 16);
                        break;
                    case "open":
                        g.FillPolygon(fill, new Point[] { new Point(4, 11), new Point(12, 11), new Point(15, 8), new Point(27, 8), new Point(27, 25), new Point(4, 25) });
                        g.DrawPolygon(pen, new Point[] { new Point(4, 11), new Point(12, 11), new Point(15, 8), new Point(27, 8), new Point(27, 25), new Point(4, 25) });
                        g.DrawLine(dark, 9, 16, 23, 16);
                        break;
                    case "save":
                        g.FillRectangle(fill, 6, 5, 20, 22);
                        g.DrawRectangle(pen, 6, 5, 20, 22);
                        g.FillRectangle(accent, 10, 5, 12, 7);
                        g.FillRectangle(fill, 10, 17, 12, 8);
                        g.DrawRectangle(dark, 10, 17, 12, 8);
                        break;
                    case "undo":
                    case "redo":
                    {
                        bool redo = name == "redo";
                        using (Pen curve = new Pen(redo ? Accent : Dark, 2.4f))
                        {
                            curve.StartCap = LineCap.Round;
                            curve.EndCap = LineCap.Round;
                            g.DrawArc(curve, 8, 10, 17, 16, redo ? 20 : 160, redo ? 250 : -250);
                        }
                        Point[] head = redo
                            ? new Point[] { new Point(27, 14), new Point(19, 8), new Point(19, 21) }
                            : new Point[] { new Point(5, 14), new Point(13, 8), new Point(13, 21) };
                        g.FillPolygon(accent, head);
                        g.DrawPolygon(dark, head);
                        break;
                    }
                    case "paste":
                        g.FillRectangle(fill, 8, 7, 17, 21);
                        g.DrawRectangle(pen, 8, 7, 17, 21);
                        g.FillRectangle(accent, 12, 4, 9, 7);
                        g.DrawRectangle(dark, 13, 16, 7, 2);
                        g.DrawRectangle(dark, 13, 21, 7, 2);
                        break;
                    case "cut":
                        g.DrawEllipse(pen, 6, 6, 7, 7);
                        g.DrawEllipse(pen, 6, 20, 7, 7);
                        g.DrawLine(pen, 12, 12, 25, 23);
                        g.DrawLine(pen, 12, 20, 25, 8);
                        break;
                    case "copy":
                        g.FillRectangle(fill, 6, 8, 15, 19);
                        g.DrawRectangle(pen, 6, 8, 15, 19);
                        g.FillRectangle(fill, 12, 4, 15, 19);
                        g.DrawRectangle(pen, 12, 4, 15, 19);
                        break;
                    case "bold":
                        DrawLetter(g, "B", FontStyle.Bold, accent);
                        break;
                    case "italic":
                        DrawLetter(g, "I", FontStyle.Italic, accent);
                        break;
                    case "underline":
                        DrawLetter(g, "U", FontStyle.Regular, accent);
                        g.DrawLine(pen, 8, 27, 24, 27);
                        break;
                    case "color":
                        DrawLetter(g, "A", FontStyle.Bold, accent);
                        using (Brush red = new SolidBrush(Color.FromArgb(197, 48, 48)))
                            g.FillRectangle(red, 6, 27, 20, 3);
                        break;
                    case "left":
                    case "center":
                    case "right":
                    case "justify":
                    {
                        for (int i = 0; i < 4; i++)
                        {
                            float y = 7 + i * 5;
                            float x = 6;
                            float w = 20;
                            if (name == "left")
                            {
                                w = (i == 1 || i == 3) ? 14 : 20;
                            }
                            else if (name == "justify")
                            {
                                w = 20;
                            }
                            else if (name == "center")
                            {
                                x = i == 3 ? 12 : 6;
                                w = i == 3 ? 8 : 20;
                            }
                            else if (name == "right")
                            {
                                x = (i == 1 || i == 3) ? 12 : 6;
                                w = (i == 1 || i == 3) ? 14 : 20;
                            }
                            g.DrawLine(i == 0 ? pen : dark, x, y, x + w, y);
                        }
                        break;
                    }
                    case "numbered":
                        using (Font numbers = new Font("Segoe UI", 10f))
                        {
                            g.DrawString("1", numbers, accent, 3, 3);
                            g.DrawString("2", numbers, accent, 3, 13);
                            g.DrawString("3", numbers, accent, 3, 23);
                        }
                        g.DrawLine(pen, 16, 8, 27, 8);
                        g.DrawLine(pen, 16, 18, 27, 18);
                        g.DrawLine(pen, 16, 28, 27, 28);
                        break;
                    case "bullet":
                        for (int i = 0; i < 3; i++)
                        {
                            float y = 8 + i * 7;
                            g.FillEllipse(accent, 6, y - 2, 4, 4);
                            g.DrawLine(i == 0 ? pen : dark, 13, y, 26, y);
                        }
                        break;
                    case "pagebreak":
                        g.DrawRectangle(pen, 5, 4, 12, 16);
                        g.DrawRectangle(pen, 15, 12, 12, 16);
                        using (Pen dash = new Pen(Muted, 1.5f) { DashStyle = DashStyle.Dash })
                            g.DrawLine(dash, 4, 22, 28, 22);
                        break;
                    case "header":
                        g.FillRectangle(fill, 6, 4, 20, 24);
                        g.DrawRectangle(pen, 6, 4, 20, 24);
                        g.FillRectangle(accent, 6, 4, 20, 6);
                        g.DrawLine(dark, 10, 15, 22, 15);
                        g.DrawLine(dark, 10, 20, 22, 20);
                        break;
                    case "page":
                        g.FillRectangle(fill, 6, 4, 20, 24);
                        g.DrawRectangle(pen, 6, 4, 20, 24);
                        g.DrawLine(dark, 10, 11, 22, 11);
                        g.DrawLine(dark, 10, 16, 22, 16);
                        g.DrawLine(dark, 10, 21, 18, 21);
                        break;
                    case "print":
                        g.FillRectangle(fill, 7, 12, 18, 12);
                        g.DrawRectangle(pen, 7, 12, 18, 12);
                        g.FillRectangle(fill, 10, 5, 12, 8);
                        g.DrawRectangle(dark, 10, 5, 12, 8);
                        g.FillRectangle(fill, 10, 21, 12, 7);
                        g.DrawRectangle(dark, 10, 21, 12, 7);
                        g.FillEllipse(accent, 21, 15, 3, 3);
                        break;
                    case "previous":
                    case "next":
                        int dir = name == "previous" ? -1 : 1;
                        g.DrawLine(pen, 16 - dir * 4, 7, 16 - dir * 4 + dir * 10, 16);
                        g.DrawLine(pen, 16 - dir * 4 + dir * 10, 16, 16 - dir * 4, 25);
                        break;
                    case "zoom":
                        g.DrawEllipse(pen, 5, 5, 15, 15);
                        g.DrawLine(pen, 18, 18, 27, 27);
                        g.DrawLine(dark, 10, 12, 16, 12);
                        g.DrawLine(dark, 13, 9, 13, 16);
                        break;
                    case "zoom-in":
                        g.DrawEllipse(pen, 4, 4, 14, 14);
                        g.DrawLine(pen, 16, 16, 27, 27);
                        g.DrawLine(dark, 7, 11, 15, 11);
                        g.DrawLine(dark, 11, 7, 11, 15);
                        break;
                    case "zoom-out":
                        g.DrawEllipse(pen, 4, 4, 14, 14);
                        g.DrawLine(pen, 16, 16, 27, 27);
                        g.DrawLine(dark, 7, 11, 15, 11);
                        break;
                    case "zoom-reset":
                        g.DrawRectangle(pen, 5, 5, 20, 22);
                        using (Font percent = new Font("Segoe UI", 7f, FontStyle.Bold))
                            g.DrawString("1:1", percent, accent, 8, 11);
                        break;
                    case "find":
                        g.DrawEllipse(pen, 5, 5, 14, 14);
                        g.DrawLine(pen, 17, 17, 27, 27);
                        g.DrawLine(dark, 9, 12, 15, 12);
                        break;
                    case "merge":
                        g.DrawRectangle(pen, 5, 7, 22, 20);
                        g.DrawLine(dark, 16, 7, 16, 27);
                        g.DrawLine(pen, 10, 17, 22, 17);
                        g.DrawLine(pen, 10, 17, 13, 14);
                        g.DrawLine(pen, 10, 17, 13, 20);
                        g.DrawLine(pen, 22, 17, 19, 14);
                        g.DrawLine(pen, 22, 17, 19, 20);
                        break;
                    case "split":
                        g.DrawRectangle(pen, 5, 7, 22, 20);
                        g.DrawLine(dark, 16, 7, 16, 27);
                        g.DrawLine(pen, 16, 17, 8, 17);
                        g.DrawLine(pen, 16, 17, 20, 13);
                        g.DrawLine(pen, 16, 17, 20, 21);
                        g.DrawLine(pen, 16, 17, 24, 13);
                        g.DrawLine(pen, 16, 17, 24, 21);
                        break;
                    case "ruler":
                        g.FillRectangle(fill, 3, 10, 26, 12);
                        g.DrawRectangle(pen, 3, 10, 26, 12);
                        g.DrawLine(pen, 3, 15, 29, 15);
                        for (int i = 0; i < 6; i++)
                            g.DrawLine(dark, 6 + i * 4, 10, 6 + i * 4, i % 2 == 0 ? 15 : 13);
                        break;
                    case "indent":
                        g.DrawLine(pen, 5, 8, 5, 25);
                        for (int i = 0; i < 4; i++) g.DrawLine(pen, 12, 8 + i * 5, 27, 8 + i * 5);
                        // Increase indent: the arrow points right, toward the text.
                        g.FillPolygon(accent, new Point[] { new Point(7, 12), new Point(7, 20), new Point(12, 16) });
                        break;
                    case "outdent":
                        g.DrawLine(pen, 5, 8, 5, 25);
                        for (int i = 0; i < 4; i++) g.DrawLine(pen, 12, 8 + i * 5, 27, 8 + i * 5);
                        // Decrease indent: the arrow points left, away from the text.
                        g.FillPolygon(accent, new Point[] { new Point(12, 12), new Point(12, 20), new Point(7, 16) });
                        break;
                    case "addrow":
                        g.DrawRectangle(pen, 7, 4, 18, 9);
                        g.DrawRectangle(pen, 7, 16, 18, 9);
                        g.DrawLine(pen, 16, 12, 16, 24);
                        g.DrawLine(pen, 12, 18, 20, 18);
                        break;
                    case "removerow":
                        g.DrawRectangle(pen, 7, 4, 18, 9);
                        g.DrawRectangle(pen, 7, 16, 18, 9);
                        g.DrawLine(pen, 12, 18, 20, 18);
                        break;
                    case "addcolumn":
                        g.DrawRectangle(pen, 4, 7, 10, 18);
                        g.DrawRectangle(pen, 16, 7, 10, 18);
                        g.DrawLine(pen, 15, 4, 15, 28);
                        g.DrawLine(pen, 8, 16, 22, 16);
                        break;
                    case "removecolumn":
                        g.DrawRectangle(pen, 4, 7, 10, 18);
                        g.DrawRectangle(pen, 16, 7, 10, 18);
                        g.DrawLine(pen, 8, 16, 22, 16);
                        break;
                    case "shapecolor":
                        g.FillRectangle(fill, 5, 9, 15, 15);
                        g.DrawRectangle(pen, 5, 9, 15, 15);
                        g.FillEllipse(accent, 17, 3, 9, 9);
                        g.DrawEllipse(dark, 17, 3, 9, 9);
                        g.FillEllipse(accent, 19, 19, 9, 9);
                        g.DrawEllipse(dark, 19, 19, 9, 9);
                        break;
                    case "paragraph":
                        g.DrawLine(pen, 7, 7, 25, 7);
                        g.DrawLine(pen, 7, 13, 25, 13);
                        g.DrawLine(pen, 7, 19, 25, 19);
                        g.DrawLine(pen, 7, 25, 18, 25);
                        g.FillEllipse(accent, 21, 21, 7, 7);
                        break;
                    case "columns":
                        g.DrawRectangle(pen, 6, 6, 8, 21);
                        g.DrawRectangle(pen, 16, 6, 8, 21);
                        g.DrawLine(dark, 14, 4, 14, 29);
                        break;
                    case "section":
                        g.DrawRectangle(pen, 5, 7, 22, 20);
                        g.DrawLine(pen, 5, 14, 27, 14);
                        g.DrawLine(pen, 11, 5, 11, 9);
                        g.DrawLine(pen, 21, 5, 21, 9);
                        break;
                    case "spell":
                        g.DrawRectangle(pen, 5, 7, 22, 19);
                        g.DrawLine(dark, 9, 12, 23, 12);
                        g.DrawLine(dark, 9, 17, 19, 17);
                        g.DrawLine(pen, 18, 8, 26, 20);
                        g.DrawLine(pen, 22, 14, 28, 14);
                        break;
                    case "help":
                        g.DrawEllipse(pen, 5, 5, 22, 22);
                        using (Pen question = new Pen(Dark, 2.4f))
                        {
                            question.StartCap = LineCap.Round;
                            question.EndCap = LineCap.Round;
                            g.DrawArc(question, 11, 8, 10, 10, 190, 220);
                            g.DrawLine(question, 16, 18, 16, 20);
                        }
                        using (Brush dot = new SolidBrush(Dark))
                            g.FillEllipse(dot, 13.7f, 23.2f, 4.6f, 4.6f);
                        break;
                    case "info":
                        g.DrawEllipse(pen, 5, 5, 22, 22);
                        using (Pen info = new Pen(Dark, 2.4f))
                        {
                            info.StartCap = LineCap.Round;
                            info.EndCap = LineCap.Round;
                            g.DrawLine(info, 16, 14, 16, 24);
                        }
                        using (Brush dot = new SolidBrush(Dark))
                            g.FillEllipse(dot, 13.8f, 8.2f, 4.4f, 4.4f);
                        break;
                    case "comment":
                        g.FillRectangle(fill, 4, 6, 22, 16);
                        g.DrawRectangle(pen, 4, 6, 22, 16);
                        g.DrawLine(dark, 8, 11, 22, 11);
                        g.DrawLine(dark, 8, 15, 19, 15);
                        g.DrawLine(pen, 8, 22, 5, 27);
                        break;
                    case "textbox":
                        g.DrawRectangle(pen, 5, 8, 22, 17);
                        g.DrawLine(dark, 9, 13, 23, 13);
                        g.DrawLine(dark, 9, 17, 19, 17);
                        g.DrawLine(pen, 5, 5, 27, 5);
                        g.DrawLine(pen, 5, 5, 5, 12);
                        break;
                    case "shape":
                        g.FillRectangle(fill, 5, 9, 22, 16);
                        g.DrawRectangle(pen, 5, 9, 22, 16);
                        g.FillEllipse(accent, 19, 4, 8, 8);
                        break;
                    case "wrap":
                        g.DrawRectangle(pen, 6, 9, 10, 14);
                        g.DrawLine(dark, 20, 10, 20, 22);
                        g.DrawLine(dark, 20, 10, 17, 13);
                        g.DrawLine(dark, 20, 10, 23, 13);
                        break;
                    case "calendar":
                        g.FillRectangle(fill, 5, 8, 22, 20);
                        g.DrawRectangle(pen, 5, 8, 22, 20);
                        g.FillRectangle(accent, 5, 8, 22, 5);
                        g.DrawLine(dark, 10, 5, 10, 10);
                        g.DrawLine(dark, 22, 5, 22, 10);
                        g.DrawLine(dark, 9, 17, 23, 17);
                        g.DrawLine(dark, 9, 22, 23, 22);
                        break;
                    case "table":
                        g.DrawRectangle(pen, 5, 6, 22, 20);
                        g.DrawLine(dark, 5, 12, 27, 12);
                        g.DrawLine(dark, 5, 18, 27, 18);
                        g.DrawLine(dark, 11, 6, 11, 26);
                        g.DrawLine(dark, 18, 6, 18, 26);
                        break;
                    case "image":
                        g.FillRectangle(fill, 5, 7, 22, 18);
                        g.DrawRectangle(pen, 5, 7, 22, 18);
                        g.FillEllipse(accent, 20, 10, 4, 4);
                        g.FillPolygon(accent, new Point[] { new Point(7, 23), new Point(14, 15), new Point(19, 20), new Point(22, 17), new Point(25, 23) });
                        break;
                    case "setup":
                        g.DrawRectangle(pen, 7, 5, 18, 22);
                        g.DrawLine(dark, 11, 11, 21, 11);
                        g.DrawLine(dark, 11, 16, 21, 16);
                        using (Pen ruler = new Pen(Muted, 1f))
                        {
                            g.DrawLine(ruler, 3, 5, 3, 27);
                            for (int i = 0; i < 5; i++) g.DrawLine(ruler, 1, 7 + i * 5, 4, 7 + i * 5);
                        }
                        break;
                    case "close":
                        g.DrawLine(pen, 8, 8, 24, 24);
                        g.DrawLine(pen, 24, 8, 8, 24);
                        break;
                    case "wordcount":
                        g.DrawString("Aa", new Font("Segoe UI", 10f, FontStyle.Bold), accent, 2, 8);
                        break;
                    case "link":
                        g.DrawArc(pen, 5, 11, 13, 10, 40, 220);
                        g.DrawArc(pen, 14, 11, 13, 10, 220, 220);
                        g.DrawLine(dark, 12, 16, 20, 16);
                        break;
                    default:
                        g.DrawEllipse(pen, 7, 7, 18, 18);
                        g.DrawLine(dark, 16, 11, 16, 21);
                        g.DrawLine(dark, 11, 16, 21, 16);
                        break;
                }
            }
        }

        private static void DrawLetter(Graphics g, string text, FontStyle style, Brush brush)
        {
            using (Font font = new Font("Segoe UI", 22f, style, GraphicsUnit.Pixel))
                g.DrawString(text, font, brush, 1, 1);
        }
    }
}
