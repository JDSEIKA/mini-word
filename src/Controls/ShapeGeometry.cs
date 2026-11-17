using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using WordPadClone.Formats;

namespace WordPadClone.Controls
{
    internal static class ShapeGeometry
    {
        public static GraphicsPath Create(FloatingShapeKind kind, RectangleF rect)
        {
            GraphicsPath path = new GraphicsPath();
            float x = rect.X, y = rect.Y, w = rect.Width, h = rect.Height;
            switch (kind)
            {
                case FloatingShapeKind.Ellipse:
                    path.AddEllipse(rect);
                    break;
                case FloatingShapeKind.Triangle:
                    path.AddPolygon(new PointF[] { new PointF(x + w / 2, y), new PointF(x + w, y + h), new PointF(x, y + h) });
                    break;
                case FloatingShapeKind.Diamond:
                    path.AddPolygon(new PointF[] { new PointF(x + w / 2, y), new PointF(x + w, y + h / 2), new PointF(x + w / 2, y + h), new PointF(x, y + h / 2) });
                    break;
                case FloatingShapeKind.Star:
                    PointF[] star = new PointF[10];
                    for (int i = 0; i < 10; i++)
                    {
                        double angle = -Math.PI / 2 + i * Math.PI / 5;
                        float radius = (i % 2 == 0) ? 0.5f : 0.22f;
                        star[i] = new PointF(x + w / 2 + (float)Math.Cos(angle) * w * radius,
                            y + h / 2 + (float)Math.Sin(angle) * h * radius);
                    }
                    path.AddPolygon(star);
                    break;
                case FloatingShapeKind.Arrow:
                    path.AddPolygon(new PointF[]
                    {
                        new PointF(x, y + h * 0.32f), new PointF(x + w * 0.62f, y + h * 0.32f),
                        new PointF(x + w * 0.62f, y), new PointF(x + w, y + h / 2),
                        new PointF(x + w * 0.62f, y + h), new PointF(x + w * 0.62f, y + h * 0.68f),
                        new PointF(x, y + h * 0.68f)
                    });
                    break;
                case FloatingShapeKind.ArrowLeft:
                    path.AddPolygon(new PointF[]
                    {
                        new PointF(x + w, y + h * 0.32f), new PointF(x + w * 0.38f, y + h * 0.32f),
                        new PointF(x + w * 0.38f, y), new PointF(x, y + h / 2),
                        new PointF(x + w * 0.38f, y + h), new PointF(x + w * 0.38f, y + h * 0.68f),
                        new PointF(x + w, y + h * 0.68f)
                    });
                    break;
                case FloatingShapeKind.Heart:
                    path.StartFigure();
                    path.AddBezier(x + w * 0.5f, y + h * 0.94f,
                        x + w * 0.06f, y + h * 0.62f, x + w * 0.07f, y + h * 0.16f,
                        x + w * 0.27f, y + h * 0.15f);
                    path.AddBezier(x + w * 0.27f, y + h * 0.15f,
                        x + w * 0.39f, y + h * 0.15f, x + w * 0.46f, y + h * 0.27f,
                        x + w * 0.5f, y + h * 0.35f);
                    path.AddBezier(x + w * 0.5f, y + h * 0.35f,
                        x + w * 0.54f, y + h * 0.27f, x + w * 0.61f, y + h * 0.15f,
                        x + w * 0.73f, y + h * 0.15f);
                    path.AddBezier(x + w * 0.73f, y + h * 0.15f,
                        x + w * 0.93f, y + h * 0.16f, x + w * 0.94f, y + h * 0.62f,
                        x + w * 0.5f, y + h * 0.94f);
                    path.CloseFigure();
                    break;
                default:
                    path.AddRectangle(rect);
                    break;
            }
            return path;
        }
    }
}
