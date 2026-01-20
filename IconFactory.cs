using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace SmartExcel
{
    internal static class IconFactory
    {
        private static readonly Lazy<Bitmap> splitRowsIcon = new Lazy<Bitmap>(CreateSplitRowsIcon);
        private static readonly Lazy<Bitmap> changeDelimiterIcon = new Lazy<Bitmap>(CreateChangeDelimiterIcon);
        private static readonly Lazy<Bitmap> trimColumnIcon = new Lazy<Bitmap>(CreateTrimColumnIcon);

        public static Bitmap SplitRowsIcon => splitRowsIcon.Value;
        public static Bitmap ChangeDelimiterIcon => changeDelimiterIcon.Value;
        public static Bitmap TrimColumnIcon => trimColumnIcon.Value;

        private static Bitmap CreateSplitRowsIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.Clear(Color.Transparent);

                var outerRect = new Rectangle(4, 5, 24, 22);
                using (var backgroundBrush = new SolidBrush(Color.FromArgb(248, 251, 255)))
                {
                    graphics.FillRectangle(backgroundBrush, outerRect);
                }

                using (var gridPen = new Pen(Color.FromArgb(60, 90, 140), 1.5f))
                {
                    graphics.DrawRectangle(gridPen, outerRect);
                    int rowHeight = outerRect.Height / 4;
                    for (int i = 1; i < 4; i++)
                    {
                        int y = outerRect.Top + (i * rowHeight);
                        graphics.DrawLine(gridPen, outerRect.Left, y, outerRect.Right, y);
                    }
                }

                var highlightedRow = new Rectangle(outerRect.Left + 1, outerRect.Top + 6, outerRect.Width - 2, 6);
                using (var highlightBrush = new SolidBrush(Color.FromArgb(210, 235, 255)))
                {
                    graphics.FillRectangle(highlightBrush, highlightedRow);
                }

                using (var arrowPen = new Pen(Color.FromArgb(30, 120, 70), 2f))
                {
                    arrowPen.CustomEndCap = new AdjustableArrowCap(3, 4);
                    graphics.DrawLine(arrowPen, 10, 26, 10, 30);
                    graphics.DrawLine(arrowPen, 22, 26, 22, 30);
                }
            }

            bitmap.MakeTransparent();
            return bitmap;
        }

        private static Bitmap CreateChangeDelimiterIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                graphics.Clear(Color.Transparent);

                using (var circlePen = new Pen(Color.FromArgb(90, 110, 145), 2f))
                {
                    graphics.DrawEllipse(circlePen, new Rectangle(4, 5, 24, 24));
                }

                using (var textBrush = new SolidBrush(Color.FromArgb(40, 55, 80)))
                using (var font = new Font("Segoe UI", 9f, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    graphics.DrawString(",", font, textBrush, new PointF(8f, 13f));
                    graphics.DrawString(";", font, textBrush, new PointF(20f, 13f));
                }

                using (var arrowPen = new Pen(Color.FromArgb(200, 120, 20), 2.5f))
                {
                    arrowPen.CustomEndCap = new AdjustableArrowCap(4, 6);
                    graphics.DrawLine(arrowPen, 9f, 23f, 23f, 23f);
                }
            }

            bitmap.MakeTransparent();
            return bitmap;
        }

        private static Bitmap CreateTrimColumnIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                graphics.Clear(Color.Transparent);

                var columnRect = new Rectangle(6, 5, 18, 22);
                using (var backgroundBrush = new LinearGradientBrush(columnRect, Color.FromArgb(250, 252, 255), Color.FromArgb(225, 235, 250), LinearGradientMode.Vertical))
                {
                    graphics.FillRectangle(backgroundBrush, columnRect);
                }

                using (var borderPen = new Pen(Color.FromArgb(75, 115, 165), 1.4f))
                {
                    graphics.DrawRectangle(borderPen, columnRect);
                }

                using (var guidePen = new Pen(Color.FromArgb(190, 205, 230), 1f))
                {
                    graphics.DrawLine(guidePen, columnRect.Left + 5, columnRect.Top + 4, columnRect.Left + 5, columnRect.Bottom - 4);
                    graphics.DrawLine(guidePen, columnRect.Left + 9, columnRect.Top + 4, columnRect.Left + 9, columnRect.Bottom - 4);
                }

                using (var textBrush = new SolidBrush(Color.FromArgb(60, 90, 140)))
                using (var font = new Font("Segoe UI", 8f, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    graphics.DrawString("LT", font, textBrush, new PointF(8f, 9f));
                    graphics.DrawString("RT", font, textBrush, new PointF(8f, 18f));
                }

                using (var trimPen = new Pen(Color.FromArgb(70, 160, 85), 2f))
                {
                    trimPen.CustomEndCap = new AdjustableArrowCap(3f, 4f);
                    graphics.DrawLine(trimPen, 25f, 10f, 25f, 24f);
                }

                using (var highlightBrush = new SolidBrush(Color.FromArgb(70, 160, 85)))
                {
                    graphics.FillRectangle(highlightBrush, new Rectangle(23, 13, 5, 2));
                    graphics.FillRectangle(highlightBrush, new Rectangle(23, 20, 5, 2));
                }
            }

            bitmap.MakeTransparent();
            return bitmap;
        }
    }
}
