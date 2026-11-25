using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelBan
{
    /// <summary>
    /// منطق مشترک «زمینه‌ی نیمه‌شفاف» برای کنترل‌های سفارشی.
    /// در WinForms معمولی، alpha بین ۱ تا ۲۵۴ در BackColor کنترل‌های پنجره‌دار
    /// با تصویر پس‌زمینه‌ی والد ترکیب نمی‌شود (هر کنترل HWND جدا دارد) و عملاً
    /// رنگ تیره/اوپک می‌بینید؛ این کلاس در OnPaintBackground کنترل، ابتدا ناحیه‌ی
    /// مربوطه از تصویر والد را رسم کرده و سپس رنگ BackColor را با GDI+ روی آن ترکیب می‌کند.
    /// والدِ دارای تصویر را در زنجیره‌ی Parent جستجو می‌کند و مختصات را تا آنجا جمع می‌زند؛
    /// اگر به والدِ اوپک بدون تصویر برسد، همان رنگ پایه می‌شود.
    /// </summary>
    internal static class TranslucentBackdrop
    {
        /// <summary>رنگ پس‌زمینه‌ی پیشنهادی بدنه‌ی پنل‌ها: همان خاکستری Control با opacity پنجاه درصد.</summary>
        public static Color HalfTransparentControl => Color.FromArgb(128, 240, 240, 240);

        public static void PaintBackground(Control self, Graphics g)
        {
            // رنگ اوپک خودِ کنترل: ترکیبی لازم نیست، همان رنگ کافی است.
            if (self.BackColor.A == 255)
            {
                using var brush = new SolidBrush(self.BackColor);
                g.FillRectangle(brush, self.ClientRectangle);
                return;
            }

            // صعود از والدها تا یافتن تصویر پس‌زمینه (یا والدِ اوپک که تصویر را می‌پوشاند).
            Control? node = self.Parent;
            int offsetX = self.Left;   // موقعیت همین کنترل در دستگاه مختصات والدِ دارای تصویر
            int offsetY = self.Top;
            Image? image = null;
            var layout = ImageLayout.Zoom;
            var baseColor = SystemColors.Control;
            var ownerClient = Rectangle.Empty;

            while (node != null)
            {
                if (node.BackgroundImage != null)
                {
                    image = node.BackgroundImage;
                    layout = node.BackgroundImageLayout;
                    baseColor = node.BackColor.A == 0 ? SystemColors.Control : node.BackColor;
                    ownerClient = new Rectangle(Point.Empty, node.ClientSize);
                    break;
                }

                if (node.BackColor.A == 255)
                {
                    // والدِ اوپک مانع دیده شدن تصویر بالاتر است؛ زمینه همین‌جا تمام می‌شود.
                    baseColor = node.BackColor;
                    ownerClient = new Rectangle(Point.Empty, node.ClientSize);
                    break;
                }

                offsetX += node.Left;
                offsetY += node.Top;
                node = node.Parent;
            }

            // ۱) رنگ پایه (ناحیه‌ی letterbox اطراف تصویر و حالت بدون تصویر)
            using (var baseBrush = new SolidBrush(baseColor))
                g.FillRectangle(baseBrush, self.ClientRectangle);

            // ۲) تصویر پس‌زمینه‌ی والد — فقط همین ناحیه، با همان layout والد
            if (image != null)
            {
                if (layout == ImageLayout.Tile)
                {
                    using var tileBrush = new TextureBrush(image, WrapMode.Tile);
                    tileBrush.TranslateTransform(-offsetX, -offsetY);
                    g.FillRectangle(tileBrush, self.ClientRectangle);
                }
                else
                {
                    var dest = GetImageRect(image, layout, ownerClient);
                    dest.Offset(-offsetX, -offsetY);
                    g.DrawImage(image, dest);
                }
            }

            // ۳) رنگ نیمه‌شفاف خودِ کنترل روی زمینه
            if (self.BackColor.A != 0)
            {
                using var overlay = new SolidBrush(self.BackColor);
                g.FillRectangle(overlay, self.ClientRectangle);
            }
        }

        /// <summary>محاسبه‌ی مستطیل رسم تصویر مطابق BackgroundImageLayout در مختصات والد.</summary>
        private static Rectangle GetImageRect(Image image, ImageLayout layout, Rectangle bounds)
        {
            switch (layout)
            {
                case ImageLayout.None:
                    return new Rectangle(bounds.X, bounds.Y, image.Width, image.Height);

                case ImageLayout.Center:
                    return new Rectangle(
                        bounds.X + (bounds.Width - image.Width) / 2,
                        bounds.Y + (bounds.Height - image.Height) / 2,
                        image.Width, image.Height);

                case ImageLayout.Stretch:
                    return bounds;

                default: // Zoom (و هر مقدار غیرمنتظره‌ی دیگر)
                    var scale = Math.Min(
                        (double)bounds.Width / image.Width,
                        (double)bounds.Height / image.Height);
                    var w = (int)Math.Round(image.Width * scale);
                    var h = (int)Math.Round(image.Height * scale);
                    return new Rectangle(
                        bounds.X + (bounds.Width - w) / 2,
                        bounds.Y + (bounds.Height - h) / 2,
                        w, h);
            }
        }
    }
}
