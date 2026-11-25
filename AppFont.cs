using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelBan
{
    /// <summary>
    /// مدیریت یکپارچه‌ی فونت فارسی برنامه.
    /// فایل‌های فونت استاندارد «وزیرمتن» (Vazirmatn) داخل خود برنامه جاسازی (Embed) شده‌اند
    /// تا روی سیستم‌هایی که فونت فارسی نصب ندارند هم به‌درستی نمایش داده شوند.
    /// ترتیب انتخاب: فونت جاسازی‌شده → فونت نصب‌شده‌ی سیستم (Vazirmatn/Saman) → Tahoma.
    /// </summary>
    internal static class AppFont
    {
        private static readonly string[] EmbeddedFontFiles =
        {
            "Vazirmatn-Regular.ttf",
            "Vazirmatn-Bold.ttf",
        };

        /// <summary>فونت‌های نصب‌شده‌ی سیستم که در صورت شکست جاسازی بررسی می‌شوند.</summary>
        private static readonly string[] InstalledFallbacks = { "Vazirmatn", "Saman", "Tahoma" };

        private static readonly object Sync = new();
        private static readonly Dictionary<(float Size, FontStyle Style), Font> Cache = new();

        private static FontFamily? resolvedFamily;
        private static PrivateFontCollection? privateCollection;

        /// <summary>
        /// ساخت (یا بازگرداندن نسخه‌ی کش‌شده‌ی) فونت با اندازه و استایل دلخواه.
        /// نمونه‌های برگشتی کش می‌شوند؛ کنترل‌های ویندوز فرمز Font خود را Dispose نمی‌کنند،
        /// بنابراین نگه‌داشتن آن‌ها برای طول عمر برنامه امن است.
        /// </summary>
        public static Font Create(float size, FontStyle style = FontStyle.Regular)
        {
            lock (Sync)
            {
                var key = (size, style);
                if (!Cache.TryGetValue(key, out var font))
                {
                    font = new Font(GetFamily(), size, style);
                    Cache[key] = font;
                }
                return font;
            }
        }

        private static FontFamily GetFamily()
        {
            if (resolvedFamily != null)
                return resolvedFamily;

            // ۱) تلاش برای بارگذاری فایل‌های جاسازی‌شده در خود برنامه
            resolvedFamily = LoadEmbedded() ?? LoadInstalled();

            if (resolvedFamily == null)
                throw new InvalidOperationException("هیچ فونت فارسی (Vazirmatn/Saman) یافت نشد و Tahoma نیز در دسترس نیست.");

            return resolvedFamily;
        }

        private static FontFamily? LoadEmbedded()
        {
            try
            {
                var assembly = typeof(AppFont).Assembly;
                var resourceNames = assembly.GetManifestResourceNames();

                var collection = new PrivateFontCollection();
                var keepAlive = new List<IntPtr>();

                foreach (var fileName in EmbeddedFontFiles)
                {
                    // نام منبع بر اساس RootNamespace و پوشه متفاوت است؛ با پسوند فایل تطبیق می‌دهیم.
                    var resourceName = resourceNames
                        .FirstOrDefault(n => n.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));
                    if (resourceName == null)
                        continue;

                    using var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream == null)
                        continue;

                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    var data = buffer.ToArray();

                    // حافظه باید تا پایان عمر برنامه معتبر بماند (به‌ازای هر فونت حدود ۱۲۰ کیلوبایت).
                    var ptr = Marshal.AllocHGlobal(data.Length);
                    Marshal.Copy(data, 0, ptr, data.Length);
                    keepAlive.Add(ptr);

                    collection.AddMemoryFont(ptr, data.Length);
                }

                if (collection.Families.Length == 0)
                {
                    collection.Dispose();
                    return null;
                }

                privateCollection = collection; // نگه‌داشتن ارجاع تا جمع‌آوری نشود
                return SelectBestFamily(collection.Families);
            }
            catch
            {
                // در صورت هرگونه مشکل، به گزینه‌های بعدی رجوع می‌کنیم.
                return null;
            }
        }

        private static FontFamily? LoadInstalled()
        {
            try
            {
                using var installed = new InstalledFontCollection();
                var families = installed.Families;

                foreach (var name in InstalledFallbacks)
                {
                    var match = families.FirstOrDefault(f =>
                        string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                        return match;
                }
            }
            catch
            {
                // نادیده گرفتن و ادامه‌ی زنجیره‌ی جایگزین
            }

            return null;
        }

        /// <summary>در میان فونت‌های بارگذاری‌شده، اول وزیرمتن، سپس صمیم و در نهایت اولین فونت موجود را برمی‌گرداند.</summary>
        private static FontFamily SelectBestFamily(FontFamily[] families)
        {
            foreach (var preferred in new[] { "Vazirmatn", "Saman" })
            {
                var match = families.FirstOrDefault(f =>
                    string.Equals(f.Name, preferred, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                    return match;
            }

            return families[0];
        }
    }
}
