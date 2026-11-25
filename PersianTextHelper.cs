using System;
using System.Globalization;
using System.Text;

namespace ExcelBan
{
    /// <summary>
    /// ابزارهای مربوط به متن و اعداد فارسی که رفتارشان صراحتاً در کد مشخص شده
    /// و به تنظیمات locale/codepage فعال ویندوز کاربر وابسته نیست.
    /// </summary>
    public static class PersianTextHelper
    {
        // نگاشت ارقام فارسی و عربی به ارقام انگلیسی (مستقل از CultureInfo سیستم)
        private const string PersianDigits = "۰۱۲۳۴۵۶۷۸۹";
        private const string ArabicDigits = "٠١٢٣٤٥٦٧٨٩";
        private const string AsciiDigits = "0123456789";

        // نویسه‌های فارسی که در codepage ویندوز-1256 (عربی) وجود ندارند.
        // توجه: حروف اختصاصی فارسی پ چ ژ گ در cp1256 موجودند و نیاز به جایگزینی ندارند؛
        // فقط «ی» و «ک» فارسی باید به معادل عربی‌شان تبدیل شوند.
        private static readonly (char From, string To)[] Cp1256SafeMap = new (char, string)[]
        {
            ('ی', "ي"),
            ('ک', "ك"),
            ('\u200c', " "), // نیم‌فاصله (ZWNJ)
            ('\u064b', ""),  // تنوین‌ها/اعراب که در cp1256 مشکل‌سازند
            ('\u064c', ""),
            ('\u064d', ""),
            ('\u064e', ""),
            ('\u064f', ""),
            ('\u0650', ""),
            ('\u0651', ""),
            ('\u0652', ""),
        };

        public static string ToAsciiDigits(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var sb = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                int pIdx = PersianDigits.IndexOf(ch);
                if (pIdx >= 0) { sb.Append(AsciiDigits[pIdx]); continue; }
                int aIdx = ArabicDigits.IndexOf(ch);
                if (aIdx >= 0) { sb.Append(AsciiDigits[aIdx]); continue; }
                sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>حذف جداکننده هزارگان/فاصله جهت آماده‌سازی برای Parse عددی.</summary>
        public static string CleanNumeric(string? text)
        {
            var s = ToAsciiDigits(text).Trim();
            foreach (var ch in new[] { ",", "٬", "،", " " })
                s = s.Replace(ch, "");
            return s;
        }

        private static string ApplyCp1256SafeMap(string text)
        {
            var result = text;
            foreach (var (from, to) in Cp1256SafeMap)
                result = result.Replace(from.ToString(), to);
            return result;
        }

        /// <summary>
        /// تبدیل امن به بایت‌های windows-1256. چون انکودینگ صراحتاً مشخص شده (نه با تکیه بر
        /// codepage فعال سیستم)، روی هر ویندوزی با هر تنظیمات regional یکسان عمل می‌کند.
        /// نویسه‌های ناسازگار باقی‌مانده با '?' جایگزین و تعدادشان برگردانده می‌شود.
        /// </summary>
        public static (byte[] Bytes, int ReplacedCount) SafeCp1256Bytes(string text)
        {
            var normalized = ApplyCp1256SafeMap(text);
            var encoding = Encoding.GetEncoding(
                1256,
                EncoderFallback.ExceptionFallback,
                DecoderFallback.ExceptionFallback);

            var output = new System.Collections.Generic.List<byte>(normalized.Length);
            int replaced = 0;
            foreach (var ch in normalized)
            {
                try
                {
                    output.AddRange(encoding.GetBytes(ch.ToString()));
                }
                catch (EncoderFallbackException)
                {
                    output.Add((byte)'?');
                    replaced++;
                }
            }
            return (output.ToArray(), replaced);
        }

        public static double? ParseAmount(string cleanedNumeric)
        {
            if (string.IsNullOrEmpty(cleanedNumeric)) return null;
            return double.TryParse(
                cleanedNumeric, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                ? value
                : (double?)null;
        }
    }
}
