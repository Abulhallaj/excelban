# Changelog

تمام تغییرات قابل توجه در این پروژه در این فایل ثبت می‌شود.

## [3.0.0] - 2026-11-25

### برنامه **ExcelBan** — نسخه اولیه (اولین انتشار رسمی)

#### امکانات جدید
- بازنویسی کامل برنامه به صورت دسکتاپ ویندوزی با C# و .NET 8 (WinForms).
- تبدیل فایل اکسل ورودی به فرمت متنی ثابت (Fixed-Width) برای واریزهای گروهی بانکی.
- ذخیره خروجی با انکودینگ `windows-1256` (با مدیریت کاراکترهای غیرقابل‌تبدیل).
- تاریخ شمسی (جلالی) داخلی بدون وابستگی خارجی (`JalaliDate.cs`).

#### رابط کاربری
- فونت فارسی «وزیرمتن» (Vazirmatn) به صورت Regular و Bold در پوشه `Fonts` جاسازی شده و به‌عنوان `EmbeddedResource` داخل exe قرار دارد (`AppFont.cs`). مدیریت فونت متمرکز است؛ در صورت شکست بارگذاری از فونت نصب‌شده سیستم و در نهایت Tahoma استفاده می‌شود.
- نمایش پیام‌ها از طریق **RtlMessageDialog** — یک جایگزین RTL-friendly برای `MessageBox` با پشتیبانی از فارسی (منبع: https://github.com/Abulhallaj/RtlMessageDialog، مجوز MIT).
- دکمه‌های دایلوگ محلی‌شده به فارسی: «تأیید»، «انصراف»، «بله»، «خیر».
- پنجره «درباره برنامه» (`AboutForm.cs`) با مشخصات نسخه 3.0.0.
- پنل‌های آکاردئونی (`CollapsiblePanel`) و پنل‌های نیمه‌شفاف (`TranslucentPanel`) برای تجربه کاربری بهتر.

#### فنی
- خواندن اکسل با **ClosedXML** (متن‌باز، MIT) به جای EPPlus (نیازمند مجوز تجاری) و ابزارهای سنگین.
- ثبت codepage قدیمی (از جمله 1256) برای دات‌نت مدرن (`System.Text.Encoding.CodePages`).
- انتشار exe در دو حالت: framework-dependent (حجم کم) و self-contained (بدون نیاز به نصب Runtime).
- مدیریت مقادیر `ContentHeight` روی `CollapsiblePanel` در `MainForm.cs`.

#### مجوزها
- **Vazirmatn**: [SIL Open Font License (OFL)](https://scripts.sil.org/OFL)
- **ClosedXML**: MIT
- **System.Text.Encoding.CodePages**: .NET Foundation (MIT)
- **RtlMessageDialog**: MIT