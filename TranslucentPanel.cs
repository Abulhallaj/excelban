using System.Windows.Forms;

namespace ExcelBan
{
    /// <summary>
    /// Panel‌ای که رنگ پس‌زمینه‌ی نیمه‌شفاف (alpha بین ۱ تا ۲۵۴) آن
    /// واقعاً روی تصویر پس‌زمینه‌ی فرم ترکیب (Blend) می‌شود.
    /// (جزئیات در <see cref="TranslucentBackdrop"/>)
    /// </summary>
    public sealed class TranslucentPanel : Panel
    {
        public TranslucentPanel()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer
                   | ControlStyles.AllPaintingInWmPaint
                   | ControlStyles.UserPaint
                   | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            try
            {
                TranslucentBackdrop.PaintBackground(this, e.Graphics);
            }
            catch
            {
                // در صورت هرگونه مشکل، رفتار پیش‌فرض WinForms
                base.OnPaintBackground(e);
            }
        }
    }
}
