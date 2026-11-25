using System;
using System.Text;
using System.Windows.Forms;

namespace ExcelBan
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            // codepage های قدیمی (از جمله 1256) به‌صورت پیش‌فرض در دات‌نت مدرن ثبت نیستند
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
