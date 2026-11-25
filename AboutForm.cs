using System.Drawing;
using System.Windows.Forms;

namespace ExcelBan
{
    public sealed class AboutForm : Form
    {
        // این مقادیر را با مشخصات خودتان جایگزین کنید
        private const string AppName = "تولید فایل واریز گروهی بانکی";
        private const string AppVersion = "3.0.0 (C#/.NET)";
        private const string AppAuthor = "علی ابوالحلاج";
        private const string AppRole = "کارشناس ارشد و مشاور فناوری اطلاعات";
        private const string AppDesc = "ابزار تبدیل فایل اکسل واریزهای گروهی به فرمت متنی ثابت (Fixed-Width) بانکی";

        public AboutForm()
        {
            Text = "درباره برنامه";
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = AppFont.Create(9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(380, 260);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                ColumnCount = 1,
                RowCount = 6,
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            layout.Controls.Add(new Label
            {
                Text = AppName,
                Font = AppFont.Create(12F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            }, 0, 0);
            layout.Controls.Add(new Label { Text = "نسخه: " + AppVersion, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
            layout.Controls.Add(new Label { Text = "تهیه‌کننده: " + AppAuthor, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
            layout.Controls.Add(new Label { Text = AppRole, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
            layout.Controls.Add(new Label
            {
                Text = AppDesc,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopRight,
            }, 0, 4);

            var closeButton = new Button { Text = "بستن", Width = 90, Height = 30 };
            closeButton.Click += (_, _) => Close();
            var buttonHost = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
            };
            buttonHost.Controls.Add(closeButton);
            layout.Controls.Add(buttonHost, 0, 5);

            Controls.Add(layout);
        }
    }
}
