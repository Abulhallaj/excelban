using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelBan
{
    /// <summary>
    /// بخش آکاردئونی: یک سربرگ قابل کلیک + یک بدنه‌ی قابل نمایش/مخفی‌سازی.
    /// برای قرار گرفتن داخل FlowLayoutPanel با جهت عمودی طراحی شده تا با
    /// تغییر ارتفاع (باز/بسته شدن)، بخش‌های بعدی به‌طور خودکار جابه‌جا شوند.
    /// </summary>
    public sealed class CollapsiblePanel : Panel
    {
        private const int HeaderHeight = 34;

        private readonly Button toggleButton;
        private readonly Label titleLabel;
        private readonly Label statusLabel;

        public Panel Body { get; }
        public bool Expanded { get; private set; }
        private int contentHeight = 150;

        /// <summary>
        /// ارتفاع محتوای بدنه (بدون احتساب سربرگ). مقداردهی، ارتفاع Body را هم همگام می‌کند —
        /// قبلاً Body همیشه با مقدار پیش‌فرض ۱۵۰ ساخته می‌شد و با ContentHeight واقعی ناهماهنگ می‌ماند.
        /// </summary>
        public int ContentHeight
        {
            get => contentHeight;
            set
            {
                contentHeight = value;
                Body.Height = value;
                if (Expanded) ApplyState();
            }
        }

        public event EventHandler? ExpandedChanged;

        public CollapsiblePanel(string title)
        {
            BorderStyle = BorderStyle.FixedSingle;
            Margin = new Padding(0, 0, 0, 6);
            Height = HeaderHeight;

            // خودِ کنترل شفاف باشد تا ناحیه‌ی بین سربرگ/بدنه و حاشیه‌ها،
            // تصویر پس‌زمینه‌ی فرم را نشان دهد (رنگ بدنه روی بدنه‌ی نیمه‌شفاف کشیده می‌شود).
            BackColor = Color.Transparent;

            var headerTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = HeaderHeight,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.Gainsboro,
            };
            headerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 36));
            headerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            headerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            headerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            toggleButton = new Button
            {
                Text = "+",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
            };
            toggleButton.Click += (_, _) => Toggle();

            titleLabel = new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = AppFont.Create(9.5F, FontStyle.Bold),
                Padding = new Padding(0, 0, 6, 0),
            };

            statusLabel = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(26, 122, 26),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = AppFont.Create(10F, FontStyle.Bold),
            };

            headerTable.Controls.Add(toggleButton, 0, 0);
            headerTable.Controls.Add(titleLabel, 1, 0);
            headerTable.Controls.Add(statusLabel, 2, 0);

            Body = new TranslucentPanel
            {
                Left = 0,
                Top = HeaderHeight,
                Width = Width,
                Height = ContentHeight,
                Visible = false,
                AutoScroll = true,
                Padding = new Padding(12, 8, 12, 8),
                // خاکستری همیشگی بدنه، این‌بار با opacity پنجاه درصد تا تصویر پس‌زمینه دیده شود
                BackColor = TranslucentBackdrop.HalfTransparentControl,
            };

            Controls.Add(Body);
            Controls.Add(headerTable);

            SizeChanged += (_, _) => { Body.Width = Width; };
        }

        public void Toggle()
        {
            Expanded = !Expanded;
            ApplyState();
        }

        public void Expand()
        {
            if (!Expanded)
            {
                Expanded = true;
                ApplyState();
            }
        }

        public void SetStatus(string text) => statusLabel.Text = text;

        /// <summary>
        /// ارتفاع ContentHeight را از روی محتوای واقعیِ فرزندانِ Body محاسبه و اعمال می‌کند؛
        /// برای کنترل‌های Dock=Fill از PreferredSize (ارتفاع طبیعی مورد نیاز) استفاده می‌شود.
        /// ۴ پیکسل نفس اضافه تا اختلافِ گرد کردن باعث ظاهر شدن اسکرول‌بار نشود.
        /// </summary>
        public void FitToContent()
        {
            int needed = 0;
            // توجه: از Visible استفاده نکن — وقتی Body مخفی است، فرزندان هم Visible=false گزارش می‌دهند
            foreach (Control c in Body.Controls)
            {
                bool stretchesVertically = c.Dock == DockStyle.Top || c.Dock == DockStyle.Bottom || c.Dock == DockStyle.Fill;
                int h = stretchesVertically ? c.PreferredSize.Height : c.Bottom;
                needed = Math.Max(needed, h);
            }
            ContentHeight = Math.Max(40, needed + Body.Padding.Vertical + 4);
        }

        private void ApplyState()
        {
            Body.Visible = Expanded;
            toggleButton.Text = Expanded ? "-" : "+";
            Body.Height = ContentHeight; // همگام‌سازی بدنه با ارتفاع محتوا
            Height = HeaderHeight + (Expanded ? ContentHeight : 0);
            ExpandedChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
