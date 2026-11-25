using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using RtlMessageDialog;
using ClosedXML.Excel;

namespace ExcelBan
{
    public sealed class MainForm : Form
    {
        private const string AppTitle = "تولید فایل واریز گروهی بانکی";

        // ---- state ----
        private string? excelPath;
        private XLWorkbook? workbook;
        private IXLWorksheet? currentSheet;
        private List<string> sheetNames = new();
        private List<string[]> previewRows = new();
        private List<string> headerValues = new();
        private int maxRow = 1;
        private int maxCol = 1;
        private readonly Dictionary<string, int> autoDetected = new();
        private readonly Dictionary<string, ComboBox> mappingCombos = new();
        private ToolStripMenuItem? fileCloseItem;
        private Panel busyOverlay = null!;
        private ProgressBar busyProgress = null!;
        private bool isBusy;

        // ---- انتخاب ستون‌ها (پیش‌نمایش + نگاشت) ----
        /// <summary>پالت رنگ‌ها؛ هر ستون اکسل بر اساس شماره‌اش یک رنگ ثابت می‌گیرد.</summary>
        private static readonly Color[] ColumnPalette =
        {
            Color.FromArgb(255, 214, 160), // نارنجی روشن
            Color.FromArgb(160, 220, 180), // سبز
            Color.FromArgb(165, 190, 255), // آبی
            Color.FromArgb(255, 170, 190), // صورتی
            Color.FromArgb(200, 170, 255), // بنفش
            Color.FromArgb(160, 230, 230), // فیروزه‌ای
            Color.FromArgb(255, 240, 150), // زرد
            Color.FromArgb(215, 210, 160), // زیتونی
            Color.FromArgb(255, 195, 155), // هلویی
            Color.FromArgb(180, 210, 225), // آبی‌خاکستری
        };

        /// <summary>هنگام بازسازی کمبوباکس‌ها روشن است تا رویدادها حلقه نشوند.</summary>
        private bool suppressMappingEvents;

        // ---- controls ----
        private FlowLayoutPanel accordionHost = null!;
        private CollapsiblePanel sec1 = null!, sec2 = null!, sec3 = null!, sec4 = null!, sec5 = null!;
        private Panel bottomPanel = null!;
        private Button generateButton = null!;
        private Label statusLabel = null!;
        private Label fileLabel = null!;
        private ComboBox sheetCombo = null!;
        private CheckBox headerCheck = null!;
        private NumericUpDown startRowSpin = null!;
        private NumericUpDown endRowSpin = null!;
        private Label autoDetectLabel = null!;
        private DataGridView previewGrid = null!;

        public MainForm()
        {
            Text = AppTitle;
            RightToLeft = RightToLeft.Yes;
            Font = AppFont.Create(9F);
            ClientSize = new Size(820, 720);
            MinimumSize = new Size(680, 560);
            StartPosition = FormStartPosition.CenterScreen;

            BackgroundImage = LoadBackground();
            BackgroundImageLayout = ImageLayout.Zoom;

            BuildLayout();
            BuildMenu();
            BuildBusyOverlay();
            Resize += (_, _) => busyOverlay.Bounds = ClientRectangle;

            Load += MainForm_Load;
        }

        private void MainForm_Load(object? sender, EventArgs e)
        {
            ResizeSections();
            RepositionBottomBar();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            workbook?.Dispose();
            base.OnFormClosed(e);
        }

        // ================= منو =================
        private void BuildMenu()
        {
            var menu = new MenuStrip();

            // ---- فایل ----
            var fileMenu = new ToolStripMenuItem("فایل");
            var openItem = new ToolStripMenuItem("بازکردن...");
            openItem.ShortcutKeys = Keys.Control | Keys.O;
            openItem.Click += (_, _) => OnPickFile();

            fileCloseItem = new ToolStripMenuItem("بستن");
            fileCloseItem.Click += (_, _) => CloseCurrentFile();
            fileCloseItem.Enabled = false;

            var exitItem = new ToolStripMenuItem("خروج");
            exitItem.Click += (_, _) => Close();

            fileMenu.DropDownItems.Add(openItem);
            fileMenu.DropDownItems.Add(fileCloseItem);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(exitItem);

            // ---- راهنما ----
            var helpMenu = new ToolStripMenuItem("راهنما");
            var aboutItem = new ToolStripMenuItem("درباره برنامه");
            aboutItem.Click += (_, _) =>
            {
                using var about = new AboutForm();
                about.ShowDialog(this);
            };
            helpMenu.DropDownItems.Add(aboutItem);

            menu.Items.Add(fileMenu);
            menu.Items.Add(helpMenu);

            MainMenuStrip = menu;
            Controls.Add(menu);
        }

        // ================= لودینگ و پس‌زمینه =================
        private static Image? LoadBackground()
        {
            try
            {
                var asm = typeof(MainForm).Assembly;
                // نام منبع بر اساس RootNamespace و پوشه متفاوت است؛ با پسوند فایل تطبیق می‌دهیم.
                var resourceName = asm.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("mellat_bg.jpg", StringComparison.OrdinalIgnoreCase));
                if (resourceName == null) return null;

                using var stream = asm.GetManifestResourceStream(resourceName);
                if (stream == null) return null;
                using var img = Image.FromStream(stream);
                return new Bitmap(img); // کپی نهایی تا بستن stream مشکلی ایجاد نکند (الزام GDI+)
            }
            catch
            {
                return null;
            }
        }

        private void BuildBusyOverlay()
        {
            busyOverlay = new Panel
            {
                BackColor = Color.FromArgb(242, 242, 242),
                Visible = false,
                Bounds = ClientRectangle,
            };

            var busyLabel = new Label
            {
                Text = "در حال پردازش فایل...",
                AutoSize = true,
                Font = AppFont.Create(11F, FontStyle.Bold),
                ForeColor = Color.FromArgb(80, 80, 80),
            };

            busyProgress = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Width = 220,
                Height = 18,
            };

            busyOverlay.Controls.Add(busyLabel);
            busyOverlay.Controls.Add(busyProgress);
            busyOverlay.Resize += (_, _) => CenterBusyIndicator();

            Controls.Add(busyOverlay); // آخرین کنترل → بالاترین z-order (روی منو هم پوشش داده شود)
        }

        private void CenterBusyIndicator()
        {
            if (busyOverlay.Controls.Count < 2) return;
            int cx = busyOverlay.ClientSize.Width / 2;
            int cy = busyOverlay.ClientSize.Height / 2;
            var lbl = busyOverlay.Controls[0];
            var prg = busyOverlay.Controls[1];
            lbl.Location = new Point(cx - lbl.Width / 2, cy - 26);
            prg.Location = new Point(cx - prg.Width / 2, cy + 6);
        }

        private void SetBusy(bool busy)
        {
            isBusy = busy;
            if (busy)
            {
                busyOverlay.Bounds = ClientRectangle;
                CenterBusyIndicator();
                busyOverlay.BringToFront();
                busyOverlay.Visible = true;
                busyOverlay.Refresh(); // رسم کامل قبل از شروع کار سنگین
            }
            else
            {
                busyOverlay.Visible = false;
            }
            UseWaitCursor = busy;
        }

        // ================= چیدمان کلی =================
        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                ColumnCount = 1,
                RowCount = 2,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));

            // میزبان نیمه‌شفاف: ترکیب واقعی alpha با تصویر پس‌زمینه (پیش‌فرض WinForms این را blend نمی‌کند)
            accordionHost = new TranslucentFlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0, 0, 0, 50),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(8),
            };
            accordionHost.Resize += (_, _) => ResizeSections();

            sec1 = new CollapsiblePanel("مرحله ۱ - انتخاب فایل اکسل") { ContentHeight = 50 };
            BuildStep1(sec1.Body);
            accordionHost.Controls.Add(sec1);
            sec1.Expand();

            sec2 = new CollapsiblePanel("مرحله ۲ - انتخاب شیت") { ContentHeight = 46 };
            BuildStep2(sec2.Body);
            accordionHost.Controls.Add(sec2);

            sec3 = new CollapsiblePanel("مرحله ۳ - سرستون‌ها و محدوده داده") { ContentHeight = 170 };
            BuildStep3(sec3.Body);
            accordionHost.Controls.Add(sec3);

            sec4 = new CollapsiblePanel("مرحله ۴ - پیش‌نمایش ۵ سطر اول") { ContentHeight = 220 };
            BuildStep4(sec4.Body);
            accordionHost.Controls.Add(sec4);

            sec5 = new CollapsiblePanel("مرحله ۵ - اتصال ستون‌ها") { ContentHeight = 210 };
            BuildStep5(sec5.Body);
            accordionHost.Controls.Add(sec5);

            bottomPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            generateButton = new Button
            {
                Text = "تولید فایل خروجی",
                Width = 170,
                Height = 34,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };
            generateButton.Click += (_, _) => OnGenerateClicked();

            statusLabel = new Label
            {
                Text = "ابتدا یک فایل اکسل انتخاب کنید.",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
            };
            bottomPanel.Controls.Add(generateButton);
            bottomPanel.Controls.Add(statusLabel);
            bottomPanel.Resize += (_, _) => RepositionBottomBar();

            root.Controls.Add(accordionHost, 0, 0);
            root.Controls.Add(bottomPanel, 0, 1);

            Controls.Add(root);
        }

        private void ResizeSections()
        {
            int width = Math.Max(320, accordionHost.ClientSize.Width - 24);
            foreach (Control c in accordionHost.Controls)
            {
                if (c is CollapsiblePanel cp) cp.Width = width;
            }
        }

        private void RepositionBottomBar()
        {
            generateButton.Left = Math.Max(10, bottomPanel.ClientSize.Width - generateButton.Width - 14);
            generateButton.Top = Math.Max(4, (bottomPanel.ClientSize.Height - generateButton.Height) / 2);
            statusLabel.Left = 14;
            statusLabel.Top = (bottomPanel.ClientSize.Height - statusLabel.Height) / 2;
        }

        // ================= مرحله ۱ =================
        private void BuildStep1(Panel body)
        {
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var pickButton = new Button { Text = "انتخاب فایل اکسل...", AutoSize = true, Padding = new Padding(6, 3, 6, 3) };
            pickButton.Click += (_, _) => OnPickFile();

            fileLabel = new Label
            {
                Text = "فایلی انتخاب نشده است.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.DimGray,
            };

            t.Controls.Add(pickButton, 0, 0);
            t.Controls.Add(fileLabel, 1, 0);
            body.Controls.Add(t);
        }

        private async void OnPickFile()
        {
            if (isBusy) return;

            using var dlg = new OpenFileDialog
            {
                Title = "انتخاب فایل اکسل",
                Filter = "Excel Files|*.xlsx;*.xlsm|All Files|*.*",
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            string path = dlg.FileName;
            SetBusy(true);
            try
            {
                XLWorkbook newWorkbook;
                try
                {
                    newWorkbook = await Task.Run(() => new XLWorkbook(path));
                }
                catch (Exception ex)
                {
                    MessageForm.ShowError(this, "خطا در باز کردن فایل", ex.Message, details: ex.ToString());
                    return;
                }

                if (IsDisposed) { newWorkbook.Dispose(); return; }

                workbook?.Dispose();
                workbook = newWorkbook;
                excelPath = path;
                fileLabel.Text = Path.GetFileName(excelPath);
                sec1.SetStatus("✓");
                UpdateFileMenuState();

                sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
                sheetCombo.Items.Clear();
                if (sheetNames.Count > 0)
                {
                    sheetCombo.Items.AddRange(sheetNames.Cast<object>().ToArray());
                    sheetCombo.SelectedIndex = 0;
                }

                sec2.Expand();
                statusLabel.Text = "فایل بارگذاری شد. شیت مورد نظر را بررسی کنید.";
                ResizeSections();
            }
            finally
            {
                SetBusy(false);
            }
        }

        // ================= مرحله ۲ =================
        private void BuildStep2(Panel body)
        {
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var caption = new Label { Text = "شیت مورد نظر:", AutoSize = true, RightToLeft = RightToLeft.Yes, TextAlign = ContentAlignment.MiddleRight };

            sheetCombo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            sheetCombo.SelectedIndexChanged += (_, _) => OnSheetSelected();

            t.Controls.Add(caption, 0, 0);
            t.Controls.Add(sheetCombo, 1, 0);
            body.Controls.Add(t);
        }

        private void CloseCurrentFile()
        {
            if (workbook == null) return;

            workbook.Dispose();
            workbook = null;
            currentSheet = null;
            excelPath = null;
            sheetNames.Clear();
            previewRows.Clear();
            headerValues.Clear();
            autoDetected.Clear();

            fileLabel.Text = "فایلی انتخاب نشده است.";
            sec1.SetStatus("");
            sec2.SetStatus("");
            sec3.SetStatus("");
            sec4.SetStatus("");
            sec5.SetStatus("");

            sheetCombo.Items.Clear();
            maxRow = 1;
            maxCol = 1;
            headerCheck.Checked = true;
            startRowSpin.Minimum = 1;
            startRowSpin.Maximum = 1;
            startRowSpin.Value = 1;
            endRowSpin.Minimum = 1;
            endRowSpin.Maximum = 1;
            endRowSpin.Value = 1;
            autoDetectLabel.Text = "";
            sec3.FitToContent(); // متن قرمز/سبز خالی شد؛ ارتفاع مرحله ۳ را دوباره تنظیم کن

            previewGrid.Columns.Clear();
            previewGrid.Rows.Clear();

            foreach (var combo in mappingCombos.Values)
            {
                combo.Items.Clear();
                combo.SelectedIndex = -1;
            }

            if (sec2.Expanded) sec2.Toggle();
            if (sec3.Expanded) sec3.Toggle();
            if (sec4.Expanded) sec4.Toggle();
            if (sec5.Expanded) sec5.Toggle();

            statusLabel.Text = "فایل بسته شد. یک فایل اکسل انتخاب کنید.";
            UpdateFileMenuState();
        }

        private void UpdateFileMenuState()
        {
            if (fileCloseItem != null) fileCloseItem.Enabled = workbook != null;
        }

        private void OnSheetSelected()
        {
            if (workbook == null || sheetCombo.SelectedItem == null) return;

            var sheetName = sheetCombo.SelectedItem.ToString()!;
            currentSheet = workbook.Worksheet(sheetName);

            var lastRow = currentSheet.LastRowUsed();
            var lastCol = currentSheet.LastColumnUsed();
            maxRow = lastRow?.RowNumber() ?? 1;
            maxCol = lastCol?.ColumnNumber() ?? 1;

            startRowSpin.Maximum = maxRow;
            endRowSpin.Maximum = maxRow;

            sec2.SetStatus("✓");
            LoadPreviewRows();
            OnHeaderToggle();
            sec3.Expand();
            sec4.Expand();
            sec5.Expand();
            ResizeSections();
        }

        private void LoadPreviewRows()
        {
            previewRows.Clear();
            var rowNumbers = new List<int>(); // شمارهٔ سطر واقعی در شیت (برای ستون «ردیف»)
            int rowsToRead = Math.Min(6, maxRow);
            for (int r = 1; r <= rowsToRead; r++)
            {
                var values = new string[maxCol];
                for (int c = 1; c <= maxCol; c++)
                {
                    var cell = currentSheet!.Cell(r, c);
                    values[c - 1] = cell.IsEmpty() ? "" : cell.GetString();
                }
                previewRows.Add(values);
                rowNumbers.Add(r);
            }

            previewGrid.Columns.Clear();
            previewGrid.Rows.Clear();

            // ستون شمارهٔ ردیف — در RTL سمت راست‌ترین ستون است؛ غیرقابل مرتب‌سازی
            previewGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "__rowNo",
                HeaderText = "ردیف",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = true,
                FillWeight = 30,
                MinimumWidth = 35,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter },
            });

            for (int c = 1; c <= maxCol; c++)
            {
                previewGrid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "col" + c,
                    HeaderText = GetColumnLetter(c),
                    SortMode = DataGridViewColumnSortMode.NotSortable, // جلوگیری از مرتب‌سازی با کلیک روی هدر
                    MinimumWidth = 55,
                });
            }

            for (int i = 0; i < previewRows.Count; i++)
            {
                var src = previewRows[i];
                var line = new string[src.Length + 1];
                line[0] = rowNumbers[i].ToString();
                Array.Copy(src, 0, line, 1, src.Length);
                previewGrid.Rows.Add(line);
            }

            // ارتفاع مرحله ۴ از روی محتوای واقعی شبکه (هدر + همهٔ سطرها) تا هیچ سطری بریده نشود
            previewGrid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);
            int rowsHeight = 0;
            foreach (DataGridViewRow row in previewGrid.Rows) rowsHeight += row.Height;
            sec4.ContentHeight = Math.Max(120,
                previewGrid.ColumnHeadersHeight + rowsHeight + previewGrid.Padding.Vertical
                + sec4.Body.Padding.Vertical + 4);
        }

        // ================= مرحله ۳ =================
        private void BuildStep3(Panel body)
        {
            // چیدمان RTL: هر ستون کپشن AutoSize است تا NumericUpDown آن بلافاصله به متنِ سمت راستش بچسبد؛
            // ستون آخر (پرکننده) فضای باقی‌مانده در سمت چپ را می‌گیرد.
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 3 };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // startCaption (راست‌ترین)
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // startRowSpin
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // endCaption
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // endRowSpin
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // پرکننده (چپ‌ترین)
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            headerCheck = new CheckBox
            {
                Text = "اولین سطر به‌عنوان عنوان ستون‌ها (Header) در نظر گرفته شود",
                AutoSize = true,
                CheckAlign = ContentAlignment.MiddleLeft,
                TextAlign = ContentAlignment.MiddleRight,
                Checked = true,
            };
            headerCheck.CheckedChanged += (_, _) => OnHeaderToggle();
            t.Controls.Add(headerCheck, 0, 0);
            t.SetColumnSpan(headerCheck, 5);

            var startCaption = new Label { Text = "سطر شروع داده:", AutoSize = true, TextAlign = ContentAlignment.MiddleRight, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            startRowSpin = new NumericUpDown { Minimum = 1, Maximum = 1, Value = 1, Dock = DockStyle.Fill };
            var endCaption = new Label { Text = "سطر پایان داده:", AutoSize = true, TextAlign = ContentAlignment.MiddleRight, Anchor = AnchorStyles.Left | AnchorStyles.Right };
            endRowSpin = new NumericUpDown { Minimum = 1, Maximum = 1, Value = 1, Dock = DockStyle.Fill };

            t.Controls.Add(startCaption, 0, 1);
            t.Controls.Add(startRowSpin, 1, 1);
            t.Controls.Add(endCaption, 2, 1);
            t.Controls.Add(endRowSpin, 3, 1);

            autoDetectLabel = new Label
            {
                Text = "",
                AutoSize = true,
                MaximumSize = new Size(650, 0),
                TextAlign = ContentAlignment.TopRight,
                ForeColor = Color.FromArgb(26, 122, 26),
            };
            t.Controls.Add(autoDetectLabel, 0, 2);
            t.SetColumnSpan(autoDetectLabel, 5); // جلوگیری از برش خوردن متن طولانی در ستون AutoSize

            body.Controls.Add(t);
            sec3.FitToContent(); // ارتفاع مرحله ۳ = محتوای واقعی (بدون فاضلهٔ اضافی زیر آخرین لیبل)
        }

        private void OnHeaderToggle()
        {
            if (headerCheck.Checked)
            {
                startRowSpin.Value = Math.Min(2, maxRow);
                headerValues = previewRows.Count > 0 ? previewRows[0].ToList() : new List<string>();
            }
            else
            {
                startRowSpin.Value = 1;
                headerValues = new List<string>();
            }
            endRowSpin.Value = maxRow;

            AttemptAutoDetect();
            RefreshColumnChoices();
        }

        private void AttemptAutoDetect()
        {
            autoDetected.Clear();
            if (headerCheck.Checked && headerValues.Count > 0)
            {
                foreach (var kv in FieldCatalog.AutoDetectKeywords)
                {
                    for (int idx = 0; idx < headerValues.Count; idx++)
                    {
                        var text = (headerValues[idx] ?? "").Trim().ToLowerInvariant();
                        if (string.IsNullOrEmpty(text)) continue;
                        if (kv.Value.Any(k => text.Contains(k.ToLowerInvariant())))
                        {
                            autoDetected[kv.Key] = idx + 1;
                            break;
                        }
                    }
                }
            }

            var found = FieldCatalog.RequiredFields.Where(f => autoDetected.ContainsKey(f.Key)).Select(f => f.Label).ToList();
            var missingRequired = FieldCatalog.RequiredFields.Where(f => f.Required && !autoDetected.ContainsKey(f.Key)).Select(f => f.Label).ToList();

            if (found.Count > 0)
            {
                var msg = "تشخیص خودکار: " + string.Join("، ", found);
                if (missingRequired.Count > 0)
                    msg += "\nستون‌های الزامی زیر را به‌صورت دستی در مرحله ۵ انتخاب کنید: " + string.Join("، ", missingRequired);
                autoDetectLabel.Text = msg;
            }
            else
            {
                autoDetectLabel.Text = "تشخیص خودکار ستون‌ها ممکن نشد؛ لطفاً در مرحله ۵ ستون‌ها را به‌صورت دستی انتخاب کنید.";
            }

            sec3.FitToContent(); // متن سبز ممکن است یک یا دو خطی شود؛ ارتفاع مرحله ۳ را با آن تنظیم کن
        }

        // ================= مرحله ۴ =================
        private void BuildStep4(Panel body)
        {
            previewGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                // Fill: ستون‌ها با هم کل عرض کانتاینر را پر می‌کنند (فضای خالی در انتها نمی‌ماند)
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                // ارتفاع هر سطر بر اساس متن (فونت Vazirmatn) تا متنِ سطرها هرگز بریده نشود
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders,
            };
            body.Controls.Add(previewGrid);
        }

        // ================= مرحله ۵ =================
        private void BuildStep5(Panel body)
        {
            var fields = FieldCatalog.RequiredFields;
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = fields.Count + 1 };
            // چیدمان RTL: کپشن (راست‌ترین) سمت راست با عرض دقیق متن؛ داپ‌داون در سمت چپِ آن بقیهٔ عرض را می‌گیرد
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // caption (راست‌ترین)
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // combo

            for (int i = 0; i < fields.Count; i++)
            {
                var field = fields[i];
                var combo = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    DrawMode = DrawMode.OwnerDrawFixed, // رنگ‌آیی آیتم‌ها با رنگ ستونِ مربوط در پیش‌نمایش
                };
                mappingCombos[field.Key] = combo;
                combo.SelectedIndexChanged += (_, _) => OnMappingChanged(field.Key);
                combo.DrawItem += OnMappingComboDrawItem;

                var mark = field.Required ? " *" : "";
                var caption = new Label { Text = field.Label + mark + ":", AutoSize = true, TextAlign = ContentAlignment.MiddleRight, Anchor = AnchorStyles.Left | AnchorStyles.Right };

                t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                t.Controls.Add(caption, 0, i);
                t.Controls.Add(combo, 1, i);
            }

            var note = new Label
            {
                Text = "فیلدهای دارای * الزامی هستند.",
                AutoSize = true,
                ForeColor = Color.FromArgb(170, 51, 51),
                TextAlign = ContentAlignment.MiddleRight,
            };
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(note, 0, fields.Count);
            t.SetColumnSpan(note, 2);

            body.Controls.Add(t);
            sec5.FitToContent(); // ارتفاع مرحله ۵ = محتوای واقعی (آخرین داپ‌داون و متن قرمز کامل دیده شوند)
        }

        /// <summary>انتخاب ستون در مرحله ۵: نگاشت را به‌روز و ستون‌های تکراری را از سایر کمبوباکس‌ها حذف می‌کند.</summary>
        private void OnMappingChanged(string fieldKey)
        {
            if (suppressMappingEvents) return;
            if (workbook == null) return; // هنگام بستن فایل، بازسازی نکن
            if (!mappingCombos.TryGetValue(fieldKey, out var combo)) return;

            var col = (combo.SelectedItem as ColumnOption)?.ColumnIndex ?? 0;
            if (col > 0) autoDetected[fieldKey] = col;
            else autoDetected.Remove(fieldKey);

            RefreshColumnChoices();
        }

        /// <summary>کمبوباکس‌های مرحله ۵ را بازسازی می‌کند به‌طوری که هر ستون فقط به یک فیلد اختصاص یابد
        /// و سپس ستون‌های نگاشت‌شده را در پیش‌نمایش هایلایت می‌کند.</summary>
        private void RefreshColumnChoices()
        {
            if (workbook == null || maxCol < 1) return;

            // ۱) رفع تعارض: اولویت با فیلدهای اول است؛ ستونِ تکراری از فیلدهای بعدی پس گرفته می‌شود
            var claimed = new HashSet<int>();
            foreach (var field in FieldCatalog.RequiredFields)
            {
                if (autoDetected.TryGetValue(field.Key, out int col) && !claimed.Add(col))
                    autoDetected.Remove(field.Key);
            }

            var options = new List<ColumnOption>();
            for (int c = 1; c <= maxCol; c++)
            {
                string headerText = "";
                if (headerCheck.Checked && c - 1 < headerValues.Count)
                    headerText = (headerValues[c - 1] ?? "").Trim();
                string label = string.IsNullOrEmpty(headerText) ? GetColumnLetter(c) : $"{GetColumnLetter(c)} - {headerText}";
                options.Add(new ColumnOption(c, label));
            }

            suppressMappingEvents = true;
            try
            {
                foreach (var field in FieldCatalog.RequiredFields)
                {
                    if (!mappingCombos.TryGetValue(field.Key, out var combo)) continue;

                    int own = autoDetected.TryGetValue(field.Key, out int sel) ? sel : 0;

                    combo.Items.Clear();
                    combo.Items.Add(new ColumnOption(0, ""));
                    foreach (var opt in options)
                    {
                        // ستونِ متعلق به فیلد دیگر را نشان نده
                        if (opt.ColumnIndex != 0 && opt.ColumnIndex != own && claimed.Contains(opt.ColumnIndex))
                            continue;
                        combo.Items.Add(opt);
                    }

                    ColumnOption? match = null;
                    if (own > 0)
                    {
                        foreach (var item in combo.Items)
                        {
                            if (item is ColumnOption co && co.ColumnIndex == own) { match = co; break; }
                        }
                    }
                    combo.SelectedItem = match ?? combo.Items[0];
                }
            }
            finally
            {
                suppressMappingEvents = false;
            }

            ApplyColumnHighlights();
        }

        /// <summary>ستون‌های نگاشت‌شده را در پیش‌نمایش با رنگ اختصاصی هر ستون هایلایت می‌کند.</summary>
        private void ApplyColumnHighlights()
        {
            if (previewGrid == null || previewGrid.Columns.Count == 0) return;

            // رنگِ متمایز برای هر ستونِ نگاشت‌شده — منبع واحد رنگ‌ها: BuildColumnColors (با کمبوباکس مرحله ۵ هماهنگ است)
            var colorByCol = BuildColumnColors();
            var labelByCol = new Dictionary<int, string>();
            foreach (var field in FieldCatalog.RequiredFields)
                if (autoDetected.TryGetValue(field.Key, out int c) && c > 0 && colorByCol.ContainsKey(c))
                    labelByCol.TryAdd(c, field.Label);

            // ستون اول (در صورت وجود) ستون شمارهٔ ردیف است و در محاسبهٔ شمارهٔ ستون اکسل دخالت ندارد
            int offset = previewGrid.Columns.Count > 0 && previewGrid.Columns[0].Name == "__rowNo" ? 1 : 0;
            if (offset == 1)
            {
                var noHeader = previewGrid.Columns[0].HeaderCell;
                noHeader.Style.BackColor = Color.FromArgb(235, 235, 235);
                noHeader.Style.ForeColor = Color.Empty;
                noHeader.ToolTipText = null;
                for (int r = 0; r < previewGrid.Rows.Count; r++)
                    previewGrid.Rows[r].Cells[0].Style.BackColor = Color.White;
            }

            for (int i = offset; i < previewGrid.Columns.Count; i++)
            {
                int excelCol = i - offset + 1;
                var headerCell = previewGrid.Columns[i].HeaderCell;
                var headerStyle = headerCell.Style;

                if (colorByCol.TryGetValue(excelCol, out var color))
                {
                    headerStyle.BackColor = color;
                    headerStyle.ForeColor = Color.Black;
                    headerCell.ToolTipText = $"نگاشت‌شده به: {labelByCol[excelCol]}";
                    var cellColor = Lighten(color);
                    for (int r = 0; r < previewGrid.Rows.Count; r++)
                        previewGrid.Rows[r].Cells[i].Style.BackColor = cellColor;
                }
                else
                {
                    headerStyle.BackColor = Color.White;
                    headerStyle.ForeColor = Color.Empty;
                    headerCell.ToolTipText = null;
                    for (int r = 0; r < previewGrid.Rows.Count; r++)
                        previewGrid.Rows[r].Cells[i].Style.BackColor = Color.White;
                }
            }
            previewGrid.Refresh();
            foreach (var c in mappingCombos.Values) c.Invalidate(); // رنگ‌آیی مجدد آیتم‌های dropdown مرحله ۵
        }

        /// <summary>نسخه روشن‌تر رنگ (مخلوط با سفید) برای سلول‌های بدنه.</summary>
        private static Color Lighten(Color c) =>
            Color.FromArgb((c.R + 255) / 2, (c.G + 255) / 2, (c.B + 255) / 2);

        /// <summary>رنگِ هر ستونِ نگاشت‌شده بر اساس وضعیت فعلی نگاشت — منبع واحد برای هایلایتِ پیش‌نمایش و رنگ آیتم‌های مرحله ۵.</summary>
        private Dictionary<int, Color> BuildColumnColors()
        {
            var colorByCol = new Dictionary<int, Color>();
            var usedColors = new HashSet<Color>();
            foreach (var field in FieldCatalog.RequiredFields)
            {
                if (!autoDetected.TryGetValue(field.Key, out int col) || col <= 0) continue;
                if (colorByCol.ContainsKey(col)) continue; // نگاشت تکراری (نباید رخ دهد)

                int idx = (col - 1) % ColumnPalette.Length;
                while (!usedColors.Add(ColumnPalette[idx]))
                    idx = (idx + 1) % ColumnPalette.Length; // در صورت تداخل، اولین رنگ آزاد
                colorByCol[col] = ColumnPalette[idx];
            }
            return colorByCol;
        }

        /// <summary>رنگ‌آیی آیتم‌های کمبوباکس مرحله ۵: زمینهٔ هر گزینه همان رنگی است که آن ستون در پیش‌نمایش دارد (نگاشت‌نشده → سفید).</summary>
        private void OnMappingComboDrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not ComboBox combo || e.Index < 0 || e.Index >= combo.Items.Count) return;

            Color bg = Color.White;
            if (combo.Items[e.Index] is ColumnOption opt && opt.ColumnIndex > 0)
            {
                var colors = BuildColumnColors();
                if (colors.TryGetValue(opt.ColumnIndex, out var mapped))
                    bg = mapped;
            }

            using (var brush = new SolidBrush(bg))
                e.Graphics.FillRectangle(brush, e.Bounds);

            var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis
                      | (combo.RightToLeft == RightToLeft.Yes ? TextFormatFlags.Right : TextFormatFlags.Left);
            var rect = new Rectangle(e.Bounds.X + 6, e.Bounds.Y, Math.Max(0, e.Bounds.Width - 10), e.Bounds.Height);
            var text = combo.Items[e.Index]?.ToString() ?? "";
            TextRenderer.DrawText(e.Graphics, text, e.Font, rect, Color.Black, flags);
        }

        // ================= تولید فایل خروجی =================
        private void OnGenerateClicked()
        {
            try
            {
                GenerateFile();
            }
            catch (Exception ex)
            {
                MessageForm.ShowError(this, "خطای غیرمنتظره", ex.Message, details: ex.ToString());
            }
        }

        private void GenerateFile()
        {
            if (currentSheet == null)
            {
                MessageForm.ShowWarning(this, "توجه", "ابتدا فایل اکسل و شیت را انتخاب کنید.");
                return;
            }

            var mapping = new Dictionary<string, int>();
            foreach (var field in FieldCatalog.RequiredFields)
            {
                var combo = mappingCombos[field.Key];
                var selected = combo.SelectedItem as ColumnOption;
                int col = selected?.ColumnIndex ?? 0;
                if (field.Required && col == 0)
                {
                    MessageForm.ShowWarning(this, "توجه", $"ستون «{field.Label}» را انتخاب کنید.");
                    return;
                }
                mapping[field.Key] = col;
            }

            int start = (int)startRowSpin.Value;
            int end = (int)endRowSpin.Value;
            if (start > end)
            {
                MessageForm.ShowWarning(this, "توجه", "سطر شروع نباید بزرگ‌تر از سطر پایان باشد.");
                return;
            }

            string CellText(int row, string key)
            {
                int col = mapping.TryGetValue(key, out var c) ? c : 0;
                if (col <= 0) return "";
                var cell = currentSheet!.Cell(row, col);
                return cell.IsEmpty() ? "" : cell.GetString();
            }

            var validRows = new List<ValidRow>();
            var errors = new List<(int Row, string Reason)>();

            for (int r = start; r <= end; r++)
            {
                string account = CellText(r, "ACCOUNT_NUMBER").Trim();
                string amountText = CellText(r, "AMOUNT");
                string amountRaw = PersianTextHelper.CleanNumeric(amountText);

                if (string.IsNullOrEmpty(account) && string.IsNullOrEmpty(amountRaw))
                    continue; // ردیف کاملاً خالی

                if (string.IsNullOrEmpty(account))
                {
                    errors.Add((r, "شماره حساب خالی است"));
                    continue;
                }

                double amountVal;
                if (string.IsNullOrEmpty(amountRaw))
                {
                    amountVal = 0;
                }
                else
                {
                    var parsed = PersianTextHelper.ParseAmount(amountRaw);
                    if (parsed == null)
                    {
                        errors.Add((r, $"مبلغ نامعتبر: {amountText}"));
                        continue;
                    }
                    amountVal = parsed.Value;
                }

                if (amountVal <= 0)
                {
                    errors.Add((r, "مبلغ باید بزرگ‌تر از صفر باشد"));
                    continue;
                }

                validRows.Add(new ValidRow
                {
                    AccountNumber = PersianTextHelper.ToAsciiDigits(account),
                    Amount = amountVal,
                    TransactionNumber = PersianTextHelper.ToAsciiDigits(CellText(r, "TRANSACTION_NUMBER").Trim()),
                    Note = CellText(r, "NOTE").Trim(),
                    DepositorName = CellText(r, "DEPOSITOR_NAME").Trim(),
                });
            }

            if (validRows.Count == 0)
            {
                MessageForm.ShowError(this, "خطا", "هیچ ردیف معتبری برای تولید فایل یافت نشد.");
                return;
            }

            if (errors.Count > 0)
            {
                var previewErrors = string.Join("\n", errors.Take(10).Select(e => $"سطر {e.Row}: {e.Reason}"));
                var more = errors.Count > 10 ? $"\n... و {errors.Count - 10} مورد دیگر" : "";
                var proceed = MessageForm.Show(this, "ردیف‌های نامعتبر",
                    $"{errors.Count} ردیف نامعتبر یافت و نادیده گرفته شد:\n{previewErrors}{more}\n\n" +
                    $"آیا می‌خواهید با {validRows.Count} ردیف معتبر ادامه دهید؟",
                    MessageKind.Warning, buttons: MessageButtons.YesNo);
                if (proceed != DialogResult.Yes) return;
            }

            double totalAmount = validRows.Sum(r => r.Amount);
            var confirm = MessageForm.ShowQuestion(this, "تأیید قبل از ذخیره",
                $"تعداد ردیف‌های معتبر: {validRows.Count}\nجمع کل مبلغ: {totalAmount:N0}\n\nآیا اطلاعات فوق تأیید می‌شود؟");
            if (confirm != DialogResult.Yes) return;

            string defaultName = "FL" + JalaliDate.TodayYyMmDd() + ".PAY";
            using var saveDlg = new SaveFileDialog
            {
                Title = "ذخیره فایل خروجی",
                FileName = defaultName,
                Filter = "Payment File|*.PAY|All Files|*.*",
                DefaultExt = "PAY",
            };
            if (saveDlg.ShowDialog(this) != DialogResult.OK) return;

            var (content, badCount) = BuildFileContent(validRows);
            try
            {
                File.WriteAllBytes(saveDlg.FileName, content);
            }
            catch (Exception ex)
            {
                MessageForm.ShowError(this, "خطا در ذخیره فایل", ex.Message, details: ex.ToString());
                return;
            }

            var msg = $"فایل با موفقیت ذخیره شد:\n{saveDlg.FileName}";
            if (badCount > 0)
                msg += $"\n\nهشدار: {badCount} نویسه غیرقابل‌تبدیل با '?' جایگزین شد.";
            MessageForm.ShowSuccess(this, "پایان عملیات", msg);
            statusLabel.Text = "فایل خروجی با موفقیت تولید شد.";
        }

        private static (byte[] Content, int BadCount) BuildFileContent(List<ValidRow> rows)
        {
            static string Pad(string value, int width, PadMode mode)
            {
                string s = value ?? "";
                if (s.Length > width)
                    s = mode == PadMode.LeftZero ? s.Substring(s.Length - width) : s.Substring(0, width);
                return s.PadLeft(width, mode == PadMode.LeftZero ? '0' : ' ');
            }

            int numDeposits = rows.Count;
            double totalAmount = rows.Sum(r => r.Amount);

            var lines = new List<string>
            {
                numDeposits.ToString().PadLeft(10, '0') + ((long)Math.Round(totalAmount)).ToString().PadLeft(15, '0'),
            };

            foreach (var r in rows)
            {
                string line =
                    Pad(r.AccountNumber, 10, PadMode.LeftZero) +
                    Pad(((long)Math.Round(r.Amount)).ToString(), 15, PadMode.LeftZero) +
                    Pad(r.TransactionNumber, 17, PadMode.LeftZero) +
                    Pad(r.Note, 30, PadMode.RightSpace) +
                    Pad(r.DepositorName, 30, PadMode.RightSpace);
                lines.Add(line);
            }

            string textContent = string.Join("\n", lines) + "\n";
            return PersianTextHelper.SafeCp1256Bytes(textContent);
        }

        private static string GetColumnLetter(int col)
        {
            string letter = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                letter = (char)('A' + rem) + letter;
                col = (col - 1) / 26;
            }
            return letter;
        }
    }
}
