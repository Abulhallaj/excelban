using System.Collections.Generic;

namespace ExcelBan
{
    public enum PadMode
    {
        LeftZero,
        RightSpace,
    }

    public sealed class FieldSpec
    {
        public string Key { get; }
        public string Label { get; }
        public bool Required { get; }
        public int Width { get; }
        public PadMode Mode { get; }

        public FieldSpec(string key, string label, bool required, int width, PadMode mode)
        {
            Key = key;
            Label = label;
            Required = required;
            Width = width;
            Mode = mode;
        }
    }

    public static class FieldCatalog
    {
        public static readonly IReadOnlyList<FieldSpec> RequiredFields = new List<FieldSpec>
        {
            new FieldSpec("ACCOUNT_NUMBER", "شماره حساب", true, 10, PadMode.LeftZero),
            new FieldSpec("AMOUNT", "مبلغ", true, 15, PadMode.LeftZero),
            new FieldSpec("TRANSACTION_NUMBER", "شماره تراکنش", false, 17, PadMode.LeftZero),
            new FieldSpec("NOTE", "توضیحات", false, 30, PadMode.RightSpace),
            new FieldSpec("DEPOSITOR_NAME", "نام واریزکننده", false, 30, PadMode.RightSpace),
        };

        public static readonly IReadOnlyDictionary<string, string[]> AutoDetectKeywords =
            new Dictionary<string, string[]>
            {
                ["ACCOUNT_NUMBER"] = new[] { "account", "acc no", "شماره حساب", "شماره_حساب", "حساب" },
                ["AMOUNT"] = new[] { "amount", "مبلغ", "وجه" },
                ["TRANSACTION_NUMBER"] = new[] { "transaction", "پیگیری", "تراکنش", "ref" },
                ["NOTE"] = new[] { "note", "توضیح", "شرح", "بابت" },
                ["DEPOSITOR_NAME"] = new[] { "depositor", "نام واریز", "واریزکننده", "واریز کننده", "name", "نام" },
            };
    }

    /// <summary>یک گزینه‌ی ستون در کومبوباکس مرحله ۵؛ ToString برای نمایش در کومبوباکس استفاده می‌شود.</summary>
    public sealed class ColumnOption
    {
        public int ColumnIndex { get; }
        public string Display { get; }

        public ColumnOption(int columnIndex, string display)
        {
            ColumnIndex = columnIndex;
            Display = display;
        }

        public override string ToString() => Display;
    }

    public sealed class ValidRow
    {
        public string AccountNumber { get; set; } = "";
        public double Amount { get; set; }
        public string TransactionNumber { get; set; } = "";
        public string Note { get; set; } = "";
        public string DepositorName { get; set; } = "";
    }
}
