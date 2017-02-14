using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace ExTools
{
    public class ContainsRuleFormat : BaseRuleFormat, IExcelConditionalFormattingWithText
    {
        public ContainsRuleFormat(int columnNumber, string text, System.Drawing.Color? backgroundColor = null)
            :base(columnNumber, backgroundColor)
        {
            Text = text;
        }

        public string Text { get; set; }
    }
}