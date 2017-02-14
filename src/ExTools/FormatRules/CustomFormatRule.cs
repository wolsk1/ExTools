using System.Drawing;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace ExTools
{
    public class CustomFormatRule : BaseRuleFormat, IExcelConditionalFormattingWithFormula
    {
        public CustomFormatRule(int columnNumber, string formula, Color? backgroundColor, bool hasHeaders = true) 
            : base(columnNumber, backgroundColor, hasHeaders)
        {
            Formula = formula;
        }

        public string Formula { get; set; }
    }
}