using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;

namespace ExTools
{
    public static class WorksheetProvider
    {
        /// <summary>
        /// Adds the duplicate value rule.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="formatSettings">The format settings.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// worksheet
        /// or
        /// formatSettings
        /// </exception>
        public static IExcelConditionalFormattingDuplicateValues AddFormatRule(this ExcelWorksheet worksheet, BaseRuleFormat formatSettings)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (formatSettings == null)
                throw new ArgumentNullException(nameof(formatSettings));

            var excelAddress = worksheet.GetCellRange(formatSettings.Address);
            var formatRule = worksheet.ConditionalFormatting.AddDuplicateValues(excelAddress);
            formatRule.Style.Fill.BackgroundColor.Color = formatSettings.BackgroundColor;

            return formatRule;
        }

        /// <summary>
        /// Adds the contains rule.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="formatSettings">The format settings.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// worksheet
        /// or
        /// formatSettings
        /// </exception>
        public static IExcelConditionalFormattingContainsText AddFormatRule(this ExcelWorksheet worksheet, ContainsRuleFormat formatSettings)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (formatSettings == null)
                throw new ArgumentNullException(nameof(formatSettings));

            var excelAddress = worksheet.GetCellRange(formatSettings.Address);
            var formatRule = worksheet.ConditionalFormatting.AddContainsText(excelAddress);
            formatRule.Style.Fill.BackgroundColor.Color = formatSettings.BackgroundColor;
            formatRule.Text = formatSettings.Text;

            return formatRule;
        }

        /// <summary>
        /// Adds the format rule.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="formatSettings">The format settings.</param>
        /// <exception cref="System.ArgumentNullException">
        /// worksheet
        /// or
        /// formatSettings
        /// </exception>
        public static void AddFormatRule(this ExcelWorksheet worksheet, CustomFormatRule formatSettings)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (formatSettings == null)
                throw new ArgumentNullException(nameof(formatSettings));

            var excelAddress = worksheet.GetCellRange(formatSettings.Address);
            var formatRule = worksheet.ConditionalFormatting.AddExpression(excelAddress);
            formatRule.Style.Fill.BackgroundColor.Color = formatSettings.BackgroundColor;
            formatRule.Formula = formatSettings.Formula;
        }

        private static ExcelRange GetCellRange(this ExcelWorksheet worksheet, string rangeAddress)
        {
            return worksheet.Cells[rangeAddress];
        }
    }
}