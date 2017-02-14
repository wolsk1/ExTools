using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace ExTools
{
    public class IntegerValidationFormula : IExcelDataValidationFormulaInt
    {
        public IntegerValidationFormula(int value)
        {
            Value = value;
        }

        /// <summary>
        /// Gets or sets the excel formula.
        /// </summary>
        /// <value>
        /// The excel formula.
        /// </value>
        public string ExcelFormula { get; set; }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public int? Value { get; set; }
    }
}