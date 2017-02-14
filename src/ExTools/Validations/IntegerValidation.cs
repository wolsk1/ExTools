using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;

namespace ExTools
{
    public class IntegerValidation : BaseValidation
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValidation" /> class.
        /// </summary>
        /// <param name="columnNumber">The column number.(Must be greater than 0)</param>
        /// <param name="validationOperator">The validation operator.</param>
        /// <param name="firstValidationFormula">The first validation formula.</param>
        /// <param name="secondValidationFormula">The second validation formula.[Optional]</param>
        /// <exception cref="System.ArgumentOutOfRangeException">{nameof(columnNumber)}</exception>
        /// <exception cref="System.ArgumentNullException"></exception>
        public IntegerValidation(
            int columnNumber,
            ExcelDataValidationOperator validationOperator,
            IntegerValidationFormula firstValidationFormula,
            IntegerValidationFormula secondValidationFormula = null)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber), columnNumber, $"{nameof(columnNumber)} value must be greater than 0");
            }

            if (firstValidationFormula == null)
            {
                throw new ArgumentNullException(nameof(firstValidationFormula));
            }

            ColumnNumber = columnNumber;
            Operator = validationOperator;
            FirstFormula = firstValidationFormula;
            SecondFormula = secondValidationFormula;
        }

        public override string Address => ExcelCellBase.GetAddress(RowNumber, ColumnNumber, ExcelPackage.MaxRows, ColumnNumber);
       
        public override DataValidationType ValidationType => DataValidationType.Text;
       
        public override ExcelDataValidationOperator Operator { get; }
        
        public override int ColumnNumber { get; }

        public IExcelDataValidationFormulaInt FirstFormula { get; set; }

        public IExcelDataValidationFormulaInt SecondFormula { get; set; }
    }
}