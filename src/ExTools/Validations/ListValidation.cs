using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;

namespace ExTools
{
    public class ListValidation : BaseValidation
    {
        private readonly bool hasHeaders;

        /// <summary>
        /// Initializes a new instance of the <see cref="ListValidation" /> class.
        /// </summary>
        /// <param name="columnNumber">The column number.(Must be greater than 0)</param>
        /// <param name="validationFormula">The validation formula.</param>
        /// <param name="hasHeaders">if set to <c>true</c> [has headers].</param>
        /// <exception cref="System.ArgumentOutOfRangeException">{nameof(columnNumber)}</exception>
        public ListValidation(int columnNumber, ListValidationFormula validationFormula, bool hasHeaders = true)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber), columnNumber, $"{nameof(columnNumber)} value must be greater than 0");
            }

            if (validationFormula == null)
            {
                throw new ArgumentNullException(nameof(validationFormula));
            }

            ColumnNumber = columnNumber;
            this.hasHeaders = hasHeaders;
            Formula = validationFormula;
        }
      
        public override string Address => ExcelCellBase.GetAddress(RowNumber, ColumnNumber, ExcelPackage.MaxRows, ColumnNumber);
       
        public override DataValidationType ValidationType => DataValidationType.List;
       
        public override int ColumnNumber { get; }
      
        public override int RowNumber => hasHeaders ? 2 : 1;
        
        public IExcelDataValidationFormulaList Formula { get; set; }
    }
}