using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Collections.Generic;

namespace ExTools
{
    public class ListValidationFormula : IExcelDataValidationFormulaList
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ListValidationFormula"/> class.
        /// </summary>
        /// <param name="excelFormula">The excel formula.</param>
        public ListValidationFormula(string excelFormula)
        {
            ExcelFormula = excelFormula;
            Values = new List<string>();
        }

        /// <summary>
        /// An excel formula
        /// </summary>
        public string ExcelFormula { get; set; }

        /// <summary>
        /// A list of value strings.
        /// </summary>
        public IList<string> Values { get; }
    }
}