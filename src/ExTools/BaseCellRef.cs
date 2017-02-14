using System;
using OfficeOpenXml;

namespace ExTools
{
    //TODO use this class as base class for DataValidation classes
    public class BaseCellRef
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="BaseCellRef"/> class.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        /// <param name="hasHeaders">if set to <c>true</c> [has headers].</param>
        /// <exception cref="System.ArgumentOutOfRangeException">columnNumber - columnNumber</exception>
        protected BaseCellRef(int columnNumber, bool hasHeaders = true)
        {
            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber), columnNumber, $"{nameof(columnNumber)} value must be greater than 0");
            }

            ColumnNumber = columnNumber;

            this.hasHeaders = hasHeaders;
        }

        private readonly bool hasHeaders;

        public string Address => ExcelCellBase.GetAddress(RowNumber, ColumnNumber, ExcelPackage.MaxRows, ColumnNumber);

        public int ColumnNumber { get; set; }

        public int RowNumber => hasHeaders ? 2 : 1;
    }
}