using System;
using OfficeOpenXml;

namespace ExTools.Models
{
    public class NamedRange
    {
        private string rangeName;

        /// <summary>
        /// Initializes a new instance of the <see cref="NamedRange" /> class.
        /// </summary>
        /// <param name="rangeName">Name of the range.</param>
        /// <param name="sourceSheetName">Name of the source sheet.</param>
        /// <param name="columnNumber">The column number.(Must be greater than 0)</param>
        /// <param name="hasHeader">if set to <c>true</c> [has header].</param>
        /// <exception cref="System.ArgumentNullException">
        /// </exception>
        /// <exception cref="System.ArgumentOutOfRangeException">{nameof(column)}</exception>
        public NamedRange(string rangeName, string sourceSheetName, int columnNumber, bool hasHeader = true)
        {
            if (string.IsNullOrEmpty(rangeName))
            {
                throw new ArgumentNullException(nameof(rangeName));
            }

            if (string.IsNullOrEmpty(sourceSheetName))
            {
                throw new ArgumentNullException(nameof(sourceSheetName));
            }

            if (columnNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber),columnNumber,$"{nameof(columnNumber)} value must be greater than 0");
            }

            var firstRow = hasHeader ? 2 : 1;
            Address = ExcelCellBase.GetAddress(firstRow, columnNumber, ExcelPackage.MaxRows, columnNumber);
            RangeName = rangeName;
            SourceSheet = sourceSheetName;
        }

        /// <summary>
        /// Gets or sets the name of the range.
        /// </summary>
        /// <value>
        /// The name of the range.
        /// </value>
        /// <exception cref="ArgumentNullException"></exception>
        public string RangeName
        {
            get { return rangeName; }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    throw new ArgumentNullException(nameof(value));
                }

                rangeName = value.Replace(" ", "_");
            }
        }

        /// <summary>
        /// Gets or sets the source sheet.
        /// </summary>
        /// <value>
        /// The source sheet.
        /// </value>
        public string SourceSheet { get; set; }

        /// <summary>
        /// Gets or sets the owner sheet.
        /// </summary>
        /// <value>
        /// The owner sheet.
        /// </value>
        public string OwnerSheet { get; set; }

        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <value>
        /// The address.
        /// </value>
        public string Address { get; }
    }
}