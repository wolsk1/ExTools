using ExTools.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExTools
{
    public static class SheetLoader
    {
        /// <summary>
        /// Loads the rows.
        /// </summary>
        /// <typeparam name="TRowData">The type of the row data.</typeparam>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="dataCollection">The data collection.</param>
        /// <param name="printHeaders">if set to <c>true</c> [print headers].</param>
        /// <exception cref="System.ArgumentNullException">worksheet
        /// or
        /// dataCollection</exception>
        /// <exception cref="ArgumentNullException"></exception>
        public static void LoadData<TRowData>(ExcelWorksheet worksheet, IEnumerable<TRowData> dataCollection, bool printHeaders = true)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }
            if (dataCollection == null)
            {
                throw new ArgumentNullException(nameof(dataCollection));
            }

            worksheet.Cells.LoadFromCollectionFiltered(dataCollection);

            if (!printHeaders)
            {
                return;
            }

            var headerNames = ExcelUtils.GetClassAsHeaders(typeof(TRowData))
                .ToList();
            worksheet.InsertRow(1, 1);
            AddWorksheetHeaders(worksheet, headerNames);
        }

        /// <summary>
        /// Extracts the data rows.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet">The worksheet.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException"></exception>
        public static List<T> ExtractData<T>(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException(nameof(worksheet));
            }
            var rowList = GetDataRows(worksheet)
                .ToList();

            return rowList
                .Select(r => (T) Activator.CreateInstance(typeof(T), r))
                .ToList();
        }

        public static IEnumerable<DataRow> GetDataRows(ExcelWorksheet worksheet, bool hasHeaders = true)
        {
            if (worksheet.Dimension == null)
                throw new ArgumentOutOfRangeException(nameof(worksheet.Dimension));

            var coulmnIndexMapping = GetColumnCellMapping(worksheet, hasHeaders);
            var fromRow = hasHeaders ? 2 : Constants.MIN_ROW_ID;

            for (var rowId = fromRow; rowId <= worksheet.Dimension.Rows; rowId++)
            {
                yield return new DataRow(ExtractDataCells(worksheet, rowId), coulmnIndexMapping);
            }

        }
        
        private static void AddWorksheetHeaders(ExcelWorksheet worksheet, IList<string> columnNames)
        {
            if (columnNames == null)
            {
                throw new ArgumentNullException(nameof(columnNames));
            }
            var columnId = Constants.MIN_COLUMN_ID;
            foreach (var columnName in columnNames)
            {
                worksheet.Cells[1, columnId].Value = columnName;
                columnId++;
            }
        }
       
        private static IEnumerable<DataCell> ExtractDataCells(ExcelWorksheet worksheet, int rowId)
        {
            if (rowId < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(rowId));
            }

            for (var colId = Constants.MIN_COLUMN_ID; colId <= worksheet.Dimension.Columns; colId++)
            {
                yield return new DataCell(worksheet.Cells[rowId, colId].Value);
            }
        }

        private static IEnumerable<string> GetHeadings(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null)
            {
                throw new ArgumentOutOfRangeException(nameof(worksheet.Dimension));
            }

            for (var columnId = 1; columnId <= worksheet.Dimension.Columns; columnId++)
            {
                var cellValue = worksheet.Cells[1, columnId].Value;
                yield return cellValue?.ToString() ?? columnId.ToString();
            }
        }
        
        private static Dictionary<string, int> GetColumnCellMapping(ExcelWorksheet worksheet, bool hasHeaders = true)
        {
            var headings = GetHeadings(worksheet);
            var coulumnIndexMapping = new Dictionary<string, int>();
            var headerList = headings.ToList();

            for (var i = 0; i < headerList.Count; i++)
            {
                //TODO Utils
                //coulumnIndexMapping.Add(hasHeaders ? Utils.RemoveWhiteSpaces(headerList[i]) : i.ToString(), i);
            }

            return coulumnIndexMapping;
        }
    }
}