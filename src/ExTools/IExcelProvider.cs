using System.Collections.Generic;
using ExTools.Models;

namespace ExTools
{
    public interface IExcelProvider
    {
        /// <summary>
        /// Sets the workbook.
        /// </summary>
        /// <param name="newWorkbook">The new workbook.</param>
        void SetWorkbook(ExcelTemplate newWorkbook);

        /// <summary>
        /// Gets the workbook.
        /// </summary>
        /// <returns></returns>
        ExcelTemplate GetWorkbook();

        /// <summary>
        /// Gets the sheet data.
        /// </summary>
        /// <typeparam name="TDocument">The type of the document.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="worksheetConfig">The worksheet configuration.</param>
        /// <returns></returns>
        Sheet<TDocument> GetSheetData<TDocument>(string sheetName, Worksheet worksheetConfig);

        /// <summary>
        /// Loads the data.
        /// </summary>
        /// <typeparam name="TRowData">The type of the row data.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="dataCollection">The data collection.</param>
        /// <param name="printHeaders">if set to <c>true</c> [print headers].</param>
        void LoadData<TRowData>(string sheetName, IEnumerable<TRowData> dataCollection, bool printHeaders = true);

        /// <summary>
        /// Extracts the data.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        List<T> ExtractData<T>(string sheetName);

        /// <summary>
        /// Autofits the columns.
        /// </summary>
        /// <param name="sheetname">The sheetname.</param>
        void AutofitColumns(string sheetname);
    }
}