using ExTools.Models;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExTools
{
    public class ExcelProvider : IExcelProvider
    {
        private ExcelTemplate workbook;
        private Dictionary<string, object[]> namedRangeValues;

        public ExcelProvider()
        {
            using (var template = new ExcelTemplate())
            {
                SetWorkbook(template);
            }
        }

        public void SetWorkbook(ExcelTemplate newWorkbook)
        {
            if (newWorkbook == null)
            {
                throw new ArgumentNullException(nameof(newWorkbook));
            }

            workbook = newWorkbook;
            GetNamedRangeValues();
        }

        public ExcelTemplate GetWorkbook()
        {
            return workbook;
        }

        public Sheet<TDocument> GetSheetData<TDocument>(string sheetName, Worksheet worksheetConfig)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (worksheetConfig == null)
            {
                throw new ArgumentNullException(nameof(worksheetConfig));
            }
           
            var worksheet = GetWorksheet(sheetName);
            
            var rows = SheetLoader.GetDataRows(worksheet)
                .ToList();
            var sheetMessages = ValidationProvider.Validate(rows, worksheetConfig.DataValidations, namedRangeValues);
            var sheetData = rows
                .Select(r => (TDocument)Activator.CreateInstance(typeof(TDocument), r))
                .ToList();
            var sheet = new Sheet<TDocument>(sheetName, sheetData);

            foreach (var sheetMessage in sheetMessages)
            {
                sheetMessage.SheetName = sheetName;
                sheet.Messages.Add(sheetMessage);
            }

            return sheet;
        }

        public void LoadData<TRowData>(string sheetName, IEnumerable<TRowData> dataCollection, bool printHeaders = true)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (dataCollection == null)
            {
                throw new ArgumentNullException(nameof(dataCollection));
            }

            var worksheet = GetWorksheet(sheetName);
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
        
        public List<T> ExtractData<T>(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var rowList = GetDataRows(sheetName)
                .ToList();

            return rowList
                .Select(r => (T)Activator.CreateInstance(typeof(T), r))
                .ToList();
        }

        public Dictionary<int, IExcelDataValidation> GetDataValidations(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var worksheet = GetWorksheet(sheetName);

            return worksheet.DataValidations.ToDictionary(d => d.Address.Start.Column, d => d);
        }

        public IEnumerable<DataRow> GetDataRows(string sheetName, bool hasHeaders = true)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var worksheet = GetWorksheet(sheetName);

            if (worksheet.Dimension == null)
                throw new ArgumentOutOfRangeException(nameof(worksheet.Dimension));

            var coulmnIndexMapping = GetColumnCellMapping(worksheet, hasHeaders);
            var fromRow = hasHeaders ? 2 : Constants.MIN_ROW_ID;

            for (var rowId = fromRow; rowId <= worksheet.Dimension.Rows; rowId++)
            {
                yield return new DataRow(ExtractDataCells(worksheet, rowId), coulmnIndexMapping);
            }

        }
       
        public void AutofitColumns(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            var worksheet = GetWorksheet(sheetName);
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        private ExcelWorksheet GetWorksheet(string sheetName)
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == sheetName);

            if (worksheet == null)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetName), ErrorMessages.SHEET_NOT_EXIST);
            }

            return worksheet;
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
                //TODO Add Utility class
                //coulumnIndexMapping.Add(hasHeaders ? Utils.RemoveWhiteSpaces(headerList[i]) : i.ToString(), i);
            }

            return coulumnIndexMapping;
        }

        private void GetNamedRangeValues()
        {
            namedRangeValues = new Dictionary<string, object[]>();

            foreach (var namedRange in workbook.NamedRanges)
            {
                var valueList = namedRange.Value as object[,];

                if (valueList == null)
                {
                    continue;
                }

                var nonNullValues = valueList.Cast<object>()
                    .Where(v => v != null)
                    .ToArray();
                namedRangeValues.Add(namedRange.Name, nonNullValues);
            }
        }
    }
}