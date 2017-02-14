using ExTools.Models;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Contracts;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExTools
{
    public class ExcelTemplate : IDisposable
    {
        private bool isDisposed;

        /// <summary>
        /// The default sheet name
        /// </summary>
        private const string DEFAULT_SHEET_NAME = "Default";

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelTemplate"/> class.
        /// </summary>
        public ExcelTemplate()
        {
            Package = new ExcelPackage();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelTemplate" /> class.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <exception cref="System.ArgumentNullException"></exception>
        public ExcelTemplate(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            Package = new ExcelPackage(new FileInfo(fileName));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelTemplate"/> class.
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        public ExcelTemplate(Stream fileStream)
        {
            if (fileStream == null)
            {
                throw new ArgumentNullException(nameof(fileStream));
            }

            Package = new ExcelPackage(fileStream);
        }

        /// <summary>
        /// Gets or sets the package.
        /// </summary>
        /// <value>
        /// The package.
        /// </value>
        private ExcelPackage Package { get; set; }

        /// <summary>
        /// Gets the worksheets.
        /// </summary>
        /// <value>
        /// The worksheets.
        /// </value>
        public ExcelWorksheets Worksheets => Package.Workbook.Worksheets;

        /// <summary>
        /// Gets the named ranges.
        /// </summary>
        /// <value>
        /// The named ranges.
        /// </value>
        public ExcelNamedRangeCollection NamedRanges => Package.Workbook.Names;

        /// <summary>
        /// Adds the worksheet.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>
        /// Created instance of the <see cref="ExcelWorksheet" />
        /// </returns>
        /// <exception cref="System.ArgumentNullException"></exception>
        public ExcelWorksheet AddWorksheet(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            return Worksheets.Add(sheetName);
        }

        /// <summary>
        /// Adds the named range.
        /// </summary>
        /// <param name="namedRange">The named range.</param>
        /// <returns>
        /// Created instance of the <see cref="ExcelTemplate" />
        /// </returns>
        /// <exception cref="System.ArgumentNullException"></exception>
        public ExcelNamedRange AddNamedRange(NamedRange namedRange)
        {
            if (namedRange == null)
            {
                throw new ArgumentNullException(nameof(namedRange));
            }

            var belongsToSheet = string.IsNullOrEmpty(namedRange.OwnerSheet) 
                ? null 
                : Worksheets[namedRange.OwnerSheet];

            using (var excelNamedRange = new ExcelNamedRange(namedRange.RangeName, belongsToSheet, Worksheets[namedRange.SourceSheet], namedRange.Address, 1))
            {
                return NamedRanges.Add(namedRange.RangeName, excelNamedRange);
            }
        }

        /// <summary>
        /// Adds the data validation.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="dataValidation">The data validation.</param>
        /// <exception cref="System.ArgumentNullException"></exception>
        /// <exception cref="System.ArgumentException">validationSettings.ValidationType {dataValidation.ValidationType}</exception>
        public void AddDataValidation(string worksheetName, IDataValidation dataValidation)
        {
            if (dataValidation == null)
            {
                throw new ArgumentNullException(nameof(dataValidation));
            }

            //TODO implement other types of data validation
            switch (dataValidation.ValidationType)
            {
                case DataValidationType.Text:
                    AddValidation(worksheetName, (IntegerValidation)dataValidation);
                    break;
                case DataValidationType.List:
                    AddValidation(worksheetName, (ListValidation)dataValidation);
                    break;
                case DataValidationType.WholeNumber:
                    AddValidation(worksheetName, (IntegerValidation) dataValidation);
                    break;
                default:
                    throw new ArgumentException($"validationSettings.ValidationType {dataValidation.ValidationType} not implemented");
            }
        }

        /// <summary>
        /// Adds the data validation collection.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="dataValidations">The data validations.</param>
        /// <exception cref="System.ArgumentNullException">
        /// </exception>
        public void AddValidationCollection(string worksheetName, IEnumerable<IDataValidation> dataValidations)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException(nameof(worksheetName));
            }
            if (dataValidations == null)
            {
                throw new ArgumentNullException(nameof(dataValidations));
            }

            foreach (var dataValidation in dataValidations)
            {
                AddDataValidation(worksheetName, dataValidation);
            }
        }

        /// <summary>
        /// Adds the named range collection.
        /// </summary>
        /// <param name="namedRanges">The named ranges.</param>
        public void AddNamedRangeCollection(IEnumerable<NamedRange> namedRanges)
        {
            if (namedRanges == null)
            {
                throw new ArgumentNullException(nameof(namedRanges));
            }

            foreach (var namedRange in namedRanges)
            {
                AddNamedRange(namedRange);
            }
        }

        /// <summary>
        /// Adds the text validation.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="validationSettings">The validation settings.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// </exception>
        public IExcelDataValidationInt AddValidation(string worksheetName, IntegerValidation validationSettings)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException(nameof(worksheetName));
            }

            if (validationSettings == null)
            {
                throw new ArgumentNullException(nameof(validationSettings));
            }

            var textValidation = Worksheets[worksheetName].DataValidations.AddTextLengthValidation(validationSettings.Address);
            textValidation.Operator = validationSettings.Operator;
            textValidation.Formula.Value = validationSettings.FirstFormula.Value;
            textValidation.ShowErrorMessage = validationSettings.ShowErrorMessage;
            textValidation.AllowBlank = validationSettings.AllowBlank;
            textValidation.ErrorTitle = validationSettings.ErrorTitle;
            textValidation.Error = validationSettings.Error;
            if (validationSettings.SecondFormula == null)
            {
                return textValidation;
            }

            textValidation.Formula2.Value = validationSettings.SecondFormula.Value;

            return textValidation;
        }

        /// <summary>
        /// Adds the list validation.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="validationSettings">The validation settings.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// </exception>
        public IExcelDataValidationList AddValidation(string worksheetName, ListValidation validationSettings)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException(nameof(worksheetName));
            }
            if (validationSettings == null)
            {
                throw new ArgumentNullException(nameof(validationSettings));
            }
            
            var listValidation = Worksheets[worksheetName].DataValidations.AddListValidation(validationSettings.Address);
            listValidation.Formula.ExcelFormula = validationSettings.Formula.ExcelFormula;
            listValidation.ShowErrorMessage = validationSettings.ShowErrorMessage;
            listValidation.AllowBlank = validationSettings.AllowBlank;
            listValidation.ErrorTitle = validationSettings.ErrorTitle;
            listValidation.Error = validationSettings.Error;
            listValidation.Validate();

            return listValidation;
        }
        /// <summary>
        /// Exports as byte array.
        /// </summary>
        /// <returns></returns>
        public Stream Export()
        {
            if (Worksheets.Count == 0)
            {
                AddDefaultSheetToWorkbook();
            }

            return new MemoryStream(Package.GetAsByteArray());
        }

        /// <summary>
        /// Adds the default sheet to workbook.
        /// </summary>
        private void AddDefaultSheetToWorkbook()
        {
            AddWorksheet(DEFAULT_SHEET_NAME);
        }

        /// <summary>
        /// Loads the worksheet configuration.
        /// </summary>
        /// <param name="worksheetConfigs">The worksheet configs.</param>
        /// <exception cref="System.ArgumentNullException">worksheetConfigs</exception>
        public void LoadWorksheetConfig(IEnumerable<Worksheet> worksheetConfigs)
        {
            if (worksheetConfigs == null)
            {
                throw new ArgumentNullException(nameof(worksheetConfigs));
            }

            foreach (var worksheetConfig in worksheetConfigs)
            {
                var worksheet = Worksheets[worksheetConfig.Name];
                AddNamedRangeCollection(worksheetConfig.NamedRanges);
                AddValidationCollection(worksheetConfig.Name, worksheetConfig.DataValidations);
                SetColumns(worksheetConfig.Name, worksheetConfig.ColumConfig);
                foreach (var settings in worksheetConfig.FormatingSettings)
                {
                    var duplicateFormat = settings as DuplicateRuleFormat;
                    if (duplicateFormat != null)
                    {
                        worksheet.AddFormatRule(duplicateFormat);
                        continue;
                    }
                    var containsFormat = settings as ContainsRuleFormat;
                    if (containsFormat != null)
                    {
                        worksheet.AddFormatRule(containsFormat);
                        continue;
                    }
                    var customFormat = settings as CustomFormatRule;
                    if (customFormat == null)
                    {
                        continue;
                    }
                    worksheet.AddFormatRule(customFormat);
                }
            }
        }

        /// <summary>
        /// Sets the column types.
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="columnConfigs">The column configs.</param>
        /// <exception cref="ArgumentNullException">
        /// </exception>
        public void SetColumns(string worksheetName, Collection<ColumnConfig> columnConfigs)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException(nameof(worksheetName));
            }
            if (columnConfigs == null)
            {
                throw new ArgumentNullException(nameof(columnConfigs));
            }

            var worksheet = Worksheets.First(w => w.Name == worksheetName);

            foreach (var columnConfig in columnConfigs)
            {
                var format = CellConfiguration.NumberFormats.First(f => f.Key == columnConfig.NumberFormat);
                worksheet.Column(columnConfig.Number)
                    .Style.Numberformat.Format = format.Value;
                worksheet.Column(columnConfig.Number)
                    .Style.WrapText = columnConfig.WrapText;
                worksheet.Column(columnConfig.Number)
                    .Width = columnConfig.Width;
            }
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (isDisposed)
            {
                return;
            }

            if (disposing)
            {
                if (Package != null)
                {
                    Package.Dispose();
                    Package = null;
                }
            }

            isDisposed = true;
        }

        ~ExcelTemplate()
        {
            Dispose(false);
        }
    }
}