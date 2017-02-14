using ExTools.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExTools
{
    internal static class ValidationProvider
    {
        //TODO consider refactor. Reason namedRange is passed through multiple methods and only used later
        public static List<SheetMessage> Validate(
            List<DataRow> dataRowCollection,
            Collection<IDataValidation> validations,
            Dictionary<string, object[]> namedRangeValues)
        {
            if (dataRowCollection == null)
            {
                throw new ArgumentNullException(nameof(dataRowCollection));
            }
            if (validations == null)
            {
                throw new ArgumentNullException(nameof(validations));
            }
            
            var sortedValidations = validations.ToDictionary(v => v.ColumnNumber);
            var errors = new List<SheetMessage>();

            for (var i = 0; i < dataRowCollection.Count; i++)
            {
                var rowId = i + 2;
                var messages = ValidateRow(dataRowCollection[i], sortedValidations, namedRangeValues)
                    .ToList();

                foreach (var sheetMessage in messages)
                {
                    sheetMessage.Row = rowId;
                    errors.Add(sheetMessage);
                }
            }

            return errors;
        }

        private static IEnumerable<SheetMessage> ValidateRow(
            DataRow dataRow,
            IReadOnlyDictionary<int, IDataValidation> dataValidations,
            IReadOnlyDictionary<string, object[]> namedRangeValues)
        {
            for (var i = 0; i < dataRow.Count; i++)
            {
                var columnId = i + 1;

                if (!dataValidations.ContainsKey(columnId))
                {
                    continue;
                }

                var errorMsg = ValidateCell(dataRow[i], dataValidations[columnId], namedRangeValues);

                if (string.IsNullOrEmpty(errorMsg))
                {
                    continue;
                }

                yield return new SheetMessage
                {
                    Column = columnId,
                    Message = errorMsg
                };
            }
        }

        private static string ValidateCell(
            DataCell cell,
            IDataValidation validation,
            IReadOnlyDictionary<string, object[]> namedRangeValues)
        {
            bool isValid;
            switch (validation.ValidationType)
            {
                case DataValidationType.Text:
                    var textValid = (IntegerValidation)validation;
                    isValid = IsValid((string)cell, textValid.FirstFormula.Value, textValid.SecondFormula.Value);
                    break;
                case DataValidationType.WholeNumber:
                    var intValid = (IntegerValidation)validation;
                    isValid = IsValid(cell.ToInt(), intValid.FirstFormula.Value, intValid.SecondFormula.Value);
                    break;
                case DataValidationType.List:
                    var listValidation = (ListValidation)validation;
                    var usedNameRange = namedRangeValues[listValidation.Formula.ExcelFormula];
                    isValid = IsValid(cell.ToString(), usedNameRange, listValidation.AllowBlank);
                    break;
                case DataValidationType.Custom:
                case DataValidationType.Date:
                case DataValidationType.Time:
                    isValid = true;
                    break;
                default:
                    isValid = true;
                    break;
            }

            return isValid
                ? string.Empty
                : validation.Error;
        }

        private static bool IsInRange(int value, int min, int max)
        {
            return value >= min && value <= max;
        }

        private static bool IsValid(string value, int? minValue, int? maxValue)
        {
            if (minValue == null)
            {
                throw new ArgumentNullException(nameof(minValue), ErrorMessages.NULL_VALUE);
            }
            if (maxValue == null)
            {
                throw new ArgumentNullException(nameof(maxValue), ErrorMessages.NULL_VALUE);
            }

            return value != null && IsInRange(value.Length, minValue.Value, maxValue.Value);
        }
        
        private static bool IsValid(string value, IEnumerable<object> listValues, bool? allowBlank)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value), ErrorMessages.NULL_VALUE);
            }

            var blanksAllowed = allowBlank != null
                                && allowBlank.Value;

            if (!blanksAllowed && string.IsNullOrEmpty(value))
            {
                return false;
            }

            return blanksAllowed
                   || listValues.Any(v => v.ToString()
                       .Equals(value, StringComparison.OrdinalIgnoreCase));
        }
        
        private static bool IsValid(int value, int? minValue, int? maxValue)
        {
            if (minValue == null)
            {
                throw new ArgumentNullException(nameof(minValue), ErrorMessages.NULL_VALUE);
            }
            if (maxValue == null)
            {
                throw new ArgumentNullException(nameof(maxValue), ErrorMessages.NULL_VALUE);
            }

            return IsInRange(value, minValue.Value, maxValue.Value);
        }
    }
}