using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExTools.Models
{
    public class Worksheet
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Worksheet" /> class.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="namedRanges">The named ranges.</param>
        /// <param name="dataValidations">The data validations.</param>
        /// <param name="columnConfigs">The column configs.</param>
        /// <param name="formatingSettings">The formating settings.</param>
        /// <exception cref="System.ArgumentNullException">sheetName</exception>
        public Worksheet(
            string sheetName,
            Collection<NamedRange> namedRanges = null,
            Collection<IDataValidation> dataValidations = null,
            Collection<ColumnConfig> columnConfigs = null,
            IEnumerable<BaseRuleFormat> formatingSettings = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            Name = sheetName;
            NamedRanges = namedRanges ?? new Collection<NamedRange>();
            DataValidations = dataValidations ?? new Collection<IDataValidation>();
            ColumConfig = columnConfigs ?? new Collection<ColumnConfig>();
            FormatingSettings = formatingSettings ?? new Collection<BaseRuleFormat>();
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; }

        /// <summary>
        /// Gets the data validations.
        /// </summary>
        /// <value>
        /// The data validations.
        /// </value>
        public Collection<IDataValidation> DataValidations { get; }

        /// <summary>
        /// Gets the named ranges.
        /// </summary>
        /// <value>
        /// The named ranges.
        /// </value>
        public Collection<NamedRange> NamedRanges { get; }

        /// <summary>
        /// Gets the colum number formats.
        /// </summary>
        /// <value>
        /// The colum number formats.
        /// </value>
        public Collection<ColumnConfig> ColumConfig { get; }

        /// <summary>
        /// Gets or sets the formating settings.
        /// </summary>
        /// <value>
        /// The formating settings.
        /// </value>
        public IEnumerable<BaseRuleFormat> FormatingSettings { get; set; }

    }
}