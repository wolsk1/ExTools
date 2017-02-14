using OfficeOpenXml.DataValidation;

namespace ExTools
{
    public interface IDataValidation
    {
        /// <summary>
        /// Data validation range address
        /// </summary>
        /// <value>
        /// The address.
        /// </value>
        string Address { get; }

        /// <summary>
        /// Gets or sets the type of the validation.
        /// </summary>
        /// <value>
        /// The type of the validation.
        /// </value>
        DataValidationType ValidationType { get; }

        /// <summary>
        /// Gets or sets the allow blank.
        /// </summary>
        /// <value>
        /// The allow blank.
        /// </value>
        bool AllowBlank { get; set; }

        /// <summary>
        /// Gets or sets the show input message.
        /// </summary>
        /// <value>
        /// The show input message.
        /// </value>
        bool ShowInputMessage { get; set; }

        /// <summary>
        /// Gets or sets the show error message.
        /// </summary>
        /// <value>
        /// The show error message.
        /// </value>
        bool ShowErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the error title.
        /// </summary>
        /// <value>
        /// The error title.
        /// </value>
        string ErrorTitle { get; set; }

        /// <summary>
        /// Gets or sets the error.
        /// </summary>
        /// <value>
        /// The error.
        /// </value>
        string Error { get; set; }

        /// <summary>
        /// Gets or sets the prompt title.
        /// </summary>
        /// <value>
        /// The prompt title.
        /// </value>
        string PromptTitle { get; set; }

        /// <summary>
        /// Gets or sets the prompt.
        /// </summary>
        /// <value>
        /// The prompt.
        /// </value>
        string Prompt { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [allows operator].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [allows operator]; otherwise, <c>false</c>.
        /// </value>
        bool AllowsOperator { get; set; }

        /// <summary>
        /// Gets or sets the operator.
        /// </summary>
        /// <value>
        /// The operator.
        /// </value>
        ExcelDataValidationOperator Operator { get; }

        /// <summary>
        /// Gets or sets the column number.
        /// </summary>
        /// <value>
        /// The column number.
        /// </value>
        int ColumnNumber { get; }

        /// <summary>
        /// Gets the row number.
        /// </summary>
        /// <value>
        /// The row number.
        /// </value>
        int RowNumber { get; }
    }
}