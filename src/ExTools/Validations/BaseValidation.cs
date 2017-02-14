using OfficeOpenXml.DataValidation;

namespace ExTools
{
    public abstract class BaseValidation : IDataValidation
    {
        protected BaseValidation()
        {
            ShowErrorMessage = true;
            AllowBlank = false;
            AllowsOperator = true;
            ShowInputMessage = true;
            ShowErrorMessage = true;
            Operator = ExcelDataValidationOperator.any;
            ColumnNumber = 1;
            RowNumber = 2;
        }

        public abstract string Address { get; }

        public abstract DataValidationType ValidationType { get; }

        public bool AllowBlank { get; set; }

        public bool ShowInputMessage { get; set; }

        public bool ShowErrorMessage { get; set; }

        public string ErrorTitle { get; set; }

        public string Error { get; set; }

        public string PromptTitle { get; set; }

        public string Prompt { get; set; }

        public bool AllowsOperator { get; set; }

        public virtual ExcelDataValidationOperator Operator { get; }
        
        public virtual int ColumnNumber { get; }

        public virtual int RowNumber { get; }

        public ExcelDataValidationWarningStyle ErrorStyle { get; set; }
    }
}