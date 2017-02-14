namespace ExTools
{
    public class BaseRuleFormat : BaseCellRef
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseRuleFormat"/> class.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        /// <param name="backgroundColor">Color of the background.</param>
        /// <param name="hasHeaders">if set to <c>true</c> [has headers].</param>
        public BaseRuleFormat(int columnNumber, System.Drawing.Color? backgroundColor, bool hasHeaders = true)
            : base(columnNumber, hasHeaders)
        {
            ColumnNumber = columnNumber;
            BackgroundColor = backgroundColor;
        }

        public System.Drawing.Color? BackgroundColor { get; set; }
    }
}