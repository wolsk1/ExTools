namespace ExTools
{
    public class DuplicateRuleFormat: BaseRuleFormat
    {
        public DuplicateRuleFormat(int columnNumber, System.Drawing.Color? backgroundColor = null)
            : base(columnNumber, backgroundColor)
        {
            if (backgroundColor == null)
            {
                BackgroundColor = Constants.DuplicateColor;
            }
        }
    }
}