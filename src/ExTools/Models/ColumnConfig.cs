namespace ExTools.Models
{
    public class ColumnConfig
    {
        private double width;

        public ColumnConfig(int number)
        {
            Number = number;
            WrapText = false;
            NumberFormat = NumberFormat.General;
            Width = Constants.DEFAULT_COL_WIDTH;
            Autofit = false;
        }

        public int Number { get; set; }

        public bool WrapText { get; set; }

        public bool Autofit { get; set; }

        public NumberFormat NumberFormat { get; set; }

        public double Width
        {
            get { return width; }
            set
            {
                width = value > Constants.MAX_COL_WIDTH
                    ? Constants.MAX_COL_WIDTH : value;
            }
        }
    }
}