namespace ExTools.Models
{
    public class SheetMessage
    {
        public string SheetName { get; set; }

        public int Row { get; set; }

        public int Column { get; set; }

        public string Message { get; set; }
    }
}