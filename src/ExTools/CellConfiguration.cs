using System.Collections.Generic;

namespace ExTools
{
    public static class CellConfiguration
    {
        /// <summary>
        /// Gets the number formats.
        /// </summary>
        /// <value>
        /// The number formats.
        /// </value>
        public static Dictionary<NumberFormat,string> NumberFormats => new Dictionary<NumberFormat, string>
        {
            { NumberFormat.General, "General" },
            { NumberFormat.Number, "0" },
            { NumberFormat.Currency, "0.00" },
            { NumberFormat.ShortDate, "dd.mm.yyyy hh:mm" },
            { NumberFormat.Text, "@" },
        };
    }
}