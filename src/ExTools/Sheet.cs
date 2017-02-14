using System;
using System.Collections.Generic;
using ExTools.Models;

namespace ExTools
{
    public class Sheet<T>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Sheet{T}" /> class.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="dataCollection">The data collection.</param>
        /// <exception cref="System.ArgumentNullException">sheetName</exception>
        public Sheet(string sheetName, IEnumerable<T> dataCollection)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }
            if (dataCollection == null)
            {
                throw new ArgumentNullException(nameof(dataCollection));
            }
            SheetName = sheetName;
            DataCollection = dataCollection;
            Messages = new List<SheetMessage>();
        }

        /// <summary>
        /// Gets or sets the name of the sheet.
        /// </summary>
        /// <value>
        /// The name of the sheet.
        /// </value>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets the data collection.
        /// </summary>
        /// <value>
        /// The data collection.
        /// </value>
        public IEnumerable<T> DataCollection { get; }

        /// <summary>
        /// Gets the messages.
        /// </summary>
        /// <value>
        /// The messages.
        /// </value>
        public List<SheetMessage> Messages { get; }
    }
}