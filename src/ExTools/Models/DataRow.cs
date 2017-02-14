using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExTools.Models
{
    public class DataRow : Collection<DataCell>, IEquatable<DataRow>
    {
        private readonly IDictionary<string, int> columnIndexMapping;

        /// <summary>
        /// Initializes a new instance of the <see cref="DataRow"/> class.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// <param name="columnIndexMapping">The column index mapping.</param>
        /// <exception cref="System.ArgumentNullException"></exception>
        public DataRow(IEnumerable<DataCell> cells, IDictionary<string, int> columnIndexMapping)
        {
            if (cells == null)
            {
                throw new ArgumentNullException(nameof(cells));
            }

            var cellList = cells.ToList();

            foreach (var cell in cellList)
            {
                Add(cell);
            }

            this.columnIndexMapping = columnIndexMapping;
        }

        /// <summary>
        /// Gets or sets the column names.
        /// </summary>
        /// <value>
        /// The column names.
        /// </value>
        public IEnumerable<string> ColumnNames => columnIndexMapping.Keys;

        /// <summary>
        /// Gets the <see cref="DataCell" /> with the specified column name.
        /// </summary>
        /// <value>
        /// The <see cref="DataCell" />.
        /// </value>
        /// <param name="columnName">Name of the column.</param>
        /// <returns></returns>
        public DataCell this[string columnName]
        {
            get
            {
                if (!columnIndexMapping.ContainsKey(columnName))
                { 
                    throw new ArgumentException($"'{columnName as object}' column name does not exist. Valid column names are '{string.Join("', '", columnIndexMapping.Keys.ToArray()) as object}'");
                }

                return this[columnIndexMapping[columnName]];
            }
        }

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other" /> parameter; otherwise, false.
        /// </returns>
        public bool Equals(DataRow other)
        {
            return other != null && ColumnNames.SequenceEqual(other.ColumnNames);
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" />, is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj))
            {
                return false;
            }
            if (ReferenceEquals(this, obj))
            {
                return true;
            }

            return obj.GetType() == GetType() && Equals(obj as DataRow);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return columnIndexMapping?.GetHashCode() ?? 0;
        }
    }
}