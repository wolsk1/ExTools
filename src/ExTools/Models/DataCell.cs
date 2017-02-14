using System;

namespace ExTools.Models
{
    public class DataCell : IEquatable<DataCell>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DataCell"/> class.
        /// </summary>
        /// <param name="value">The value.</param>
        public DataCell(object value)
        {
            Value = value;
        }

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public object Value { get; }

        /// <summary>
        /// Performs an implicit conversion from <see cref="DataCell"/> to <see cref="System.String"/>.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The result of the conversion.
        /// </returns>
        /// <exception cref="System.ArgumentNullException"></exception>
        public static implicit operator string(DataCell cell)
        {
            return cell?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Performs an implicit conversion from <see cref="DataCell"/> to <see cref="System.Int32"/>.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>
        /// The result of the conversion.
        /// </returns>
        public static implicit operator int(DataCell cell)
        {
            return cell?.ToInt() ?? 0;
        }

        /// <summary>
        /// Casts this instance.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T Cast<T>()
        {
            return Value == null || Value is DBNull 
                ? default(T) 
                : (T) Convert.ChangeType(Value, typeof(T));
        }

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other" /> parameter; otherwise, false.
        /// </returns>
        public bool Equals(DataCell other)
        {
            if (other == null)
            {
                return false;
            }

            return Value == other.Value;
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

            return obj.GetType() == GetType() 
                && Equals(obj as DataCell);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return Value?.GetHashCode() ?? 0;
        }

        /// <summary>
        /// Returns a <see cref="string" /> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="string" /> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            return string.IsNullOrEmpty(Value?.ToString()) 
                ? string.Empty 
                : Value?.ToString().Trim(' ');
        }

        /// <summary>
        /// To the int.
        /// </summary>
        /// <returns>Number</returns>
        public int ToInt()
        {
            int number;

            return int.TryParse(Value?.ToString(), out number)
                ? number
                : 0;
        }
    }
}