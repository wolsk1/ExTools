using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ExTools
{
    public static class ExcelUtils
    {
        /// <summary>
        /// Gets the number format.
        /// </summary>
        /// <param name="numberFormat">The number format.</param>
        /// <returns></returns>
        public static string GetNumberFormat(NumberFormat numberFormat)
        {
            var format = CellConfiguration.NumberFormats.FirstOrDefault(f => f.Key == numberFormat);

            return format.Value;
        }

        /// <summary>
        /// Gets the object property names.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="bindingFlags">The binding flags.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException"></exception>
        public static Collection<string> GetTypePropertyNames(Type type, BindingFlags bindingFlags = BindingFlags.Public)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            //TODO extract property info extraction in seperate method
            var propertyInfos = type.GetProperties(bindingFlags);

            if (propertyInfos.Any())
            {
                return new Collection<string>(propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList());
            }

            var runtime = type.GetRuntimeProperties()
                .ToList();
            
            var runtimeProperties = runtime.Where(p => !Attribute.IsDefined(p, typeof (ExcelExportIgnore)));

            return new Collection<string>(runtimeProperties.Select(propertyInfo => propertyInfo.Name).ToList());
        }

        /// <summary>
        /// Gets the property value.
        /// </summary>
        /// <param name="classObject">The class object.</param>
        /// <param name="propName">Name of the property.</param>
        /// <returns></returns>
        public static object GetPropertyValue(this object classObject, string propName)
        {
            if (classObject == null)
            {
                throw new ArgumentNullException(nameof(classObject));
            }

            if (string.IsNullOrEmpty(propName))
            {
                throw new ArgumentNullException(nameof(propName));
            }

            var type = classObject.GetType();
            var propertyInfo = type.GetProperty(propName);

            return propertyInfo.GetValue(classObject);
        }

        /// <summary>
        /// Transforms string from the camel case to human readable text.
        /// </summary>
        /// <param name="sourceString">The source string.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static string FromPascalCase(this string sourceString)
        {
            if (string.IsNullOrEmpty(sourceString))
            {
                throw new ArgumentNullException(nameof(sourceString));
            }

            var regex = new Regex(@"
                (?<=[A-Z])(?=[A-Z][a-z]) |
                 (?<=[^A-Z])(?=[A-Z]) |
                 (?<=[A-Za-z])(?=[^A-Za-z])", RegexOptions.IgnorePatternWhitespace);

            return regex.Replace(sourceString, " ");
        }

        /// <summary>
        /// To the pascal case.
        /// </summary>
        /// <param name="sourceString">The source string.</param>
        /// <returns></returns>
        public static string ToPascalCase(this string sourceString)
        {
            if (sourceString == null)
            {
                return null;
            }

            if (sourceString.Length < 2)
            {
                return sourceString.ToUpper();
            }
           
            var words = sourceString.Split(
                new char[] { },
                StringSplitOptions.RemoveEmptyEntries);

            return words.Aggregate(string.Empty, (current, word) => 
            current + word.Substring(0, 1).ToUpper() + word.Substring(1));
        }

        /// <summary>
        /// Gets the class as headers.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        public static IEnumerable<string> GetClassAsHeaders(Type type)
        {
            var propertyNames = GetTypePropertyNames(type);

            return propertyNames.Select(n => n.FromPascalCase());
        }

        public static ExcelRangeBase LoadFromCollectionFiltered<T>(
            this ExcelRangeBase cellRange,
            IEnumerable<T> collection,
            bool printHeaders = false)
        {
            if (cellRange == null)
            {
                throw new ArgumentNullException(nameof(cellRange));
            }
            if (collection == null)
            {
                throw new ArgumentNullException(nameof(collection));
            }

            MemberInfo[] membersToInclude = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => !Attribute.IsDefined(p, typeof(ExcelExportIgnore)))
                .ToArray();

            return cellRange.LoadFromCollection(collection, printHeaders,
                OfficeOpenXml.Table.TableStyles.None,
                BindingFlags.Instance | BindingFlags.Public,
                membersToInclude);
        }
    }
}