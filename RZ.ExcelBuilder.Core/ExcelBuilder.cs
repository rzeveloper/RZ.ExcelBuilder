using ClosedXML.Excel;
using RZ.ExcelBuilder.Core.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace RZ.ExcelBuilder.Core
{
    public class ExcelBuilder
    {
        /// <summary>
        /// Builds an Excel file from a collection of objects and returns its content as a Base64 string.
        /// </summary>
        /// <typeparam name="T">Type of objects in the collection.</typeparam>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="headers">Array of header names for the columns.</param>
        /// <param name="columns">List of column definitions.</param>
        /// <param name="collection">List of objects to export.</param>
        /// <returns>Base64 string representing the Excel file.</returns>
        public static string BuildExcel<T>(string worksheetName, string[] headers, List<ColumnBuilder> columns, List<T> collection) where T : class
        {
            try
            {
                using XLWorkbook workbook = new();
                IXLWorksheet worksheet = workbook.Worksheets.Add(worksheetName);
                worksheet.AddHeaders(headers);

                if (collection != null && collection.Count > 0)
                {
                    var propertiesInformation = GetPropertiesInformation(collection[0]);

                    if (!IsSubset([.. propertiesInformation.Select(x => x.Name)], [.. columns.Select(x => x.PropertyName)]))
                    {
                        throw new Exception("Properties Names does not exist.");
                    }
                }

                int row = 1;
                foreach (T instance in collection)
                {
                    var dictionary = DictionaryFromType(instance);

                    int col = 0;
                    foreach (ColumnBuilder column in columns)
                    {
                        object value = dictionary.TryGetValue(column.PropertyName, out var val) ? val : null;

                        var cell = worksheet.Cell(row + 1, ++col);

                        if (column.ColumnType == ColumnType.RUT)
                        {
                            cell.Value = value?.ToString().FormatRut() ?? string.Empty;
                        }
                        else
                        {
                            cell.Value = value switch
                            {
                                string s => s,
                                int i => i,
                                decimal d => d,
                                DateTime dt => dt,
                                _ => value?.ToString() ?? string.Empty
                            };
                        }

                        switch (column.ColumnType)
                        {
                            case ColumnType.STRING:
                            case ColumnType.RUT:
                                cell.Style.NumberFormat.Format = "@";
                                break;
                            case ColumnType.DATE:
                                cell.Style.ToDateFormat(column.Format);
                                break;
                            case ColumnType.DECIMAL:
                                cell.Style.ToMoneyFormat(true);
                                break;
                            case ColumnType.INTEGER:
                                cell.Style.ToNumberFormat(false);
                                break;
                        }
                    }

                    row++;
                }

                worksheet.RangeUsed().SetAutoFilter();
                worksheet.Columns().AdjustToContents();
                worksheet.SheetView.FreezeRows(1);

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);

                byte[] content = stream.ToArray();

                return Convert.ToBase64String(content, 0, content.Length);
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Converts an object instance to a dictionary with property names as keys and property values as values.
        /// </summary>
        /// <typeparam name="T">Type of the object.</typeparam>
        /// <param name="instance">Object instance to convert.</param>
        /// <returns>Dictionary of property names and values.</returns>
        private static Dictionary<string, object> DictionaryFromType<T>(T instance)
        {
            try
            {
                if (instance == null)
                {
                    return [];
                }

                Type type = instance.GetType();

                PropertyInfo[] properties = type.GetProperties();
                
                Dictionary<string, object> dictionary = [];
                
                foreach (PropertyInfo property in properties)
                {
                    object value = property.GetValue(instance, []);

                    dictionary.Add(property.Name, value);
                }

                return dictionary;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Gets the property information of an object instance.
        /// </summary>
        /// <typeparam name="T">Type of the object.</typeparam>
        /// <param name="instance">Object instance.</param>
        /// <returns>Array of PropertyInfo objects.</returns>
        private static PropertyInfo[] GetPropertiesInformation<T>(T instance)
        {
            try
            {
                if (instance == null)
                {
                    return null;
                }

                Type type = instance.GetType();

                PropertyInfo[] propertiesInformation = type.GetProperties();

                return propertiesInformation;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Checks if all elements in collection2 are present in collection1.
        /// </summary>
        /// <param name="collection1">The superset collection.</param>
        /// <param name="collection2">The subset collection.</param>
        /// <returns>True if collection2 is a subset of collection1, otherwise false.</returns>
        private static bool IsSubset(IEnumerable<string> collection1, IEnumerable<string> collection2)
        {
            var set2 = collection2.ToHashSet();
            
            return set2.IsSubsetOf(collection1);
        }
    }
}
