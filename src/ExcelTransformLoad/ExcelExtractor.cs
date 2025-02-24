
using System.Reflection;

using ClosedXML.Excel;

namespace ExcelTransformLoad;

public static class ExcelExtractor
{
    public static IReadOnlyList<T> ExtractDataFromStream<T>(Stream stream) where T : new()
    {
        ArgumentNullException.ThrowIfNull(stream, nameof(stream));
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);
        return GetExtractedData<T>(worksheet);
    }

    public static IReadOnlyList<T> ExtractDataFromFile<T>(string filepath) where T : new()
    {
        ArgumentNullException.ThrowIfNullOrWhiteSpace(filepath, nameof(filepath));
        using var workbook = new XLWorkbook(filepath);
        var worksheet = workbook.Worksheet(1);
        return GetExtractedData<T>(worksheet);
    }


    private static IReadOnlyList<T> GetExtractedData<T>(IXLWorksheet worksheet) where T : new()
    {
        var extractedData = new List<T>();
        var excelRange = worksheet.RangeUsed();

        if (excelRange is not null)
        {
            var properties = GetExcelColumnProperties<T>();
            // Cache header lookup
            var columnIndices = worksheet.Row(1).CellsUsed()
                .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);
            // Cache hot path property setters 
            var propertySetters = properties.ToDictionary(
                propInfo => propInfo,
                propInfo => new Action<T, object>(
                    (obj, value) => propInfo.Property.SetValue(obj, Convert.ChangeType(value, Nullable.GetUnderlyingType(propInfo.Property.PropertyType) ?? propInfo.Property.PropertyType))));

            foreach (var row in excelRange.RowsUsed().Skip(1))
            {
                var obj = new T();

                foreach (var propInfo in properties)
                {
                    foreach (var columnName in propInfo.Attribute?.ColumnNames ?? Array.Empty<string>())
                    {
                        if (columnIndices.TryGetValue(columnName, out int colIndex))
                        {
                            var cell = row.Cell(colIndex);
                            var value = GetCellValue(cell);

                            if (value is not null)
                            {
                                //Set the  value based on the proper type -> 23 to int, 23.5 to double, "23" to string
                                propertySetters[propInfo](obj, value);
                            }
                            else if (Nullable.GetUnderlyingType(propInfo.Property.PropertyType) is not null)
                            {
                                //Set the value to null if nullable
                                propInfo.Property.SetValue(obj, null);
                            }
                            else
                            {
                                //Set the default values if not nullable -> bool to false, int to 0, string to ""
                                propInfo.Property.SetValue(obj, Activator.CreateInstance(propInfo.Property.PropertyType));
                            }

                            break;
                        }
                    }
                }

                extractedData.Add(obj);
            }
        }

        return extractedData.AsReadOnly();
    }

    private static object? GetCellValue(IXLCell cell)
    {
        return cell.Value.Type switch
        {
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Text => cell.GetString(),
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            XLDataType.Error => cell.GetError(),
            XLDataType.Blank => null,
            _ => null
        };
    }

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnProperties<T>()
    {
        var propertiesWithAttributes = typeof(T).GetProperties()
            .Select(p => new
            {
                Property = p,
                Attribute = p.GetCustomAttribute<ExcelColumnAttribute>()
            })
            .Where(p => p.Attribute != null)
            .Select(p => (p.Property, p.Attribute!))
            .ToList();

        if (propertiesWithAttributes.Count == 0)
        {
            throw new InvalidOperationException($"No properties with {nameof(ExcelColumnAttribute)} found on type {typeof(T).Name}");
        }

        return propertiesWithAttributes;
    }

}
