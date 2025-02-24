
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
            // Cache column lookup
            var columnIndices = worksheet.Row(1).CellsUsed()
                .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);

            foreach (var row in excelRange.RowsUsed().Skip(1))
            {
                var obj = new T();

                foreach (var propInfo in properties)
                {
                    if (columnIndices.TryGetValue(propInfo.Attribute!.ColumnName, out int colIndex))
                    {
                        var cell = row.Cell(colIndex);
                        var value = GetCellValue(cell);

                        var propertyType = propInfo.Property.PropertyType;

                        // Handle nullabe type , get the actual type if it is nullable
                        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

                        if (value is not null)
                        {
                            //Convert the proper value -> 23 to int, 23.5 to double, "23" to string
                            propInfo.Property.SetValue(obj, Convert.ChangeType(value, targetType));
                        }
                        else if (Nullable.GetUnderlyingType(propertyType) is not null)
                        {
                            // If nullable and its value is null set it to null
                            propInfo.Property.SetValue(obj, null);
                        }
                        else
                        {
                            //Set the default values if not nullable
                            propInfo.Property.SetValue(obj, Activator.CreateInstance(propertyType));
                        }
                    }
                }

                extractedData.Add(obj);
            }
        }

        return extractedData.AsReadOnly<T>();
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
