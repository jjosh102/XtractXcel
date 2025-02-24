
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
            var mappings = GetColumnMappings<T>(worksheet);

            foreach (var row in excelRange.RowsUsed().Skip(1))
            {
                var obj = new T();

                foreach (var (colIndex, setter) in mappings)
                {
                    var cell = row.Cell(colIndex);
                    var value = GetCellValue(cell);
                    setter(obj, value);
                }

                extractedData.Add(obj);
            }
        }

        return extractedData.AsReadOnly();
    }

    private static Dictionary<int, Action<T, object?>> GetColumnMappings<T>(IXLWorksheet worksheet)
    {
        // The purpose of this is to precompile property.SetValue to use only ONCE and avoid excessive use of reflection during runtime
        var mappings = new Dictionary<int, Action<T, object?>>();
        var properties = GetExcelColumnProperties<T>();

        // Cache header lookup
        var columnIndices = worksheet.Row(1).CellsUsed()
            .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);

        foreach (var propInfo in properties)
        {
            foreach (var columnName in propInfo.Attribute.ColumnNames)
            {
                if (columnIndices.TryGetValue(columnName, out int colIndex))
                {
                    var setter = CreateSetter<T>(propInfo.Property);
                    mappings[colIndex] = setter;
                    break;
                }
            }
        }

        return mappings;
    }

    private static Action<T, object?> CreateSetter<T>(PropertyInfo property)
    {
        return (instance, value) =>
        {
            var targetType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            var convertedValue = value is not null ? Convert.ChangeType(value, targetType) : null;
            property.SetValue(instance, convertedValue); // Reflection only happens here ONCE
        };
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
        var properties = typeof(T).GetProperties()
           .Where(p => p.GetCustomAttribute<ExcelColumnAttribute>() is not null)
           .ToDictionary(p => p.GetCustomAttribute<ExcelColumnAttribute>()!.ColumnNames.First(), p => p);

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
