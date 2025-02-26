
using System.Linq.Expressions;
using System.Reflection;

using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public sealed class ExcelDataExtractor<T> where T : new()
{
    private readonly ExcelExtractorOptions _options;

    public ExcelDataExtractor(ExcelExtractorOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }
    
    public IReadOnlyList<T> ExtractData()
    {
        using var workbook = _options.Source switch
        {
            SourceType.FilePath => new XLWorkbook(_options.FilePath!),
            SourceType.Stream => new XLWorkbook(_options.Stream!),
            SourceType.None => throw new InvalidOperationException("Either a file path or a stream must be provided."),
            _ => throw new InvalidOperationException("Invalid source type.")
        };

        var worksheet = workbook.Worksheet(_options.SheetIndex);
        return ExtractDataFromWorksheet(worksheet);
    }

    private List<T> ExtractDataFromWorksheet(IXLWorksheet worksheet)
    {
        List<T> extractedData = [];
        var excelRange = worksheet.RangeUsed();

        if (excelRange is not null)
        {
            var excelRows = _options.ReadHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();
            var mappings = GetColumnMappings(worksheet);

            foreach (var row in excelRows)
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

        return extractedData;
    }

    private Dictionary<int, Action<T, object?>> GetColumnMappings(IXLWorksheet worksheet)
    {
        var mappings = new Dictionary<int, Action<T, object?>>();
        var columnIndices = worksheet.Row(1).CellsUsed()
            .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);

        if (_options.ReadHeader)
        {
            foreach (var propInfo in GetExcelColumnAttributeProperties())
            {
                foreach (var columnName in propInfo.Attribute.ColumnNames)
                {
                    if (columnIndices.TryGetValue(columnName, out int colIndex))
                    {
                        var setter = CreateSetter(propInfo.Property);
                        mappings[colIndex] = setter;
                        break;
                    }
                }
            }
        }
        else
        {
            //Only retrieve the number of properties that match the number of used columns
            var properties = GetProperties().Take(columnIndices.Count + 1);
            foreach (var (propInfo, columnIndex) in properties.Select((p, i) => (p, i + 1)))
            {
                mappings[columnIndex] = CreateSetter(propInfo);
            }
        }

        return mappings;
    }
    
    private static Action<T, object?> CreateSetter(PropertyInfo property)
    {
        var instance = Expression.Parameter(typeof(T), "instance");
        var value = Expression.Parameter(typeof(object), "value");

        var propertyType = property.PropertyType;
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

        var isNullable = propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
        var defaultValue = Expression.Default(propertyType);

        var convertedValue = Expression.Condition(
            // Check if value is null
            Expression.Equal(value, Expression.Constant(null, typeof(object))),
            //If is nullable, assign null, else assign default(T)
            isNullable ? Expression.Constant(null, propertyType) : defaultValue,
            // Convert to actual value based on type
            Expression.Convert(
                Expression.Call(
                    typeof(Convert).GetMethod(nameof(Convert.ChangeType), [typeof(object), typeof(Type)])!,
                    value,
                    Expression.Constant(targetType)
                ),
                propertyType
            )
        );

        var propertyAccess = Expression.Property(instance, property);
        var assign = Expression.Assign(propertyAccess, convertedValue);

        return Expression.Lambda<Action<T, object?>>(assign, instance, value).Compile();
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

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnAttributeProperties()
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

        return propertiesWithAttributes.Count == 0
            ? throw new InvalidOperationException($"No properties with {nameof(ExcelColumnAttribute)} found on type {typeof(T).Name}")
            : propertiesWithAttributes;
    }

    private static List<PropertyInfo> GetProperties()
    {
        var properties = typeof(T).GetProperties();

        return properties.Length > 0
            ? properties.ToList()
            : throw new InvalidOperationException($"No properties found on type {typeof(T).Name}");
    }

}
