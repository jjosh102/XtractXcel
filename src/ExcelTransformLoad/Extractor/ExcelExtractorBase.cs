
using System.Linq.Expressions;
using System.Reflection;

using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public abstract class ExcelExtractorBase<T> where T : new()
{
    protected abstract XLWorkbook GetWorkbook();

    public IReadOnlyList<T> Extract()
    {
        using var workbook = GetWorkbook();
        var worksheet = workbook.Worksheet(1);
        return ExtractDataFromWorksheet(worksheet);
    }

    private IReadOnlyList<T> ExtractDataFromWorksheet(IXLWorksheet worksheet)
    {
        var extractedData = new List<T>();
        var excelRange = worksheet.RangeUsed();

        if (excelRange is not null)
        {
            var mappings = GetColumnMappings(worksheet);

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

    private Dictionary<int, Action<T, object?>> GetColumnMappings(IXLWorksheet worksheet)
    {
        // Precompile property.SetValue to use only ONCE and avoid excessive use of reflection during runtime
        var mappings = new Dictionary<int, Action<T, object?>>();
        var properties = GetExcelColumnProperties();

        // Cache header lookup
        var columnIndices = worksheet.Row(1).CellsUsed()
            .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);

        foreach (var propInfo in properties)
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
            // Convert to actual value based on targetType
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

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnProperties()
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
