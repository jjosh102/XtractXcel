using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

internal class TypedWorksheetExtractor<T> where T : new()
{
    private static readonly ConcurrentDictionary<PropertyInfo, Action<T, object?>> _propertySetters = [];

    // Cache the setters for the most common types in Excel
    private static readonly ConcurrentDictionary<Type, Func<PropertyInfo, Action<T, object?>>> _setterFactories = new()
    {
        [typeof(string)] = prop => (obj, value) => prop.SetValue(obj, value?.ToString()),
        [typeof(int)] = prop => CreateValueSetter(prop, Convert.ToInt32),
        [typeof(int?)] = prop => CreateNullableValueSetter(prop, Convert.ToInt32),
        [typeof(double)] = prop => CreateValueSetter(prop, Convert.ToDouble),
        [typeof(double?)] = prop => CreateNullableValueSetter(prop, Convert.ToDouble),
        [typeof(decimal)] = prop => CreateValueSetter(prop, Convert.ToDecimal),
        [typeof(decimal?)] = prop => CreateNullableValueSetter(prop, Convert.ToDecimal),
        [typeof(DateTime)] = prop => CreateValueSetter(prop, Convert.ToDateTime),
        [typeof(DateTime?)] = prop => CreateNullableValueSetter(prop, Convert.ToDateTime),
        [typeof(bool)] = prop => CreateValueSetter(prop, Convert.ToBoolean),
        [typeof(bool?)] = prop => CreateNullableValueSetter(prop, Convert.ToBoolean),
        [typeof(TimeSpan)] = prop => CreateValueSetter(prop, value => TimeSpan.Parse(value.ToString()!)),
        [typeof(TimeSpan?)] = prop => CreateNullableValueSetter(prop, value => TimeSpan.Parse(value.ToString()!))
    };

    public List<T> ExtractDataFromWorksheet(IXLWorksheet worksheet, bool readHeader)
    {
        var extractedData = new List<T>();
        var excelRange = worksheet.RangeUsed();
        if (excelRange is null) return extractedData;

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();
        var mappings = BuildAttributeMappings(worksheet, readHeader);

        foreach (var row in excelRows)
        {
            var obj = new T();
            foreach (var (colIndex, setter) in mappings)
            {
                setter(obj, GetCellValue(row.Cell(colIndex)));
            }
            extractedData.Add(obj);
        }

        return extractedData;
    }

    public List<T> ExtractDataFromWorksheet(IXLWorksheet worksheet, Func<IXLRangeRow, T> mapRow, bool readHeader)
    {
        var excelRange = worksheet.RangeUsed();

        if (excelRange is null) return [];

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();

        return excelRows.Select(mapRow).ToList();
    }

    private static Dictionary<int, Action<T, object?>> BuildAttributeMappings(IXLWorksheet worksheet, bool readHeader)
    {

        var mappings = new Dictionary<int, Action<T, object?>>();
        var columnIndices = worksheet.Row(1).CellsUsed()
                  .ToDictionary(c => c.GetString(), c => c.Address.ColumnNumber);

        if (readHeader)
        {
            foreach (var propInfo in GetExcelColumnAttributeProperties())
            {
                foreach (var columnName in propInfo.Attribute.ColumnNames)
                {
                    if (columnIndices.TryGetValue(columnName, out int colIndex))
                    {
                        mappings[colIndex] = GetSetterForProperty(propInfo.Property);
                        break;
                    }
                }
            }
        }
        else
        {
            var properties = GetProperties();

            foreach (var (propInfo, columnIndex) in properties.Select((p, i) => (p, i + 1)))
            {
                mappings[columnIndex] = GetSetterForProperty(propInfo);
            }
        }

        return mappings;
    }


    private static  Action<T, object?> GetSetterForProperty(PropertyInfo property)
    {
        if (_propertySetters.TryGetValue(property, out var cachedSetter)) return cachedSetter;

        var setter = _setterFactories.TryGetValue(property.PropertyType, out var factory)
            ? factory(property)
            : CreateGenericSetter(property);

        // Caches the generic setter
        return _propertySetters[property] = setter;
    }

    private static Action<T, object?> CreateValueSetter<TValue>(PropertyInfo property, Func<object, TValue> converter) where TValue : struct
    {
        return (obj, value) =>
        {
            if (value is not null)
            {
                property.SetValue(obj, converter(value));
            }
            else
            {
                property.SetValue(obj, default(TValue));
            }
        };
    }

    private static Action<T, object?> CreateNullableValueSetter<TValue>(PropertyInfo property, Func<object, TValue> converter) where TValue : struct
    {
        return (obj, value) =>
        {
            if (value is not null)
            {
                property.SetValue(obj, converter(value));
            }
            else
            {
                property.SetValue(obj, null);
            }
        };
    }

    private static Action<T, object?> CreateGenericSetter(PropertyInfo property)
    {
        // This is a fallback for a generic setter in case no specific setter is found in _setterFactories
        var instance = Expression.Parameter(typeof(T), "instance");
        var value = Expression.Parameter(typeof(object), "value");

        var propertyType = property.PropertyType;
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        var convertedValue = Expression.Convert(
            Expression.Call(typeof(Convert).GetMethod(nameof(Convert.ChangeType), [typeof(object), typeof(Type)])!,
                value, Expression.Constant(targetType)),
            propertyType);

        var assign = Expression.Assign(Expression.Property(instance, property), convertedValue);
        return Expression.Lambda<Action<T, object?>>(assign, instance, value).Compile();
    }

    private static object? GetCellValue(IXLCell cell)
    {
        var type = cell.Value.Type;

        if (type == XLDataType.Blank)
            return null;

        return type switch
        {
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Text => cell.GetString(),
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            XLDataType.Error => cell.GetError(),
            _ => cell.GetString()
        };
    }

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnAttributeProperties()
    {
        var propertiesWithAttributes = typeof(T).GetProperties()
            .Select(p => new { Property = p, Attribute = p.GetCustomAttribute<ExcelColumnAttribute>() })
            .Where(p => p.Attribute != null)
            .Select(p => (p.Property, p.Attribute!))
            .ToList();

        return propertiesWithAttributes.Count > 0
            ? propertiesWithAttributes
            : throw new InvalidOperationException($"No properties with {nameof(ExcelColumnAttribute)} found on type {typeof(T).Name}");
    }

    private static List<PropertyInfo> GetProperties()
    {
        var properties = typeof(T).GetProperties();
        return properties.Length > 0
            ? properties.ToList()
            : throw new InvalidOperationException($"No properties found on type {typeof(T).Name}");
    }
}
