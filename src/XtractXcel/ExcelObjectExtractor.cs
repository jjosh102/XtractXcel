using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace XtractXcel;

internal class ExcelObjectExtractor<TObject> where TObject : new()
{
    private static readonly ConcurrentDictionary<PropertyInfo, Action<TObject, object?>> PropertySetters = [];

    private static readonly ConcurrentDictionary<Type, Func<PropertyInfo, Action<TObject, object?>>> SetterFactories = new();

    private static readonly ConcurrentDictionary<Type, List<(PropertyInfo Property, ExcelColumnAttribute Attribute)>> CachedExcelColumnProperties = new();

    static ExcelObjectExtractor()
    {
        InitializeDefaultSetterFactories();
    }

    private static void InitializeDefaultSetterFactories()
    {
        // Default setters for common types
        SetterFactories[typeof(string)] = prop => (obj, value) => prop.SetValue(obj, value?.ToString());

        SetterFactories[typeof(int)] = prop => CreateValueSetter(prop, Convert.ToInt32);
        SetterFactories[typeof(int?)] = prop => CreateNullableValueSetter(prop, Convert.ToInt32);

        SetterFactories[typeof(double)] = prop => CreateValueSetter(prop, Convert.ToDouble);
        SetterFactories[typeof(double?)] = prop => CreateNullableValueSetter(prop, Convert.ToDouble);

        SetterFactories[typeof(decimal)] = prop => CreateValueSetter(prop, Convert.ToDecimal);
        SetterFactories[typeof(decimal?)] = prop => CreateNullableValueSetter(prop, Convert.ToDecimal);

        SetterFactories[typeof(DateTime)] = prop => CreateValueSetter(prop, Convert.ToDateTime);
        SetterFactories[typeof(DateTime?)] = prop => CreateNullableValueSetter(prop, Convert.ToDateTime);

        SetterFactories[typeof(bool)] = prop => CreateValueSetter(prop, Convert.ToBoolean);
        SetterFactories[typeof(bool?)] = prop => CreateNullableValueSetter(prop, Convert.ToBoolean);

        SetterFactories[typeof(TimeSpan)] = prop => CreateValueSetter(prop, value => TimeSpan.Parse(value.ToString()!));
        SetterFactories[typeof(TimeSpan?)] = prop => CreateNullableValueSetter(prop, value => TimeSpan.Parse(value.ToString()!));

        SetterFactories[typeof(Guid)] = prop => CreateValueSetter(prop, value => Guid.Parse(value.ToString()!));
        SetterFactories[typeof(Guid?)] = prop => CreateNullableValueSetter(prop, value => Guid.Parse(value.ToString()!));
    }
    
    public static void RegisterConverter<TProperty>(Func<PropertyInfo, Action<TObject, object?>> converter)
    {
        var propType = typeof(TProperty);

        if (!SetterFactories.TryAdd(propType, converter))
        {
            throw new InvalidOperationException($"A converter for type '{propType.Name}' is already registered.");
        }
    }

    public List<TObject> ExtractDataFromWorksheet(IXLWorksheet worksheet, bool readHeader)
    {
        var extractedData = new List<TObject>();
        var excelRange = worksheet.RangeUsed();
        if (excelRange is null) return extractedData;

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();
        var mappings = BuildAttributeMappings(worksheet, readHeader);

        foreach (var row in excelRows)
        {
            var obj = new TObject();
            foreach ((int colIndex, Action<TObject, object?> setter) in mappings)
            {
                try
                {
                    setter(obj, GetCellValue(row.Cell(colIndex)));
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Error setting value for column {colIndex} in row {row.RowNumber()}", ex);
                }
            }

            extractedData.Add(obj);
        }

        return extractedData;
    }

    public List<TObject> ExtractDataFromWorksheet(IXLWorksheet worksheet, Func<IXLRangeRow, TObject> mapRow,
        bool readHeader)
    {
        var excelRange = worksheet.RangeUsed();

        if (excelRange is null) return [];

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();

        return excelRows.Select(mapRow).ToList();
    }

    private static Dictionary<int, Action<TObject, object?>> BuildAttributeMappings(IXLWorksheet worksheet,
        bool readHeader)
    {
        var mappings = new Dictionary<int, Action<TObject, object?>>();
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

            foreach ((PropertyInfo propInfo, int columnIndex) in properties.Select((p, i) => (p, i + 1)))
            {
                mappings[columnIndex] = GetSetterForProperty(propInfo);
            }
        }

        return mappings;
    }


    private static Action<TObject, object?> GetSetterForProperty(PropertyInfo property)
    {
        return PropertySetters.GetOrAdd(property, CreateSetterForProperty);

        static Action<TObject, object?> CreateSetterForProperty(PropertyInfo property)
        {
            if (SetterFactories.TryGetValue(property.PropertyType, out var factory))
            {
                return factory(property);
            }
            else
            {
                //If target object properties do not have the type set in SetterFactories, resolve here.
                return CreateGenericSetter(property);
            }
        }
    }

    private static Action<TObject, object?> CreateValueSetter<TValue>(PropertyInfo property,
        Func<object, TValue> converter) where TValue : struct
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

    private static Action<TObject, object?> CreateNullableValueSetter<TValue>(PropertyInfo property,
        Func<object, TValue> converter) where TValue : struct
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


    private static Action<TObject, object?> CreateGenericSetter(PropertyInfo property)
    {
        var propertyType = property.PropertyType;
        var underlyingType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

        // Handle Enums here to resolve the correct type 
        if (underlyingType.IsEnum)
        {
            return (obj, value) =>
            {
                switch (value)
                {
                    case null:
                        property.SetValue(obj, propertyType.IsGenericType ? null : Activator.CreateInstance(propertyType));
                        break;
                    case string strValue when Enum.TryParse(underlyingType, strValue, true, out var result):
                        property.SetValue(obj, result);
                        break;
                    default:
                        // Set to default enum value
                        property.SetValue(obj, Enum.GetValues(underlyingType).GetValue(0)); 
                        break;
                }
            };
        }

        // This is a fallback for a generic setter in case no specific setter is found in SetterFactories
        var instance = Expression.Parameter(typeof(TObject), "instance");
        var value = Expression.Parameter(typeof(object), "value");


        var convertedValue = Expression.Convert(
            Expression.Call(typeof(Convert).GetMethod(nameof(Convert.ChangeType), [typeof(object), typeof(Type)])!,
                value, Expression.Constant(underlyingType)),
            propertyType);

        var assign = Expression.Assign(Expression.Property(instance, property), convertedValue);
        return Expression.Lambda<Action<TObject, object?>>(assign, instance, value).Compile();
    }

    private static object? GetCellValue(IXLCell cell)
    {
        var type = cell.Value.Type;

        if (type == XLDataType.Blank)
            return null;

        return type switch
        {
            XLDataType.Text => cell.GetString(),
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            XLDataType.Error => cell.GetError(),
            _ => null
        };
    }

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnAttributeProperties()
    {
        return CachedExcelColumnProperties.GetOrAdd(typeof(TObject), type =>
        {
            var propertiesWithAttributes = type.GetProperties()
                .Select(p => new { Property = p, Attribute = p.GetCustomAttribute<ExcelColumnAttribute>() })
                .Where(p => p.Attribute != null)
                .Select(p => (p.Property, p.Attribute!))
                .ToList();

            if (propertiesWithAttributes.Count == 0)
            {
                throw new InvalidOperationException($"No properties with {nameof(ExcelColumnAttribute)} found on type {type.Name}");
            }

            return propertiesWithAttributes;
        });
    }

    public static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetCachedExcelColumnProperties()
    {
        return GetExcelColumnAttributeProperties();
    }

    private static List<PropertyInfo> GetProperties()
    {
        var properties = typeof(TObject).GetProperties();
        return properties.Length > 0
            ? properties.ToList()
            : throw new InvalidOperationException($"No properties found on type {typeof(TObject).Name}");
    }
}