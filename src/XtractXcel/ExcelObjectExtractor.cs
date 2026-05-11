using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace XtractXcel;

internal sealed class ExcelObjectExtractor<TObject>
    where TObject : new()
{
    internal static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetCachedExcelColumnProperties()
    {
        return GetExcelColumnAttributeProperties();
    }

    internal static void ClearCache()
    {
        ExcelObjectExtractorCache.Clear();
    }

    internal List<TObject> ExtractDataFromWorksheet(IXLWorksheet worksheet, bool readHeader, Action<TObject>? onRowExtracted = null)
    {
        var extractedData = new List<TObject>();
        var excelRange = worksheet.RangeUsed();
        if (excelRange is null)
        {
            return extractedData;
        }

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();
        var (mappings, columnNames) = BuildAttributeMappings(worksheet, readHeader);

        foreach (var row in excelRows)
        {
            var obj = new TObject();
            foreach (var mapping in mappings)
            {
                int colIndex = mapping.Key;
                Action<TObject, object?> setter = mapping.Value;
                object? cellValue = null;
                try
                {
                    cellValue = GetCellValue(row.Cell(colIndex));
                    setter(obj, cellValue);
                }
                catch (Exception ex)
                {
                    columnNames.TryGetValue(colIndex, out var headerName);
                    throw new ExcelExtractionException(
                        $"Error setting value for column {colIndex} {(headerName != null ? $"(Header: '{headerName}')" : string.Empty)} in row {row.RowNumber()}. Value: '{cellValue}'",
                        row.RowNumber(),
                        colIndex,
                        headerName,
                        cellValue,
                        ex);
                }
            }

            onRowExtracted?.Invoke(obj);
            extractedData.Add(obj);
        }

        return extractedData;
    }

    internal List<TObject> ExtractDataFromWorksheet(IXLWorksheet worksheet, Func<IXLRangeRow, TObject> mapRow, bool readHeader, Action<TObject>? onRowExtracted = null)
    {
        var excelRange = worksheet.RangeUsed();

        if (excelRange is null)
        {
            return [];
        }

        var excelRows = readHeader ? excelRange.RowsUsed().Skip(1) : excelRange.RowsUsed();

        var result = new List<TObject>();
        foreach (var row in excelRows)
        {
            var obj = mapRow(row);
            onRowExtracted?.Invoke(obj);
            result.Add(obj);
        }

        return result;
    }

    private static (Dictionary<int, Action<TObject, object?>> Mappings, Dictionary<int, string> ColumnNames) BuildAttributeMappings(
        IXLWorksheet worksheet,
        bool readHeader)
    {
        var mappings = new Dictionary<int, Action<TObject, object?>>();
        var columnNames = worksheet.Row(1).CellsUsed()
            .ToDictionary(c => c.Address.ColumnNumber, c => c.GetString());
        
        var nameToIndices = columnNames.ToDictionary(kvp => kvp.Value, kvp => kvp.Key, StringComparer.OrdinalIgnoreCase);

        if (readHeader)
        {
            foreach (var propInfo in GetExcelColumnAttributeProperties())
            {
                foreach (var columnName in propInfo.Attribute.ColumnNames)
                {
                    if (nameToIndices.TryGetValue(columnName, out int colIndex))
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

        return (mappings, columnNames);
    }

    private static Action<TObject, object?> GetSetterForProperty(PropertyInfo property)
    {
        return ExcelObjectExtractorCache.GetOrAddSetter<TObject>(property, CreateSetterForProperty);
    }

    private static Action<TObject, object?> CreateSetterForProperty(PropertyInfo property)
    {
        var propertyType = property.PropertyType;
        var underlyingType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

        // Check for custom converters first
        var customConverter = ExcelConverterRegistry.GetConverter(underlyingType, property);
        if (customConverter != null)
        {
            return (obj, val) => customConverter(obj!, val);
        }

        var instanceParam = Expression.Parameter(typeof(TObject), "instance");
        var valueParam = Expression.Parameter(typeof(object), "value");

        Expression valueExpr;

        if (propertyType == typeof(string))
        {
            // (obj, val) => obj.Prop = val?.ToString()
            valueExpr = Expression.Condition(
                Expression.Equal(valueParam, Expression.Constant(null)),
                Expression.Constant(null, typeof(string)),
                Expression.Call(valueParam, typeof(object).GetMethod("ToString")!));
        }
        else if (underlyingType.IsEnum)
        {
            // Special handling for Enums
            return CreateEnumSetter(property, underlyingType, propertyType);
        }
        else if (underlyingType == typeof(Guid))
        {
            valueExpr = CreateConversionExpr(valueParam, underlyingType, val => Guid.Parse(val.ToString()!));
        }
        else if (underlyingType == typeof(TimeSpan))
        {
            valueExpr = CreateConversionExpr(valueParam, underlyingType, val =>
                val is TimeSpan ? (TimeSpan)val : TimeSpan.Parse(val.ToString()!));
        }
        else
        {
            // Use Convert.ChangeType for other basic types
            var changeTypeMethod = typeof(Convert).GetMethod(nameof(Convert.ChangeType), [typeof(object), typeof(Type)])!;
            valueExpr = Expression.Call(changeTypeMethod, valueParam, Expression.Constant(underlyingType));
            valueExpr = Expression.Convert(valueExpr, underlyingType);
        }

        Expression finalValueExpr;
        if (propertyType.IsValueType && Nullable.GetUnderlyingType(propertyType) == null)
        {
            finalValueExpr = Expression.Condition(
                Expression.Equal(valueParam, Expression.Constant(null)),
                Expression.Default(propertyType),
                Expression.Convert(valueExpr, propertyType));
        }
        else
        {
            finalValueExpr = Expression.Condition(
                Expression.Equal(valueParam, Expression.Constant(null)),
                Expression.Constant(null, propertyType),
                Expression.Convert(valueExpr, propertyType));
        }

        var assignExpr = Expression.Assign(Expression.Property(instanceParam, property), finalValueExpr);
        return Expression.Lambda<Action<TObject, object?>>(assignExpr, instanceParam, valueParam).Compile();
    }

    private static Expression CreateConversionExpr(ParameterExpression valueParam, Type targetType, Expression<Func<object, object>> converter)
    {
        return Expression.Convert(Expression.Invoke(converter, valueParam), targetType);
    }

    private static Action<TObject, object?> CreateEnumSetter(PropertyInfo property, Type enumType, Type propertyType)
    {
        return (obj, value) =>
        {
            if (value == null)
            {
                property.SetValue(obj, propertyType.IsGenericType ? null : Activator.CreateInstance(propertyType));
                return;
            }

            if (value is string strValue && Enum.TryParse(enumType, strValue, true, out var result))
            {
                property.SetValue(obj, result);
            }
            else if (value is double dblValue)
            {
                property.SetValue(obj, Enum.ToObject(enumType, Convert.ToInt64(dblValue)));
            }
            else if (value is int intValue)
            {
                property.SetValue(obj, Enum.ToObject(enumType, intValue));
            }
            else
            {
                property.SetValue(obj, Enum.GetValues(enumType).GetValue(0));
            }
        };
    }

    private static object? GetCellValue(IXLCell cell)
    {
        var value = cell.CachedValue; 

        if (value.IsBlank)
        {
            return null;
        }

        if (value.IsText)
        {
            return value.GetText();
        }

        if (value.IsDateTime)
        {
            return value.GetDateTime();
        }

        if (value.IsNumber)
        {
            return value.GetNumber();
        }

        if (value.IsBoolean)
        {
            return value.GetBoolean();
        }

        if (value.IsTimeSpan)
        {
            return value.GetTimeSpan();
        }

        if (value.IsError)
        {
            return value.GetError();
        }

        return null;
    }

    private static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetExcelColumnAttributeProperties()
    {
        return ExcelObjectExtractorCache.GetOrAddProperties(typeof(TObject), type =>
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

    private static List<PropertyInfo> GetProperties()
    {
        var properties = typeof(TObject).GetProperties();
        return properties.Length > 0
            ? [.. properties]
            : throw new InvalidOperationException($"No properties found on type {typeof(TObject).Name}");
    }
}
