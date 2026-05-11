using System.Collections.Concurrent;
using System.Reflection;

namespace XtractXcel;

internal static class ExcelObjectExtractorCache
{
    private static readonly ConcurrentDictionary<Type, ConcurrentDictionary<PropertyInfo, object>> PropertySetters = new();
    private static readonly ConcurrentDictionary<Type, object> CachedExcelColumnProperties = new();

    internal static Action<T, object?> GetOrAddSetter<T>(PropertyInfo property, Func<PropertyInfo, Action<T, object?>> factory)
    {
        var typeSetters = PropertySetters.GetOrAdd(typeof(T), _ => new ConcurrentDictionary<PropertyInfo, object>());
        return (Action<T, object?>)typeSetters.GetOrAdd(property, p => factory((PropertyInfo)p));
    }

    internal static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetOrAddProperties(Type type, Func<Type, List<(PropertyInfo Property, ExcelColumnAttribute Attribute)>> factory)
    {
        return (List<(PropertyInfo Property, ExcelColumnAttribute Attribute)>)CachedExcelColumnProperties.GetOrAdd(type, t => factory((Type)t));
    }

    internal static void Clear()
    {
        PropertySetters.Clear();
        CachedExcelColumnProperties.Clear();
    }
}
