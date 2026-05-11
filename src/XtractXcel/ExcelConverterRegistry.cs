using System.Collections.Concurrent;
using System.Reflection;

namespace XtractXcel;

public static class ExcelConverterRegistry
{
    private static readonly ConcurrentDictionary<Type, Func<PropertyInfo, Action<object, object?>>> Converters = new();

    public static void RegisterConverter<T>(Func<PropertyInfo, Action<object, object?>> converter)
    {
        Converters[typeof(T)] = converter;
    }

    public static void Clear()
    {
        Converters.Clear();
    }

    internal static Action<object, object?>? GetConverter(Type type, PropertyInfo property)
    {
        if (Converters.TryGetValue(type, out var factory))
        {
            return factory(property);
        }

        return null;
    }
}
