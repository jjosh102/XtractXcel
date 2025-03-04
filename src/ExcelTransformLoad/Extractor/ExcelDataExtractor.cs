
using System.Collections.Concurrent;
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

internal sealed class ExcelDataExtractor : ExcelDataExtractorBase
{
    private static readonly ConcurrentDictionary<Type, object> _typedExtractorCache = new();

    public ExcelDataExtractor(ExcelDataSourceOptions options) : base(options) { }

    public List<T> ExtractData<T>(string worksheetName, bool readHeader = true) where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (!workbook.TryGetWorksheet(worksheetName, out var worksheet))
            throw new ArgumentException($"Worksheet '{worksheetName}' not found", nameof(worksheetName));

        var typedExtractor = GetOrCreateTypedExtractor<T>();
        return typedExtractor.ExtractDataFromWorksheet(worksheet, readHeader);
    }

    public List<T> ExtractData<T>(string worksheetName, Func<IXLRangeRow, T> mapRow, bool readHeader = true) where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (!workbook.TryGetWorksheet(worksheetName, out var worksheet))
            throw new ArgumentException($"Worksheet '{worksheetName}' not found", nameof(worksheetName));

        var typedExtractor = GetOrCreateTypedExtractor<T>();
        return typedExtractor.ExtractDataFromWorksheet(worksheet, mapRow, readHeader);
    }

    public List<T> ExtractData<T>(int worksheetIndex, bool readHeader = true) where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (worksheetIndex < 1 || worksheetIndex > workbook.Worksheets.Count)
            throw new ArgumentOutOfRangeException(nameof(worksheetIndex),
                $"Worksheet index must be between 1 and {workbook.Worksheets.Count}");

        var worksheet = workbook.Worksheet(worksheetIndex);
        var typedExtractor = GetOrCreateTypedExtractor<T>();
        return typedExtractor.ExtractDataFromWorksheet(worksheet, readHeader);
    }

    public List<T> ExtractData<T>(int worksheetIndex, Func<IXLRangeRow, T> mapRow, bool readHeader = true) where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (worksheetIndex < 1 || worksheetIndex > workbook.Worksheets.Count)
            throw new ArgumentOutOfRangeException(nameof(worksheetIndex),
                $"Worksheet index must be between 1 and {workbook.Worksheets.Count}");

        var worksheet = workbook.Worksheet(worksheetIndex);
        var typedExtractor = GetOrCreateTypedExtractor<T>();
        return typedExtractor.ExtractDataFromWorksheet(worksheet, mapRow, readHeader);
    }

    private TypedWorksheetExtractor<T> GetOrCreateTypedExtractor<T>() where T : new()
    {
        // Reuse TypedWorksheetExtractor for the same type, some cases involve using the same type on a different worksheet
        return (TypedWorksheetExtractor<T>)_typedExtractorCache.GetOrAdd(typeof(T), _ => new TypedWorksheetExtractor<T>());
    }
}




