using System.Reflection;
using ClosedXML.Excel;

namespace XtractXcel;

internal sealed class ExcelDataExtractor : ExcelDataExtractorBase
{
    internal ExcelDataExtractor(ExcelDataSourceOptions options)
        : base(options)
    {
    }

    internal static List<(PropertyInfo Property, ExcelColumnAttribute Attribute)> GetCachedExcelColumnProperties<T>()
            where T : new()
    {
        return ExcelObjectExtractor<T>.GetCachedExcelColumnProperties();
    }

    internal List<T> ExtractData<T>(string worksheetName, bool readHeader = true)
        where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (!workbook.TryGetWorksheet(worksheetName, out var worksheet))
        {
            throw new ArgumentException($"Worksheet '{worksheetName}' not found", nameof(worksheetName));
        }

        return new ExcelObjectExtractor<T>().ExtractDataFromWorksheet(worksheet, readHeader);
    }

    internal List<T> ExtractData<T>(string worksheetName, Func<IXLRangeRow, T> mapRow, bool readHeader = true)
        where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (!workbook.TryGetWorksheet(worksheetName, out var worksheet))
        {
            throw new ArgumentException($"Worksheet '{worksheetName}' not found", nameof(worksheetName));
        }

        return new ExcelObjectExtractor<T>().ExtractDataFromWorksheet(worksheet, mapRow, readHeader);
    }

    internal List<T> ExtractData<T>(int worksheetIndex, bool readHeader = true)
        where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (worksheetIndex < 1 || worksheetIndex > workbook.Worksheets.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(worksheetIndex),
                $"Worksheet index must be between 1 and {workbook.Worksheets.Count}");
        }

        var worksheet = workbook.Worksheet(worksheetIndex);
        return new ExcelObjectExtractor<T>().ExtractDataFromWorksheet(worksheet, readHeader);
    }

    internal List<T> ExtractData<T>(int worksheetIndex, Func<IXLRangeRow, T> mapRow, bool readHeader = true)
        where T : new()
    {
        var workbook = GetOrCreateWorkbook();
        if (worksheetIndex < 1 || worksheetIndex > workbook.Worksheets.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(worksheetIndex),
                $"Worksheet index must be between 1 and {workbook.Worksheets.Count}");
        }

        var worksheet = workbook.Worksheet(worksheetIndex);
        return new ExcelObjectExtractor<T>().ExtractDataFromWorksheet(worksheet, mapRow, readHeader);
    }
}