using System.Text.Json;
using System.Xml.Serialization;
using ClosedXML.Excel;

namespace ExcelTransformLoad.Extractor;

public record ExcelExtractor(
    string? FilePath = null,
    Stream? Stream = null,
    bool ReadHeader = true,
    int? WorksheetIndex = null,
    string? WorksheetName = null
)
{
    public ExcelExtractor WithHeader(bool readHeader)
    {
        EnsureSourceNotSet();
        return this with { ReadHeader = readHeader };
    }


    public ExcelExtractor WithWorksheetIndex(int workSheetIndex)
    {
        EnsureSourceNotSet();
        return this with { WorksheetIndex = workSheetIndex, WorksheetName = null };
    }

    public ExcelExtractor WithWorksheetName(string workSheetName)
    {
        EnsureSourceNotSet();
        return this with { WorksheetName = workSheetName, WorksheetIndex = null };
    }

    public ExcelExtractor FromFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        EnsureSourceNotSet();
        return this with { FilePath = filePath, Stream = null };
    }

    public ExcelExtractor FromStream(Stream stream)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream), "Stream cannot be null");

        EnsureSourceNotSet();
        return this with { Stream = stream, FilePath = null };
    }

    public List<T> Extract<T>() where T : new()
    {
        EnsureSourceIsSet();

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };

        if (WorksheetIndex.HasValue)
        {
            return new ExcelDataExtractor(options).ExtractData<T>(WorksheetIndex.Value, ReadHeader);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return new ExcelDataExtractor(options).ExtractData<T>(WorksheetName, ReadHeader);
        }

        return new ExcelDataExtractor(options).ExtractData<T>(1, ReadHeader);
    }

    public List<T> ExtractWithManualMapping<T>(Func<IXLRangeRow, T> manualMapping) where T : new()
    {
        EnsureSourceIsSet();

        if (manualMapping is null)
        {
            throw new InvalidOperationException(
                "A row mapping function must be provided when manual mapping is enabled.");
        }

        var options = new ExcelDataSourceOptions { FilePath = FilePath, Stream = Stream };

        if (WorksheetIndex.HasValue)
        {
            return new ExcelDataExtractor(options).ExtractData(WorksheetIndex.Value, manualMapping, ReadHeader);
        }

        if (!string.IsNullOrEmpty(WorksheetName))
        {
            return new ExcelDataExtractor(options).ExtractData(WorksheetName, manualMapping, ReadHeader);
        }

        return new ExcelDataExtractor(options).ExtractData(1, manualMapping, ReadHeader);
    }

    public string ExtractAsJson<T>() where T : new()
    {
        EnsureSourceIsSet();

        return SerializeToJson(Extract<T>());

        static string SerializeToJson(List<T> data)
        {
            using var stream = new MemoryStream();
            using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false });
            JsonSerializer.Serialize(writer, data);
            return System.Text.Encoding.UTF8.GetString(stream.ToArray());
        }
    }

    public string ExtractAsXml<T>() where T : new()
    {
        EnsureSourceIsSet();

        return SerializeToXml(Extract<T>());

        static string SerializeToXml(List<T> data)
        {
            using var stringWriter = new StringWriter();
            new XmlSerializer(typeof(List<T>)).Serialize(stringWriter, data);
            return stringWriter.ToString();
        }
    }


    private void EnsureSourceNotSet()
    {
        if (FilePath is not null || Stream is not null)
            throw new InvalidOperationException(
                "Source (file or stream) has already been set. Cannot modify settings after source is set.");
    }

    private void EnsureSourceIsSet()
    {
        if (FilePath is null && Stream is null)
            throw new InvalidOperationException("Data source (file or stream) is required before extraction.");
    }

}
