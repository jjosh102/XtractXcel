using System.Text.Json;
using System.Xml.Serialization;
using ClosedXML.Excel;

namespace XtractXcel;

public static class ExcelDataExtractorExtensions
{
    public static string SaveAsJson<T>(this List<T> data)
        where T : new()
    {
        return data is null ? throw new ArgumentNullException(nameof(data), "Data cannot be null.") : SerializeToJson(data);

        static string SerializeToJson(List<T> data)
        {
            try
            {
                using var stream = new MemoryStream();
                using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false });
                JsonSerializer.Serialize(writer, data);
                return System.Text.Encoding.UTF8.GetString(stream.ToArray());
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to serialize data to JSON.", ex);
            }
        }
    }

    public static string SaveAsXml<T>(this List<T> data)
        where T : new()
    {
        return data is null ? throw new ArgumentNullException(nameof(data), "Data cannot be null.") : SerializeToXml(data);

        static string SerializeToXml(List<T> data)
        {
            try
            {
                using var stringWriter = new StringWriter();
                new XmlSerializer(typeof(List<T>)).Serialize(stringWriter, data);
                return stringWriter.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to serialize data to XML.", ex);
            }
        }
    }

    public static void SaveAsXml<T>(this List<T> data, string filePath)
        where T : new()
    {
        if (data is null)
        {
            throw new ArgumentNullException(nameof(data), "Data cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        try
        {
            using var writer = new StreamWriter(filePath);
            var serializer = new XmlSerializer(typeof(List<T>));
            serializer.Serialize(writer, data);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save data as XML.", ex);
        }
    }

    public static void SaveAsXlsx<T>(this List<T> data, string filePath)
        where T : new()
    {
        if (data is null)
        {
            throw new ArgumentNullException(nameof(data), "Data cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add();

            var propertiesWithAttributes = ExcelDataExtractor.GetCachedExcelColumnProperties<T>();

            for (int i = 0; i < propertiesWithAttributes.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = propertiesWithAttributes[i].Attribute.ColumnNames.FirstOrDefault() ?? propertiesWithAttributes[i].Property.Name;
            }

            // Add data starting from cell A2, below the header
            worksheet.Cell(2, 1).InsertData(data);
            workbook.SaveAs(filePath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save data as XLSX.", ex);
        }
    }

    public static void SaveAsXlsxWithoutHeader<T>(this List<T> data, string filePath)
        where T : new()
    {
        if (data is null)
        {
            throw new ArgumentNullException(nameof(data), "Data cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add();

            // Add data starting from cell A1 without headers
            worksheet.Cell(1, 1).InsertData(data);
            workbook.SaveAs(filePath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save data as XLSX without headers.", ex);
        }
    }
}