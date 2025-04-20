using System.Text.Json;
using System.Xml.Serialization;
using ClosedXML.Excel;

namespace XtractXcel;

public static class ExcelDataExtractorExtensions
{
    public static string SaveAsJson<T>(this List<T> data)
        where T : new()
    {
        return SerializeToJson(data);

        static string SerializeToJson(List<T> data)
        {
            using var stream = new MemoryStream();
            using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false });
            JsonSerializer.Serialize(writer, data);
            return System.Text.Encoding.UTF8.GetString(stream.ToArray());
        }
    }

    public static string SaveAsXml<T>(this List<T> data)
        where T : new()
    {
        return SerializeToXml(data);

        static string SerializeToXml(List<T> data)
        {
            using var stringWriter = new StringWriter();
            new XmlSerializer(typeof(List<T>)).Serialize(stringWriter, data);
            return stringWriter.ToString();
        }
    }

    public static void SaveAsXml<T>(this List<T> data, string filePath)
        where T : new()
    {
        using var writer = new StreamWriter(filePath);
        var serializer = new XmlSerializer(typeof(List<T>));
        serializer.Serialize(writer, data);
    }

    public static void SaveAsXlsx<T>(this List<T> data, string filePath)
        where T : new()
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
}