using ClosedXML.Excel;

namespace ExcelTransformLoad.UnitTests.TestHelpers;

public static class TestExcelGenerator
{
    public static MemoryStream CreateTestExcelFile()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            // Headers
            worksheet.Cell(1, 1).Value = "Full Name";
            worksheet.Cell(1, 2).Value = "Age";
            worksheet.Cell(1, 3).Value = "Salary";
            worksheet.Cell(1, 4).Value = "Join Date";
            worksheet.Cell(1, 5).Value = "Last Active";

            // Data rows
            worksheet.Cell(2, 1).Value = "Alice";
            worksheet.Cell(2, 2).Value = 25;
            worksheet.Cell(2, 3).Value = 50000.75;
            worksheet.Cell(2, 4).Value = new DateTime(2020, 5, 1);
            worksheet.Cell(2, 5).Clear();

            worksheet.Cell(3, 1).Value = "Bob";
            worksheet.Cell(3, 2).Clear();
            worksheet.Cell(3, 3).Clear();
            worksheet.Cell(3, 4).Value = new DateTime(2018, 10, 15);
            worksheet.Cell(3, 5).Value = new DateTime(2023, 3, 10);

            // Edge cases
            worksheet.Cell(4, 1).Value = "Charlie";
            worksheet.Cell(4, 2).Value = 0;
            worksheet.Cell(4, 3).Value = -100.50;
            worksheet.Cell(4, 4).Value = new DateTime(2022, 1, 1);
            worksheet.Cell(4, 5).Clear();

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        return stream;
    }

    public static MemoryStream CreateTestExcelFileWithNoHeader()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Sheet1");

            worksheet.Cell(1, 1).Value = "Dave";
            worksheet.Cell(1, 2).Value = 42;
            worksheet.Cell(1, 3).Value = 75000.50;
            worksheet.Cell(1, 4).Value = new DateTime(2019, 3, 15);
            worksheet.Cell(1, 5).Value = new DateTime(2024, 1, 10);

            worksheet.Cell(2, 1).Value = "Eve";
            worksheet.Cell(2, 2).Value = 38;
            worksheet.Cell(2, 3).Value = 82000.25;
            worksheet.Cell(2, 4).Value = new DateTime(2020, 7, 22);
            worksheet.Cell(2, 5).Value = new DateTime(2024, 2, 5);

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        return stream;
    }
}