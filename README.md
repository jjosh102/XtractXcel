# XtractXcel

[![NuGet](https://img.shields.io/nuget/v/XtractXcel.svg)](https://www.nuget.org/packages/XtractXcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/XtractXcel?logo=nuget)](https://www.nuget.org/packages/XtractXcel)

## Overview

XtractXcel is a simple library for extracting data from Excel files using [ClosedXML](https://github.com/ClosedXML/ClosedXML), transforming it as needed, and loading it into your objects with minimal effort. It supports a variety of data types, flexible column mapping, and both attribute-based and manual mapping approaches.

## Getting Started

### Installing

To install the package add the following line inside your csproj file with the latest version.

```xml
<PackageReference Include="XtractXcel" Version="x.x.x" />
```

An alternative is to install via the .NET CLI with the following command:

```xml
dotnet add package XtractXcel
```

### Basic Usage

#### 1. Define your model with ExcelColumn attributes

```csharp
public class Person {
    [ExcelColumn("Full Name", "Name", "Employee Name")]
    public string? Name { get; set; }
    
    [ExcelColumn("Age", "Employee Age")]
    public int? Age { get; set; }
    
    [ExcelColumn("Salary")]
    public decimal? Salary { get; set; }
    
    [ExcelColumn("Join Date")]
    public DateTime JoinDate { get; set; }
    
    [ExcelColumn("Last Active", "Last Activity")]
    public DateTime? LastActive { get; set; }
}
```

The `ExcelColumn` attribute maps Excel column headers to C# properties. You can provide multiple possible column names to handle variations in your Excel files gracefully.

#### 2. Extract data from Excel

##### From a Stream

```csharp
// Get a stream from a file, memory, or any other source
using var stream = File.OpenRead("employees.xlsx");

// Extract the data using a fluent API
var people = new ExcelExtractor()
    .WithHeader(true)               // Excel file contains headers
    .WithWorksheetIndex(1)          // Use the first worksheet (1-based index)
    .FromStream(stream)             // Set the source stream
    .Extract<Person>();             // Perform the extraction

// Use the extracted data
foreach (var person in people) {
    Console.WriteLine($"Name: {person.Name}, Age: {person.Age}, Joined: {person.JoinDate:d}");
}
```

##### From a File

```csharp
var people = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromFile("employees.xlsx")
    .Extract<Person>();
```

## Advanced Features

### Working with Files Without Headers

For Excel files without headers, you can use column position for extraction:

```csharp
public class PersonNoHeader {
    // No attributes needed - properties are mapped by column position (1-based)
    // First column (A) maps to first property, second column (B) to second property, etc.
    public string? Name { get; set; }
    public int? Age { get; set; }
    public decimal? Salary { get; set; }
    public DateTime JoinDate { get; set; }
    public DateTime? LastActive { get; set; }
}

// Extract the data
var people = new ExcelExtractor()
    .WithHeader(false)              // Specify that there's no header row
    .WithWorksheetIndex(1)
    .FromFile("employees-no-header.xlsx")
    .Extract<PersonNoHeader>();
```

### Manual Mapping

For more control over the extraction process, you can use manual mapping:

```csharp
// Extract data with manual mapping
var people = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream)
    .ExtractWithManualMapping(row => new Person {
        Name = row.Cell(1).GetString(),
        Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
        Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
        JoinDate = row.Cell(4).GetDateTime(),
        LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
    });
```

### Supported Data Types

`XtractXcel` supports a wide range of data types:

- Basic types: `string`, `int`, `decimal`, `double`, `DateTime`
- Nullable variants: `int?`, `decimal?`, `DateTime?`, etc.
- `TimeSpan` for time values
- `Guid` for unique identifiers
- Enums for categorized data

### Selecting Specific Columns

If you only need certain columns from an Excel file:

```csharp
public class PersonWithSpecificColumns {
    [ExcelColumn("Name")]
    public string? NameOnly { get; set; }
    
    [ExcelColumn("Salary")]
    public decimal SalaryOnly { get; set; }
}

var partialData = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromFile("employees.xlsx")
    .Extract<PersonWithSpecificColumns>();
```

### Handling Enums

Enums are supported out of the box:

```csharp
public enum UserStatus {
    None,
    Active,
    Inactive,
    Suspended
}

public class PersonWithEnumStatus {
    [ExcelColumn("Name")]
    public string? Name { get; set; }
    
    [ExcelColumn("Status")]
    public UserStatus Status { get; set; }
}

var people = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromFile("employees.xlsx")
    .Extract<PersonWithEnumStatus>();
```

### Data Transformation During Extraction

Transform data as it's being extracted:

```csharp
var transformedData = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream)
    .ExtractWithManualMapping(row => new Person {
        // Convert names to uppercase
        Name = row.Cell(1).GetString().ToUpper(),
        // Double the age values
        Age = !row.Cell(2).IsEmpty() ? (int)(row.Cell(2).GetDouble() * 2) : null,
        // Halve the salary values
        Salary = !row.Cell(3).IsEmpty() ? (decimal)(row.Cell(3).GetDouble() / 2) : null,
        // Add a year to join dates
        JoinDate = row.Cell(4).GetDateTime().AddYears(1),
        // Use current date for missing activity dates
        LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : DateTime.Now
    });
```

### Converting to Different Target Types

You can map Excel data to any object type:

```csharp
public class CustomPerson {
    public string? FullName { get; set; }
    public int YearsOld { get; set; }
    public decimal AnnualSalary { get; set; }
    public DateTime StartDate { get; set; }
    public bool IsActive { get; set; }
}

var customData = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream)
    .ExtractWithManualMapping(row => new CustomPerson {
        FullName = row.Cell(1).GetString(),
        YearsOld = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : 0,
        AnnualSalary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : 0,
        StartDate = row.Cell(4).GetDateTime(),
        IsActive = !row.Cell(5).IsEmpty()
    });
```

### Saving Extracted Data

The `ExcelExtractor` class now supports saving extracted data into various formats using extension methods. Here are some examples:

#### Save as JSON

```csharp
using var stream = File.OpenRead("employees.xlsx");
var extractor = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream);

string json = extractor.Extract<Person>().SaveAsJson();
Console.WriteLine(json);
```

#### Save as XML

```csharp
using var stream = File.OpenRead("employees.xlsx");
var extractor = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream);

string xml = extractor.Extract<Person>().SaveAsXml();
Console.WriteLine(xml);
```

#### Save as XLSX

```csharp
using var stream = File.OpenRead("employees.xlsx");
var extractor = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream);

extractor.Extract<Person>().SaveAsXlsx("output.xlsx");
Console.WriteLine("Data saved to output.xlsx");
```

#### Save as XLSX Without Headers

If your data does not have headers, you can save it directly into an XLSX file without adding headers:

```csharp
using var stream = File.OpenRead("employees-no-header.xlsx");
var extractor = new ExcelExtractor()
    .WithHeader(false)
    .WithWorksheetIndex(1)
    .FromStream(stream);

var data = extractor.Extract<PersonNoHeader>();
data.SaveAsXlsxWithoutHeader("output-no-header.xlsx");
Console.WriteLine("Data saved to output-no-header.xlsx without headers");
```

#### Save Manually Mapped Data as XLSX

```csharp
using var stream = File.OpenRead("employees.xlsx");
var extractor = new ExcelExtractor()
    .WithHeader(true)
    .WithWorksheetIndex(1)
    .FromStream(stream);

var manuallyMappedData = extractor.ExtractWithManualMapping(row => new Person
{
    Name = row.Cell(1).GetString()?.ToUpper(), // Convert names to uppercase
    Age = !row.Cell(2).IsEmpty() ? (int)(row.Cell(2).GetDouble() * 2) : null, // Double the age
    Salary = !row.Cell(3).IsEmpty() ? (decimal)(row.Cell(3).GetDouble() / 2) : null, // Halve the salary
    JoinDate = row.Cell(4).GetDateTime().AddYears(1), // Add a year to join dates
    LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : DateTime.Now // Use current date for missing activity dates
});

manuallyMappedData.SaveAsXlsx("manually_mapped_output.xlsx");
Console.WriteLine("Manually mapped data saved to manually_mapped_output.xlsx");
```

## Performance Considerations

Benchmark results show that both attribute-based and manual mapping perform well, but manual mapping has a slight edge in certain cases.

| Method                               | Mean       | Error     | StdDev    | Gen0      | Gen1      | Gen2      | Allocated |
|-------------------------------------|-----------:|----------:|----------:|----------:|----------:|----------:|----------:|
| SmallFile_AttributeMapping           |   2.685 ms | 0.0504 ms | 0.0539 ms |  148.4375 |   46.8750 |         - |   1.88 MB |
| SmallFile_ManualMapping              |   2.600 ms | 0.0423 ms | 0.0375 ms |  148.4375 |   46.8750 |         - |   1.86 MB |
| SmallFile_ManualMapping_NoAttributes |   2.613 ms | 0.0183 ms | 0.0171 ms |  148.4375 |   46.8750 |         - |   1.86 MB |
| MediumFile_AttributeMapping          |  16.793 ms | 0.3305 ms | 0.4740 ms | 1000.0000 |  545.4545 |   90.9091 |  13.79 MB |
| MediumFile_ManualMapping             |  16.172 ms | 0.3170 ms | 0.4647 ms | 1000.0000 |  454.5455 |   90.9091 |  13.66 MB |
| LargeFile_AttributeMapping           | 174.138 ms | 3.4191 ms | 3.9374 ms | 9000.0000 | 5000.0000 | 2000.0000 | 130.90 MB |
| LargeFile_ManualMapping              | 169.002 ms | 3.3734 ms | 5.4474 ms | 9000.0000 | 5000.0000 | 2000.0000 | 129.48 MB |
| ManyColumns_AttributeMapping         |  27.434 ms | 0.5463 ms | 0.9711 ms | 1500.0000 |  750.0000 |  250.0000 |  18.69 MB |
| ManyColumns_ManualMapping            |  27.409 ms | 0.5114 ms | 1.0095 ms | 1375.0000 |  625.0000 |  250.0000 |  18.65 MB |

## Why Use XtractXcel?

If you're already using [ClosedXML](https://github.com/ClosedXML/ClosedXML) or similar libraries extensively, this one might not add much extra value. But if you're looking for a simple way to read an Excel file and load it into your objects without any hassle, this might worth checking out .
