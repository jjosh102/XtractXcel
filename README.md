# ExcelTransformLoad

## Overview
ExcelTransformLoad is a robust .NET library for extracting data from Excel files using [ClosedXML](https://github.com/ClosedXML/ClosedXML), transforming it as needed, and loading it into your objects with minimal effort. It supports a variety of data types, flexible column mapping, and both attribute-based and manual mapping approaches.

## Getting Started

### Installation (Coming soon)


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

ExcelTransformLoad supports a wide range of data types:

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

## Performance Considerations

Based on benchmarks, both attribute-based and manual mapping provide good performance:

| Method                               | Mean       | Error     | StdDev    | Gen0      | Gen1      | Gen2      | Allocated |
|------------------------------------- |-----------:|----------:|----------:|----------:|----------:|----------:|----------:|
| SmallFile_AttributeMapping           |   3.330 ms | 0.0482 ms | 0.0403 ms |  140.6250 |   46.8750 |         - |   1.89 MB |
| SmallFile_ManualMapping              |   2.740 ms | 0.0197 ms | 0.0154 ms |  148.4375 |   46.8750 |         - |   1.86 MB |
| SmallFile_ManualMapping_NoAttributes |   2.766 ms | 0.0550 ms | 0.0540 ms |  148.4375 |   46.8750 |         - |   1.86 MB |
| MediumFile_AttributeMapping          |  16.327 ms | 0.3255 ms | 0.5615 ms | 1000.0000 |  727.2727 |   90.9091 |  13.66 MB |
| MediumFile_ManualMapping             |  15.806 ms | 0.3136 ms | 0.5492 ms | 1000.0000 |  700.0000 |  100.0000 |  13.67 MB |
| LargeFile_AttributeMapping           | 177.912 ms | 3.4578 ms | 4.8473 ms | 9000.0000 | 4000.0000 | 2000.0000 | 129.31 MB |
| LargeFile_ManualMapping              | 183.702 ms | 3.1083 ms | 2.7555 ms | 9000.0000 | 5000.0000 | 2000.0000 | 129.61 MB |
| ManyColumns_AttributeMapping         |  27.877 ms | 0.5434 ms | 0.5815 ms | 1444.4444 |  888.8889 |  222.2222 |  18.74 MB |
| ManyColumns_ManualMapping            |  27.434 ms | 0.4533 ms | 0.4241 ms | 1444.4444 |  888.8889 |  222.2222 |  18.69 MB |

Manual mapping provides a slight performance edge for small files, while both approaches perform similarly for larger datasets.

## Why Use ExcelTransformLoad?
If you're already using [ClosedXML](https://github.com/ClosedXML/ClosedXML) or similar libraries extensively, this one might not add much extra value. But if you're looking for a simple way to read an Excel file and load it into your objects without any hassle, this library is worth checking out!

It's user-friendly and follows a fluent pattern, making it easy to define your options in a natural, intuitive way.