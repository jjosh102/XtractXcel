# ExcelTransformLoad (In-progress)

## Overview
ExcelTransformLoad is a simple .NET library for extracting data from Excel files using [ClosedXML](https://github.com/ClosedXML/ClosedXML), transforming it as needed, and loading it into your object.

## Getting Started

### Installation (Coming soon)


### Basic Usage

#### 1. Define your model with ExcelColumn attributes
```csharp
public class Person {
    [ExcelColumn("Full Name", "Name", "Employee Name")]
    public string? Name { get; init; }
    
    [ExcelColumn("Age", "Employee Age")]
    public int? Age { get; init; }
    
    [ExcelColumn("Salary")]
    public decimal? Salary { get; init; }
    
    [ExcelColumn("Join Date")]
    public DateTime JoinDate { get; init; }
    
    [ExcelColumn("Last Active", "Last Activity")]
    public DateTime? LastActive { get; init; }
}
```

The `ExcelColumn` attribute maps Excel column headers to C# properties. Ever had someone from HR send you reports where column names change daily? "EMPLOYEE NAME" on Monday, "Employee Name" on Tuesday, and "employee_name" on Wednesday? No problem! You can map multiple possible column names to handle these situations gracefully.

#### 2. Extract data from Excel

##### From a Stream
```csharp
// Get a stream from a file, memory, or any other source
using var stream = File.OpenRead("employees.xlsx");

// Extract the data using a fluent API
var people = new ExcelExtractor<Person>()
    .WithHeader(true)               // Excel file contains headers
    .WithSheetIndex(1)              // Use the first sheet (1-based index)
    .FromStream(stream)             // Set the source stream
    .Extract();                     // Perform the extraction

// Use the extracted data
foreach (var person in people) {
    Console.WriteLine($"Name: {person.Name}, Age: {person.Age}, Joined: {person.JoinDate:d}");
}
```

##### From a File
```csharp
var people = new ExcelExtractor<Person>()
    .WithHeader(true)
    .WithSheetIndex(1)
    .FromFile("employees.xlsx")
    .Extract();
```

### Advanced Features

#### Working with Files Without Headers
For Excel files without headers, you can use the column indices instead:

```csharp
// Define a model without header mapping
public class PersonNoHeader {
    // Properties will be mapped by column position (1-based)
    public string? Name { get; init; }
    public int? Age { get; init; }
    public decimal? Salary { get; init; }
    public DateTime JoinDate { get; init; }
    public DateTime? LastActive { get; init; }
}

// Extract the data
var people = new ExcelExtractor<PersonNoHeader>()
    .WithHeader(false)              // Specify that there's no header row
    .WithSheetIndex(1)
    .FromFile("employees-no-header.xlsx")
    .Extract();
```

#### Manual Mapping
For more control over the extraction process, you can use manual mapping with a custom function:

```csharp
// Extract data with manual mapping
var people = new ExcelExtractor<Person>()
    .WithHeader(true)
    .WithSheetIndex(1)
    .WithManualMapping(row => new Person
    {
        Name = row.Cell(1).GetString(),
        Age = !row.Cell(2).IsEmpty() ? (int)row.Cell(2).GetDouble() : null,
        Salary = !row.Cell(3).IsEmpty() ? (decimal)row.Cell(3).GetDouble() : null,
        JoinDate = row.Cell(4).GetDateTime(),
        LastActive = !row.Cell(5).IsEmpty() ? row.Cell(5).GetDateTime() : null
    })
    .FromFile("employees.xlsx")
    .Extract();
```

## Examples

### Processing Employee Data with Headers
```csharp
// Read employee data with headers
var employees = new ExcelExtractor<Person>()
    .WithHeader(true)
    .WithSheetIndex(1)
    .FromFile("employees.xlsx")
    .Extract();

// Calculate average salary
var averageSalary = employees
    .Where(e => e.Salary.HasValue)
    .Average(e => e.Salary.Value);
```

### Processing Raw Data Without Headers
```csharp
// Read raw data without headers
var data = new ExcelExtractor<PersonNoHeader>()
    .WithHeader(false)
    .WithSheetIndex(1)
    .FromFile("data.xlsx")
    .Extract();
```

### Using Manual Mapping for Selective Column Reading
```csharp
// Read only specific columns
var partialData = new ExcelExtractor<Person>()
    .WithHeader(true)
    .WithSheetIndex(1)
    .WithManualMapping(row => new Person
    {
        // Only map name and join date
        Name = row.Cell(1).GetString(),
        JoinDate = row.Cell(4).GetDateTime()
    })
    .FromFile("employees.xlsx")
    .Extract();
```

## Performance Considerations

Based on benchmarks, manual mapping provides a small performance advantage over attribute-based mapping. Choose the approach that works best for your specific use case and code organization preferences:

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


## Why Use ExcelTransformLoad?
If you're already using [ClosedXML](https://github.com/ClosedXML/ClosedXML) or similar libraries extensively, this one might not add much extra value. But if you're looking for a simple way to read an Excel file and load it into your objects without any hassle, this library is worth checking out!

It's user-friendly and follows a fluent pattern, making it easy to define your options in a natural, intuitive way.