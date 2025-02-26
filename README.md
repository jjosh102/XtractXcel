# ExcelTransformLoad (In-progress)

## Overview
ExcelTransformLoad is a simple .NET library for extracting data from Excel files using [ClosedXML](https://github.com/ClosedXML/ClosedXML), transforming it as needed, and loading it into your object.

## Getting Started

### Installation(Coming soon)


### Basic Usage

#### 1. Define your model with ExcelColumn attributes
```csharp
public class Person
{
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
foreach (var person in people)
{
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
public class PersonNoHeader
{
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