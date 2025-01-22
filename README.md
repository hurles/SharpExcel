# SharpExcel

SharpExcel is a powerful, easy-to-use .NET Standard 2.0 library designed to simplify the process of importing, exporting, styling, and validating Excel files. SharpExcel uses ClosedXml to handle reading and writing Excel files. 

### Main focus
The library is focused on mapping a collection of C# models to a corresponding Excel file. 

**ExcelSharp makes sure that every Excel file you export, can also be re-imported and converted to the same data as was used to export it. This is useful for providing a template for a user or client to provide data to load into a program.**


### Validation
The library uses FluentValidation to validate imported data. This will generate a list of exactly which cells are invalid, and why.
We can even output a new Excel file, where all invalid cells have a red color, or any other defined style.

### Styling
SharpExcel also provides a fluent API to define styles. We can set default data and header styles and even override styles based on specific rules (for example: make a cell red when a number is below zero).

### Auto Dropdowns
Enum properties in your model will be automatically mapped into dropdown lists for a user to select.

---
## Install SharpExcel

If you want to include SharpExcel in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/SharpExcel)

To install SharpExcel, run the following command in the Package Manager Console
```
PM> Install-Package SharpExcel
```
---
## Usage

There are a couple of simple steps to start using SharpExcel:

### Step 1: Define a data model

When defining a data model, we can use the ``[ExcelColumnDefinition]`` attribute to map Excel columns to model properties.
We can also use Data annotation attributes to generate validation errors when reading Excel files.

*In this example model we create a model for an employee:*

```csharp
public class EmployeeModel
{
    [ExcelColumnDefinition(columnName: "ID", width: 45)]
    public int Id { get; set; }

    [ExcelColumnDefinition(columnName: "First Name", width: 30)]
    public string FirstName { get; set; } = null!;

    [StringLength(12)]
    [ExcelColumnDefinition(columnName: "Last Name", width: 50)]
    public string LastName { get; set; } = null!;
    
    [ExcelColumnDefinition(columnName: "Budget", width: 15)]
    public decimal Budget { get; set; }
    
    //SharpExcel also supports enum values (these will be displayed as dropdowns in Excel)
    [ExcelColumnDefinition(columnName: "Employment status", width: 50)]
    public EmploymentStatus Status { get; set; } = EmploymentStatus.Employed;
}
```

*important note: In the current implementation, only classes are supported as data models. structs are currently not supported.*

---
### Step 2: Register a SharpExcel synchronizer
In the simplest case we can register a synchronizer for the given model to the service collection.
This is a default implementation and can be used for simple imports/exports.
```csharp
builder.Services.AddSynchronizer<EmployeeModel>()
```
Optionally, we can configure the synchronizer further:
```csharp
builder.Services.AddSynchronizer<TestExportModel>(options =>
{
    //apply default styling
    options.WithDataStyle(SharpExcelCellStyleConstants.DefaultDataStyle);
    
    //in this case we customize the styling for the header
    options.WithHeaderStyle(new SharpExcelCellStyle()
        .WithTextStyle(TextStyle.Bold)
        .WithFontSize(18.0));
});
```
If we want to switch styling conditionally, styling rules can be added in the following way:

*In this example, we want the text in the cell to be red when the budget is < 0*
```csharp
builder.Services.AddSynchronizer<EmployeeModel>(options =>
{
    options.WithStylingRule(rule =>
        {
            //select property of model by name
            rule.ForProperty(nameof(TestExportModel.Budget));
            //provide a condition
            rule.WithCondition(x => x.Budget < 0);
            //color text red when condition is true
            rule.WhenTrue(SharpExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(255, 100, 100)));
            //color text green when condition is false
            rule.WhenFalse(SharpExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(80, 160, 80)));
        });
});
```

---
### Step 3: Read / Write Excel files

To use the synchronizer we must first inject a ``ISharpExcelSynchronizer<TModel>`` service, where ``TModel`` is the provided data model decorated with the SharpExcel attributes.

*For the service we registered in step 2 we can use ``ISharpExcelSynchronizer<EmployeeModel>``*
```csharp
public class ApplicationService
{
    private readonly ISharpExcelSynchronizer<EmployeeModel> _synchronizer;
    
    public ApplicationService(ISharpExcelSynchronizer<EmployeeModel> synchronizer)
    {
        _synchronizer = synchronizer;
    }
}
```
#### Writing Excel files

*The following example shows how to write a collection of ``EmployeeModel`` to an excel file:*

```csharp
   var arguments = new SharpExcelArguments()
   {
       //sheet to read from
       SheetName = "Employees",
       //optional culture to use when reading
       CultureInfo = CultureInfo.CurrentCulture
   };

   //this doesn't have to be a list, an IEnumerable will also do
    var data = new List<EmployeeModel>()
    {
        new () { Id = 1, FirstName = "John", LastName = "Doe", Budget = 12.0m }
    };

    //write the collection to an XLWorkbook
    using var workbook = await _synchronizer.GenerateWorkbookAsync(arguments, data);
   
    //in this case we save to a file, but this can also be a stream
    workbook.SaveAs("C:/Documents/filename.xslx");
```
The saving of the excel file (and the XLWorkbook type) are provided by ClosedXml
For more information and documentation on these types visit [ClosedXml](https://github.com/ClosedXML/ClosedXML)

#### Reading Excel files

To read an Excel file we must provide an Excel workbook, 
then use the previously injected service to start parsing it. 
The excel file must have a header row with the column names defined in the model, in any order. 

```csharp
    // in this case we load from a file, but this can also be a stream
    using var workbook = new XLWorkbook("C:/Documents/filename.xslx");
    
    var arguments = new SharpExcelArguments()
    {
        //which sheet to read data from
        SheetName = "Employees",
        //optional
        CultureInfo = CultureInfo.CurrentCulture
    };
    
    await _synchronizer.ReadWorkbookAsync(arguments, workbook);
```
The loading of the excel file (and the XLWorkbook type) are provided by ClosedXml 
For more information and documentation on these types visit [ClosedXml](https://github.com/ClosedXML/ClosedXML)

#### Validating Excel files

```csharp

```
