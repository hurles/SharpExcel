using ExcelSharp.Models;
using ExcelSharp.TestApplication;

Console.WriteLine("Hello, World!");

var exporter = new TestExporter();

using var workbook = await exporter.BuildWorkbookAsync(CreateTestData());

workbook.SaveAs($"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx");

var importedWorkbook = await exporter.ReadWorkbookAsync("Budgets", workbook);

foreach (var dataItem in importedWorkbook)
{
    Console.WriteLine($"Id: {dataItem.Id} -- First name: {dataItem.FirstName} -- Last name: {dataItem.LastName} -- Email: {dataItem.Email} -- Budget: {dataItem.Budget}");
}


ExcelArguments<TestExportModel> CreateTestData()
{
    return new ExcelArguments<TestExportModel>()
    {
        SheetName = "Budgets",
        Data = new List<TestExportModel>()
        {
            new() { Id  = 0, FirstName = "John", LastName = "Doe", Budget = 2400.34m, Email = "john.doe@example.com" },
            new() { Id  = 1, FirstName = "Jane", LastName = "Doe", Budget = 200.42m, Email = "jane.doe@example.com" },
            new() { Id  = 2, FirstName = "Jimmy", LastName = "Neutron", Budget = 0.0m, Email = null },
            new() { Id  = 3, FirstName = "Ash", LastName = "Ketchum", Budget = 69m, Email = "ash@example.com" },
            new() { Id  = 4, FirstName = "Inspector", LastName = "Gadget", Budget = 1337m, Email = "gogogadget@example.com" },
            new() { Id  = 5, FirstName = "Mickey", LastName = "Mouse", Budget = 2400.34m, Email = "mmouse@example.com" },
        }
    };
}