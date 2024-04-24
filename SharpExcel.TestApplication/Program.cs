using SharpExcel.Models;
using SharpExcel.TestApplication;

Console.WriteLine("Writing .xslsx file");

var exporter = new TestExporter();

using var workbook = await exporter.BuildWorkbookAsync(CreateTestData());

workbook.SaveAs($"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx");

var importedWorkbook = await exporter.ReadWorkbookAsync("Budgets", workbook);

foreach (var dataItem in importedWorkbook.Records)
{
    Console.WriteLine($"Id: {dataItem?.Id} -- First name: {dataItem?.FirstName} -- Last name: {dataItem?.LastName} -- Email: {dataItem?.Email} -- Budget: {dataItem?.Budget}");
    if (importedWorkbook.ValidationResults.TryGetValue(dataItem, out var validationResults))
    {
        foreach (var validationResult in validationResults.ValidationResults)
        {
            Console.WriteLine($"\tValidation Error in on row {validationResults.Address.RowNumber} in column {validationResults.Address.ColumnName} ({validationResults.Address.HeaderName}) {validationResult.ErrorMessage}");
        }
    }
}

ExcelArguments<TestExportModel> CreateTestData()
{
    return new ExcelArguments<TestExportModel>()
    {
        SheetName = "Budgets",
        Data = new List<TestExportModel>()
        {
            new() { Id  = 0, FirstName = "John", LastName = "Doe", Budget = 2400.34m, Email = "john.doe@example.com" },
            new() { Id  = 1, FirstName = "Jane", LastName = "Doe", Budget = -200.42m, Email = "jane.doe@example.com" },
            new() { Id  = 2, FirstName = "Jimmy", LastName = "Neutron", Budget = 0.0m, Email = null },
            new() { Id  = 3, FirstName = "Ash", LastName = "Ketchum", Budget = 69m, Email = "ash@example.com" },
            new() { Id  = 4, FirstName = "Inspector", LastName = "Gadget", Budget = 1337m, Email = "gogogadget@example.com" },
            new() { Id  = 5, FirstName = "Mickey", LastName = "Mouse", Budget = 2400.34m, Email = "mmouse@example.com" },
            new() { Id  = 6, FirstName = "ThisIsLongerThan10", LastName = "Mouse", Budget = 2400.34m, Email = "mmouse@example.com" },
        }
    };
}