using SharpExcel.Models;
using SharpExcel.Models.Results;
using SharpExcel.TestApplication;

var testData = new ExcelArguments<TestExportModel>()
{
    SheetName = "Budgets",
    Data = new List<TestExportModel>()
    {
        new() { Id  = 0, FirstName = "John", LastName = "Doe", Budget = 2400.34m, Email = "john.doe@example.com", Department = Department.Unknown },
        new() { Id  = 1, FirstName = "Jane", LastName = "Doe", Budget = -200.42m, Email = "jane.doe@example.com", Department = Department.ValueB },
        new() { Id  = 2, FirstName = "John", LastName = "Neutron", Budget = 0.0m, Email = null, Department = Department.ValueB },
        new() { Id  = 3, FirstName = "Ash", LastName = "Ketchum", Budget = 69m, Email = "ash@example.com", Department = Department.ValueC },
        new() { Id  = 4, FirstName = "Inspector", LastName = "Gadget", Budget = 1337m, Email = "gogogadget@example.com", Department = Department.ValueC },
        new() { Id  = 5, FirstName = "Mindy", LastName = "", Budget = 2400.34m, Email = "mmouse@example.com", Department = Department.ValueA },
        new() { Id  = 6, FirstName = "ThisIsLongerThan10", LastName = "Mouse", Budget = 2400.34m, Email = "mmouse@example.com", Department = Department.ValueA },
        new() { Id  = 7, FirstName = "Name", LastName = "LasName", Budget = 2400.34m, Email = null, Department = Department.ValueB },
    }
};

//create exporter. This is of type BaseExcelExporter<TestExportModel>
var exporter = new TestExporter();

//Step 1 -- Writing:
//Create and save a new workbook based on the test data outlined above
var exportPath = $"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx";
Console.WriteLine("-- Writing test data to workbook.. --");
using var workbook = await exporter.GenerateWorkbookAsync(testData);
workbook.SaveAs(exportPath);
Console.WriteLine($"-- Saved successfully: {exportPath} --");

//Step 2 -- Validation:
//Modify the original workbook so it highlights and annotates invalid cells.
//The conditions for field validation are determined by data annotations on TestExportModel.
Console.WriteLine($"-- Validating workbook --");
var validationExportPath = $"./OutputFolder/ErrorChecked-{Guid.NewGuid()}.xlsx";
var errorCheckedWorkbook = await exporter.ValidateAndAnnotateWorkbookAsync("Budgets", workbook);
errorCheckedWorkbook.SaveAs(validationExportPath);
Console.WriteLine($"-- Saved successfully: {validationExportPath} --");

//Step 3 -- Reading
//Read and import the workbook, and turn it back into a collection of TestExportModel
Console.WriteLine($"-- Reading workbook --");
var importedWorkbook = await exporter.ReadWorkbookAsync("Budgets", workbook);

Console.WriteLine($"-------------------------------------");
Console.WriteLine($"Results:\n");
//write headers
Console.WriteLine($"Id | First name | Last name | Email | Budget | Department");

//write read rows
foreach (var dataItem in importedWorkbook.Records)
{
    WriteOutputRow(dataItem, importedWorkbook);
}


//This method is just here to write the results of the read operation.
void WriteOutputRow(TestExportModel testExportModel, ExcelReadResult<TestExportModel> excelReadResult)
{
    Console.WriteLine($"{testExportModel?.Id} | {testExportModel?.FirstName} | {testExportModel?.LastName} | {testExportModel?.Email} | {testExportModel?.Budget} | {testExportModel?.Department}");
    
    //print validation errors if needed
    if (testExportModel != null && excelReadResult.ValidationResults.TryGetValue(testExportModel, out var validationResults))
    {
        foreach (var validationResult in validationResults.ValidationResults)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"\tValidation Error on row {validationResults.Address.RowNumber} in column {validationResults.Address.ColumnName} ({validationResults.Address.HeaderName}): {validationResult.ErrorMessage}");
            Console.ResetColor();
        }
    }
}