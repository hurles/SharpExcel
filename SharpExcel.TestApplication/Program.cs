using SharpExcel.Models;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Arguments.Extensions;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Extensions;
using SharpExcel.TestApplication;
using SharpExcel.TestApplication.TestData;

var excelArguments = new ExcelArguments<TestExportModel>()
{
    SheetName = "Budgets",
    Data = TestDataProvider.GetTestData()
};

excelArguments.AddStylingRule()
    .ForProperty(nameof(TestExportModel.Budget))
    .WithCondition(x => x.Budget < 0.0m)
    .WhenTrue(new SharpExcelCellStyle()
    {
        TextColor = SharpExcelColorConstants.Red
    })
    .WhenFalse(new SharpExcelCellStyle()
    {
        TextColor = SharpExcelColorConstants.Green
    });

excelArguments.AddStylingRule()
    .ForProperty(nameof(TestExportModel.Status))
    .WithCondition(x => x.Status == TestStatus.Employed)
    .WhenTrue(new SharpExcelCellStyle()
    {
        TextColor = SharpExcelColorConstants.Green
    });

//create exporter. This is of type BaseExcelExporter<TestExportModel>
var exporter = new TestExporter();

//Step 1 -- Writing:
//Create and save a new workbook based on the test data outlined above
var exportPath = $"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx";
Console.WriteLine("-- Writing test data to workbook.. --");
using var workbook = await exporter.GenerateWorkbookAsync(excelArguments);
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
    Console.WriteLine($"{testExportModel?.Id} | {testExportModel?.FirstName} | {testExportModel?.LastName} | {testExportModel?.Email} | {testExportModel?.Budget} | {testExportModel?.TestDepartment}");
    
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