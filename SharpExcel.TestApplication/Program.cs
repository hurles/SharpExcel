using System.Globalization;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.TestApplication.TestData;

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharpExcel.Abstraction;
using SharpExcel.DependencyInjection;
using SharpExcel.Models.Styling.Constants;

HostApplicationBuilder builder = Host.CreateEmptyApplicationBuilder(null);

builder.Services.AddExporter<TestExportModel>(options =>
{
    options.WithDataStyle(SharpExcelCellStyleConstants.DefaultDataStyle)
        .WithHeaderStyle(SharpExcelCellStyleConstants.DefaultHeaderStyle)
        .WithErrorStyle(
            SharpExcelCellStyleConstants.DefaultDataStyle
                .WithTextColor(new SharpExcelColor(255, 100, 100))
                .WithBackgroundColor(new SharpExcelColor(255, 100, 100, 70))
        )
        .WithStylingRule(rule =>
        {
            rule.ForProperty(nameof(TestExportModel.Budget));
            rule.WithCondition(x => x.Budget < 0);
            rule.WhenTrue(SharpExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(255, 100, 100)));
            rule.WhenFalse(SharpExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(80, 160, 80)));
        });
});

using IHost host = builder.Build();
await RunApp(host.Services);
await host.RunAsync();

async Task RunApp(IServiceProvider services)
{
    var exportPath = $"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx";
    var validationExportPath = $"./OutputFolder/ErrorChecked-{Guid.NewGuid()}.xlsx";
    var exportService = services.GetRequiredService<IExcelExporter<TestExportModel>>();

    #region write-workbook
    var excelArguments = new SharpExcelArguments()
    {
        SheetName = "Budgets",
        CultureInfo = CultureInfo.CurrentCulture
    };
    
    using var workbook = await exportService.GenerateWorkbookAsync(excelArguments, TestDataProvider.GetTestData());
    workbook.SaveAs(exportPath);
    #endregion
    
    #region validate-workbook
    using var errorCheckedWorkbook = await exportService.ValidateAndAnnotateWorkbookAsync(excelArguments, workbook);
    errorCheckedWorkbook.SaveAs(validationExportPath);
    #endregion

    #region read-workbook
    var importedWorkbook = await exportService.ReadWorkbookAsync(excelArguments, workbook);
    #endregion

    #region write_output
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
    #endregion
}