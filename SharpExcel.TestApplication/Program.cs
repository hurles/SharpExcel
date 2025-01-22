using System.Globalization;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.TestApplication.TestData;

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharpExcel.Abstraction;
using SharpExcel.DependencyInjection;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Constants;
using SharpExcel.Models.Styling.Text;

HostApplicationBuilder builder = Host.CreateEmptyApplicationBuilder(null);

//Here we add a synchronizer to the service collection
builder.Services.AddSharpExcelSynchronizer<TestExportModel>(options =>
{
    //set default style of
    options.WithDataStyle(ExcelCellStyleConstants.DefaultDataStyle);
    options.WithHeaderStyle(new ExcelCellStyle()
        .WithTextStyle(TextStyle.Bold | TextStyle.Underlined)
        .WithFontSize(18.0));
    
    //here we define the style of an errored cell.
    //This is only applicable when we want to return a validated excel file.
    //Any cells that have validation errors will have this style
    options.WithErrorStyle(
            ExcelCellStyleConstants.DefaultDataStyle
                .WithTextColor(new ExcelColor(255, 100, 100))
                .WithBackgroundColor(new ExcelColor(255, 100, 100, 70))
        );
    
    //We can also define rules for styling
    options.WithStylingRule(rule =>
        {
            //define which property we want to check
            rule.ForProperty(nameof(TestExportModel.Budget));
            
            //define the condition for this rule
            rule.WithCondition(x => x.Budget < 0);
            
            //define style to do when the rule is true
            rule.WhenTrue(ExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(255, 100, 100)));
            
            //define style for when the rule is false
            //can be omitted to use default style
            rule.WhenFalse(ExcelCellStyleConstants.DefaultDataStyle.WithTextColor(new(80, 160, 80)));
        });
});

using IHost host = builder.Build();
await RunApp(host.Services);
await host.RunAsync();

//this is our test application
async Task RunApp(IServiceProvider services)
{
    var exportPath = $"./OutputFolder/TestExport-{Guid.NewGuid()}.xlsx";
    var validationExportPath = $"./OutputFolder/ErrorChecked-{Guid.NewGuid()}.xlsx";
    var exportService = services.GetRequiredService<ISharpExcelSynchronizer<TestExportModel>>();

    var excelArguments = new ExcelArguments()
    {
        SheetName = "Budgets",
        CultureInfo = CultureInfo.CurrentCulture
    };
    
    using var workbook = await exportService.GenerateWorkbookAsync(excelArguments, TestDataProvider.GetTestData());
    workbook.SaveAs(exportPath);
    
    using var errorCheckedWorkbook = await exportService.ValidateAndAnnotateWorkbookAsync(excelArguments, workbook);
    errorCheckedWorkbook.SaveAs(validationExportPath);

    var importedWorkbook = await exportService.ReadWorkbookAsync(excelArguments, workbook);

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