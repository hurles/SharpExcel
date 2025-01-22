using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using SharpExcel.Tests.Shared;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Configuration.Constants;
using Shouldly;
using Xunit;

namespace SharpExcel.Tests;

public class ExcelImportTests
{
    private TestSynchronizer _synchronizer = null!;

    public ExcelImportTests()
    {
        var options = Options.Create(ExporterOptionsConstants.GetDefaultOptions<TestModel>());

        _synchronizer = new TestSynchronizer(options);

    }

    [Fact]
    public async Task CreateWorkbookTest()
    {
        var workbook = await _synchronizer.GenerateWorkbookAsync(new ExcelArguments(){ SheetName = "TestSheet"}, CreateTestData());
        workbook.Worksheets.FirstOrDefault(x => x.Name == "TestSheet").ShouldNotBeNull();
        
        workbook.ShouldNotBeNull();
        //there should be 2 worksheets, a visible one for the data, and a hidden one to pull data from for the enum dropdowns
        workbook.Worksheets.Count.ShouldBe(2);
        
        //main data worksheet
        workbook.Worksheet(1).Name.ShouldBe("TestSheet");
        workbook.Worksheet(1).Visibility.ShouldBe(XLWorksheetVisibility.Visible);
        
        //hidden worksheet for enum dropdowns
        workbook.Worksheet(2).Visibility.ShouldBe(XLWorksheetVisibility.Hidden);
    }
    
    [Fact]
    public async Task ReadWorkbookTest()
    {
        var args = new ExcelArguments() { SheetName = "TestSheet" };
        //create test workbook
        var workbook = await _synchronizer.GenerateWorkbookAsync( args, CreateTestData());

        //read workbook
        var output = await _synchronizer.ReadWorkbookAsync(args, workbook);
        
        
        output.Records.Count.ShouldBe(2);
        output.Records[0].Id.ShouldBe(1);
        output.Records[0].FirstName.ShouldBe("John");
        output.Records[0].LastName.ShouldBe("Doe");
    }

    private static List<TestModel> CreateTestData()
    {
        return new List<TestModel>()
        {
            new () { Id = 1, FirstName = "John", LastName = "Doe", TestValue = TestEnum.ValueA },
            new () { Id = 2, FirstName = "Jane", LastName = "Doe", TestValue = TestEnum.ValueB },
        };
    }
}