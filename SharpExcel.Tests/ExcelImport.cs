using Microsoft.Extensions.Options;
using SharpExcel.Tests.Shared;
using SharpExcel.Models;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Configuration.Constants;

namespace SharpExcel.Tests;

public class Tests
{
    private TestExporter _exporter = null!;


    [SetUp]
    public void Setup()
    {
        var options = Options.Create(ExporterOptionsConstants.GetDefaultOptions<TestModel>());

        _exporter = new TestExporter(options);

    }

    [Test]
    public async Task CreateWorkbookTest()
    {
        var workbook = await _exporter.GenerateWorkbookAsync(new SharpExcelArguments(){ SheetName = "TestSheet"}, CreateTestData());
        Assert.IsTrue(workbook.Worksheets.FirstOrDefault(x => x.Name == "TestSheet") is not null);
    }
    
    [Test]
    public async Task ReadWorkbookTest()
    {
        //create test workbook
        var workbook = await _exporter.GenerateWorkbookAsync( new SharpExcelArguments(){ SheetName = "TestSheet"}, CreateTestData());

        //read workbook
        var output = await _exporter.ReadWorkbookAsync("TestSheet", workbook);
        
        Assert.Multiple(() =>
        {
            Assert.That(output.Records.Count, Is.EqualTo(2));
            Assert.That(output.Records[0]?.Id, Is.EqualTo(1));
            Assert.That(output.Records[0]?.FirstName, Is.EqualTo("John"));
            Assert.That(output.Records[0]?.LastName, Is.EqualTo("Doe"));
        });
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