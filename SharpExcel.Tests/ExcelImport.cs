using SharpExcel.Tests.Shared;
using SharpExcel.Models;
using SharpExcel.Models.Arguments;

namespace SharpExcel.Tests;

public class Tests
{
    private TestExporter _exporter = new();


    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public async Task CreateWorkbookTest()
    {
        var workbook = await _exporter.GenerateWorkbookAsync(CreateTestData());
        Assert.IsTrue(workbook.Worksheets.FirstOrDefault(x => x.Name == "TestSheet") is not null);
    }
    
    [Test]
    public async Task ReadWorkbookTest()
    {
        //create test workbook
        var workbook = await _exporter.GenerateWorkbookAsync(CreateTestData());

        //read workbook
        var output = await _exporter.ReadWorkbookAsync("TestSheet", workbook);
        
        Assert.Multiple(() =>
        {
            Assert.That(output.Records.Count, Is.EqualTo(1));
            Assert.That(output.Records[0]?.Id, Is.EqualTo(0));
            Assert.That(output.Records[0]?.FirstName, Is.EqualTo("John"));
            Assert.That(output.Records[0]?.LastName, Is.EqualTo("Doe"));
        });
    }

    private static ExcelArguments<TestModel> CreateTestData()
    {
        return new ExcelArguments<TestModel>()
        {
            SheetName = "TestSheet",
            Data = new List<TestModel>()
            {
                new() { Id  = 0, FirstName = "John", LastName = "Doe" }
            }
        };
    }
}