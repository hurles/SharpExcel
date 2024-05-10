using Microsoft.Extensions.Options;
using SharpExcel;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;

namespace SharpExcel.Tests.Shared;

public class TestExporter : BaseExcelExporter<TestModel>
{
    public TestExporter(IOptions<ExporterOptions<TestModel>> options) : base(options)
    {
    }
}