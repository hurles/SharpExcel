using Microsoft.Extensions.Options;
using SharpExcel.Exporters;
using SharpExcel.Models.Configuration;
using SharpExcel.TestApplication.TestData;

namespace SharpExcel.TestApplication;

public class TestExporter : BaseExcelExporter<TestExportModel>
{
    public TestExporter(IOptions<ExporterOptions<TestExportModel>> options) : base(options)
    {
    }
}