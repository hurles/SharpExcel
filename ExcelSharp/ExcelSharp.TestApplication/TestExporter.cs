using ExcelSharp.Styling;
using ExcelSharp.Styling.Colorization;
using ExcelSharp.Styling.Text;

namespace ExcelSharp.TestApplication;

public class TestExporter : BaseExcelExporter<TestExportModel>
{
    public override ExcelSharpCellStyle GetHeaderStyle()
    {
        var headerStyle = ExcelSharpCellStyleConstants.DefaultHeaderStyle;
        headerStyle.FontSize = 13.0f;
        headerStyle.TextStyle = TextStyle.Bold;
        headerStyle.BackgroundColor = new ExcelSharpColor(200, 200, 200, 255);

        return headerStyle;
    }

    public override ExcelSharpCellStyle GetDataStyle(string propertyName, TestExportModel record)
    {
        var dataStyle = ExcelSharpCellStyleConstants.DefaultDataStyle;

        if (propertyName == nameof(TestExportModel.Budget))
        {
            switch (record.Budget)
            {
                case < 0:
                    dataStyle.TextColor = ExcelSharpColorConstants.Red;
                    break;
                case > 0:
                    dataStyle.TextColor = ExcelSharpColorConstants.Green;
                    break;
                default:
                    dataStyle.TextColor = ExcelSharpColorConstants.Black;
                    break;
            }
        }

        return dataStyle;
    }
}