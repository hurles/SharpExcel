using ExcelSharp.Styling.Borders;
using ExcelSharp.Styling.Colorization;
using ExcelSharp.Styling.Text;

namespace ExcelSharp.Styling;

public static class ExcelSharpCellStyleConstants
{
    public static ExcelSharpCellStyle DefaultHeaderStyle =
        new(ExcelSharpColorConstants.Black, ExcelSharpColorConstants.Transparent, TextStyle.Bold, borders: BorderCollection.HeaderDefault);
    
    public static ExcelSharpCellStyle DefaultDataStyle =
        new(ExcelSharpColorConstants.Black, ExcelSharpColorConstants.Transparent, TextStyle.None, borders: BorderCollection.DataDefault);
}