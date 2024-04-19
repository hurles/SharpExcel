using SharpExcel.Styling.Borders;
using SharpExcel.Styling.Colorization;
using SharpExcel.Styling.Text;

namespace SharpExcel.Styling;

public static class SharpExcelCellStyleConstants
{
    public static SharpExcelCellStyle DefaultHeaderStyle =
        new(SharpExcelColorConstants.Black, SharpExcelColorConstants.Transparent, TextStyle.Bold, borders: BorderCollection.HeaderDefault);
    
    public static SharpExcelCellStyle DefaultDataStyle =
        new(SharpExcelColorConstants.Black, SharpExcelColorConstants.Transparent, TextStyle.None, borders: BorderCollection.DataDefault);
}