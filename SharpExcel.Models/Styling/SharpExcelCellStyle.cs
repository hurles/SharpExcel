using SharpExcel.Models.Styling.Borders;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Text;

namespace SharpExcel.Models.Styling;

/// <summary>
/// Struct defining the styling of a cell
/// </summary>
public struct SharpExcelCellStyle
{
    public SharpExcelCellStyle(SharpExcelColor textColor, SharpExcelColor backgroundColor, TextStyle textStyle, double? rowHeight = null, double? fontSize = null,
        BorderCollection? borders = null)
    {
        Borders = borders ?? BorderCollection.DataDefault;
        TextColor = textColor;
        BackgroundColor = backgroundColor;
        TextStyle = textStyle;
        RowHeight = rowHeight;
        FontSize = fontSize;
    }

    /// <summary>
    /// Color used for text within a cell
    /// if null, default color is used
    /// </summary>
    public SharpExcelColor? TextColor { get; set; } = SharpExcelColorConstants.Black;
    
    /// <summary>
    /// Background color used for cell
    /// if null, default color is used
    /// </summary>
    public SharpExcelColor? BackgroundColor { get; set;} = SharpExcelColorConstants.TransparentWhite;

    /// <summary>
    /// Text style (Bold, Italic, etc)
    /// Default: TextStyle.None
    /// </summary>
    public TextStyle TextStyle { get; set; } = TextStyle.None;

    /// <summary>
    /// Height of cell
    /// if null, default height is used
    /// </summary>
    public double? RowHeight { get; set; } = 20;

    /// <summary>
    /// font size
    /// if null, default size is used
    /// </summary>
    public double? FontSize { get; set; } = null;

    /// <summary>
    /// Border styling
    /// </summary>
    public BorderCollection Borders { get; set; } = BorderCollection.DataDefault;
    
    public SharpExcelCellStyle WithTextColor(SharpExcelColor color)
    {
        return new SharpExcelCellStyle(
            color,
            BackgroundColor ?? SharpExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }
    
    public SharpExcelCellStyle WithBackgroundColor(SharpExcelColor color)
    {
        return new SharpExcelCellStyle(
            TextColor ?? SharpExcelColorConstants.Black,
            color,
            TextStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }
    
    public SharpExcelCellStyle WithTextStyle(TextStyle textStyle)
    {
        return new SharpExcelCellStyle(
            TextColor ?? SharpExcelColorConstants.Black,
            BackgroundColor ?? SharpExcelColorConstants.TransparentWhite,
            textStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }

    public SharpExcelCellStyle WithRowHeight(double height)
    {
        return new SharpExcelCellStyle(
            TextColor ?? SharpExcelColorConstants.Black,
            BackgroundColor ?? SharpExcelColorConstants.TransparentWhite,
            TextStyle,
            height,
            FontSize,
            Borders
        );
    }
    
    public SharpExcelCellStyle WithFontSize(double fontSize)
    {
        return new SharpExcelCellStyle(
            TextColor ?? SharpExcelColorConstants.Black,
            BackgroundColor ?? SharpExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            fontSize,
            Borders
        );
    }
    
    public SharpExcelCellStyle WithBorders(BorderCollection borders)
    {
        return new SharpExcelCellStyle(
            TextColor ?? SharpExcelColorConstants.Black,
            BackgroundColor ?? SharpExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            FontSize,
            borders
        );
    }

}