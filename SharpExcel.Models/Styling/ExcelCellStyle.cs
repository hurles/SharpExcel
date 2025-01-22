using SharpExcel.Models.Styling.Borders;
using SharpExcel.Models.Styling.Colorization;
using SharpExcel.Models.Styling.Text;

namespace SharpExcel.Models.Styling;

/// <summary>
/// Struct defining the styling of a cell
/// </summary>
public record struct ExcelCellStyle
{
    public ExcelCellStyle() { }

    public ExcelCellStyle(ExcelColor textColor, ExcelColor backgroundColor, TextStyle textStyle, double? rowHeight = null, double? fontSize = null,
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
    public ExcelColor? TextColor { get; set; } = ExcelColorConstants.Black;
    
    /// <summary>
    /// Background color used for cell
    /// if null, default color is used
    /// </summary>
    public ExcelColor? BackgroundColor { get; set;} = ExcelColorConstants.TransparentWhite;

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
    public BorderCollection? Borders { get; set; } = BorderCollection.DataDefault;
    
    public ExcelCellStyle WithTextColor(ExcelColor color)
    {
        return new ExcelCellStyle(
            color,
            BackgroundColor ?? ExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }
    
    public ExcelCellStyle WithBackgroundColor(ExcelColor color)
    {
        return new ExcelCellStyle(
            TextColor ?? ExcelColorConstants.Black,
            color,
            TextStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }
    
    public ExcelCellStyle WithTextStyle(TextStyle textStyle)
    {
        return new ExcelCellStyle(
            TextColor ?? ExcelColorConstants.Black,
            BackgroundColor ?? ExcelColorConstants.TransparentWhite,
            textStyle,
            RowHeight,
            FontSize,
            Borders
        );
    }

    public ExcelCellStyle WithRowHeight(double height)
    {
        return new ExcelCellStyle(
            TextColor ?? ExcelColorConstants.Black,
            BackgroundColor ?? ExcelColorConstants.TransparentWhite,
            TextStyle,
            height,
            FontSize,
            Borders
        );
    }
    
    public ExcelCellStyle WithFontSize(double fontSize)
    {
        return new ExcelCellStyle(
            TextColor ?? ExcelColorConstants.Black,
            BackgroundColor ?? ExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            fontSize,
            Borders
        );
    }
    
    public ExcelCellStyle WithBorders(BorderCollection borders)
    {
        return new ExcelCellStyle(
            TextColor ?? ExcelColorConstants.Black,
            BackgroundColor ?? ExcelColorConstants.TransparentWhite,
            TextStyle,
            RowHeight,
            FontSize,
            borders
        );
    }

}