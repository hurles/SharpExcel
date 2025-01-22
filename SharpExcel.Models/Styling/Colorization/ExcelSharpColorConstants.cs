namespace SharpExcel.Models.Styling.Colorization;

/// <summary>
/// Constant values for a set of colors
/// </summary>
public static class ExcelColorConstants
{
    public static ExcelColor White = new(255, 255, 255);
    public static ExcelColor Black = new(0, 0, 0);
    public static ExcelColor Red = new(255, 0, 0);
    public static ExcelColor Lime = new(0, 255, 0);
    public static ExcelColor Blue = new(0, 0, 255);
    public static ExcelColor Yellow = new(255, 255, 0);
    public static ExcelColor Cyan = new(0, 255, 255);
    public static ExcelColor Magenta = new(255, 0, 255);
    public static ExcelColor Silver = new(192, 192, 192);
    public static ExcelColor Gray = new(128, 128, 128);
    public static ExcelColor Maroon = new(128, 0, 0);
    public static ExcelColor Olive = new(128, 128, 0);
    public static ExcelColor Green = new(0, 128, 0);
    public static ExcelColor Purple = new(128, 0, 128);
    public static ExcelColor Teal = new(0, 128, 128);
    public static ExcelColor Navy = new(0, 0, 128);
    
    //transparent
    public static ExcelColor Transparent = new(0, 0, 0, 0);
    public static ExcelColor TransparentWhite = new(255, 255, 255, 0);

}

