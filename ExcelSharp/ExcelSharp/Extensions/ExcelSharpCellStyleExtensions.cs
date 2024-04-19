using ClosedXML.Excel;
using ExcelSharp.Styling;
using ExcelSharp.Styling.Borders;
using ExcelSharp.Styling.Text;

namespace ExcelSharp.Extensions;

internal static class ExcelSharpCellStyleExtensions
{
    internal static void ApplyStyle(this IXLStyle excelStyle, ExcelSharpCellStyle cellStyle)
    {
        if (cellStyle.BackgroundColor.HasValue)
        {
            excelStyle.Fill.BackgroundColor = XLColor.FromArgb(
                cellStyle.BackgroundColor.Value.A,
                cellStyle.BackgroundColor.Value.R,
                cellStyle.BackgroundColor.Value.G,
                cellStyle.BackgroundColor.Value.B);
        }
        
        if (cellStyle.TextColor.HasValue)
        {
            excelStyle.Font.FontColor = XLColor.FromArgb(
                cellStyle.TextColor.Value.A,
                cellStyle.TextColor.Value.R,
                cellStyle.TextColor.Value.G,
                cellStyle.TextColor.Value.B);
        }

        if (cellStyle.FontSize.HasValue)
        {
            excelStyle.Font.FontSize = (double)cellStyle.FontSize;
        }
        
        if (cellStyle.TextStyle != TextStyle.None)
        {
            switch (cellStyle.TextStyle)
            {
                case TextStyle.Bold:
                    excelStyle.Font.Bold = true;
                    break;
                case TextStyle.Italic:
                    excelStyle.Font.Italic = true;
                    break;
                case TextStyle.Underlined:
                    excelStyle.Font.SetUnderline(XLFontUnderlineValues.Single);
                    break;
            }
        }

        excelStyle.Border.SetBottomBorder(
            GetBorderStyleValue(cellStyle.Borders.GetBorderStyle(BorderDirection.Bottom)));
        excelStyle.Border.SetLeftBorder(
            GetBorderStyleValue(cellStyle.Borders.GetBorderStyle(BorderDirection.Left)));
        excelStyle.Border.SetRightBorder(
            GetBorderStyleValue(cellStyle.Borders.GetBorderStyle(BorderDirection.Right)));
        excelStyle.Border.SetTopBorder(
            GetBorderStyleValue(cellStyle.Borders.GetBorderStyle(BorderDirection.Top)));
    }

    private static XLBorderStyleValues GetBorderStyleValue(BorderStyle borderStyle)
    {
        switch (borderStyle)
        {
            case BorderStyle.None:
                return XLBorderStyleValues.None;
            case BorderStyle.DashDot:
                return XLBorderStyleValues.DashDot;
            case BorderStyle.DashDotDot:
                return XLBorderStyleValues.DashDotDot;
            case BorderStyle.Dashed:
                return XLBorderStyleValues.Dashed;
            case BorderStyle.Dotted:
                return XLBorderStyleValues.Dotted;
            case BorderStyle.Double:
                return XLBorderStyleValues.Double;
            case BorderStyle.Hair:
                return XLBorderStyleValues.Hair;
            case BorderStyle.Medium:
                return XLBorderStyleValues.Medium;
            case BorderStyle.MediumDashDot:
                return XLBorderStyleValues.MediumDashDot;
            case BorderStyle.MediumDashDotDot:
                return XLBorderStyleValues.MediumDashDotDot;
            case BorderStyle.MediumDashed:
                return XLBorderStyleValues.MediumDashed;
            case BorderStyle.SlantDashDot:
                return XLBorderStyleValues.SlantDashDot;
            case BorderStyle.Thick:
                return XLBorderStyleValues.Thick;
            default:
                return XLBorderStyleValues.Thin;

        }
    }
}

