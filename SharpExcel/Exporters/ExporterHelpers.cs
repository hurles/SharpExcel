using System.Text;
using ClosedXML.Excel;
using SharpExcel.Extensions;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling.Constants;

namespace SharpExcel.Exporters;

internal static class ExporterHelpers
{
    public static void ApplyCellValidation<TModel>(string sheetName, XLWorkbook workbook, ExcelReadResult<TModel> parsedWorkbook)
        where TModel : class, new()
    {
        foreach (var result in parsedWorkbook.ValidationResults)
        {
            var cell = workbook.Worksheet(sheetName).Cell(result.Value.Address.RowNumber, result.Value.Address.ColumnId);
            var stringBuilder = new StringBuilder();
            foreach (var item in result.Value.ValidationResults)
            {
                stringBuilder.AppendLine(item.ErrorMessage);
            }
            cell.Style.ApplyStyle(SharpExcelCellStyleConstants.DefaultErrorStyle);
            cell.CreateComment().AddText(stringBuilder.ToString());
        }
    }
}