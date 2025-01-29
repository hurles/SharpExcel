namespace SharpExcel.Models.Targeting;

public enum SharpExcelCellTargetType
{
    //Processes range of cells
    CellRange,
    //Processes a given worksheet
    Worksheet,
    //Processes all unhidden sheets in the workbook
    Workbook
}