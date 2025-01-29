namespace SharpExcel.Models.Targeting;

public class WorkSheetTargetRule : SharpExcelCellTargetRule
{
    public override SharpExcelCellTargetType TargetType  => SharpExcelCellTargetType.Worksheet;

    public string WorksheetName { get; set; } = null!;
}

public class WorkbookTargetRule : SharpExcelCellTargetRule
{
    public override SharpExcelCellTargetType TargetType => SharpExcelCellTargetType.Workbook;
}

// Defines a Rectangular selection of cells