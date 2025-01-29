namespace SharpExcel.Models.Targeting;

public class CellRangeTargetRule : SharpExcelCellTargetRule
{
    public override SharpExcelCellTargetType TargetType => SharpExcelCellTargetType.CellRange;

    private CellSelection _cellSelection;
    
    public string WorksheetName { get; set; }
    
    public CellRangeTargetRule(int startX, int startY, int endX, int endY, string worksheetName)
    {
        _cellSelection = new CellSelection(startX, startY, endX, endY);
        WorksheetName = worksheetName;
    }
    
    public CellRangeTargetRule(string addressStart, string addressEnd, string worksheetName)
    {
        _cellSelection = new CellSelection(addressStart, addressEnd);
        WorksheetName = worksheetName;
    }
    
    public CellSelection GetCellSelection()
    {
        return _cellSelection;
    }
}