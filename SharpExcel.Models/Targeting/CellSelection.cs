using SharpExcel.Models.Helpers;

namespace SharpExcel.Models.Targeting;

public struct CellSelection
{
    public CellSelection() { }

    public CellSelection(string addressStart, string addressEnd)
    {
        var startAddress = ColumnHelper.ConvertFromAddress(addressStart);
        var endAddress = ColumnHelper.ConvertFromAddress(addressEnd);
        StartCellX = startAddress.column;
        StartCellY = startAddress.row;
        EndCellX = endAddress.column;
        EndCellY = endAddress.row;
    }
    
    public CellSelection(int startRow, int startColumn, int endRow, int endColumn)
    {
        StartCellX = startColumn;
        StartCellY = startRow;
        EndCellX = endColumn;
        EndCellY = endRow;
    }

    public int StartCellX { get; set; } = 1;
    public int StartCellY { get; set; } = 1;
    public int EndCellX { get; set; } = 1;
    public int EndCellY { get; set; } = 1;
}