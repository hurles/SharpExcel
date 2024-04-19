using System.Globalization;
using System.Reflection;
using ClosedXML.Excel;
using SharpExcel.Extensions;
using SharpExcel.Abstraction;
using SharpExcel.Attributes;
using SharpExcel.Models;
using SharpExcel.Styling;

namespace SharpExcel;

/// <summary>
/// Base class for creating excel workbooks
/// </summary>
public abstract class BaseExcelExporter<TModel> : IExcelExporter<TModel>
    where TModel : class, new()
{
    /// <summary>
    /// method to actually build workbook
    /// </summary>
    /// <param name="arguments"></param>
    /// <param name="optionalColumnFunc">Functions that returns a boolean based on Property name of TModel (not column name), indicating whether or not to write the specified property based on a condition</param>
    /// <returns></returns>
    public async Task<XLWorkbook> BuildWorkbookAsync(ExcelArguments<TModel> arguments,
        Func<string, Task<bool>>? optionalColumnFunc = null)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet(arguments.SheetName);

        var headerStyle = GetHeaderStyle();

        if (headerStyle.RowHeight.HasValue)
        {
            worksheet.Rows().Height = headerStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var propertyMappings = GetModelMetaData();
        var optionalColumns = await GetOptionalPropertiesToExport(optionalColumnFunc);

        int offsetColumns = 0;
        for (var columnIndex = 0; columnIndex < propertyMappings.Count; columnIndex++)
        {

            var mapping = propertyMappings[columnIndex];

            if (mapping.Optional && !optionalColumns.Contains(mapping.PropertyInfo.Name))
            {
                offsetColumns++;
                continue;
            }

            var cell = worksheet.Cell(rowIndex, columnIndex + 1 - offsetColumns /* use +1 because Excel starts at 1 */);
            
            
            cell.Style.ApplyStyle(headerStyle);

            if (mapping.ColumnWidth > 0)
            {
                worksheet.Column(columnIndex + 1).Width = mapping.ColumnWidth;
            }

            cell.SetValue(mapping.Name);
        }

        //go to next row to start inserting data
        rowIndex++;

        if (arguments?.Data is null)
            return workbook;

        foreach (var dataItem in arguments?.Data!)
        {
            var dataOffset = 0;
            for (var i = 0; i < propertyMappings.Count; i++)
            {
                var mapping = propertyMappings[i];
                var row = worksheet.Row(rowIndex);
                
                if (mapping.Optional && !optionalColumns.Contains(mapping.PropertyInfo.Name))
                {
                    dataOffset++;
                    continue;
                }

                var cell = worksheet.Cell(rowIndex, i + 1 - dataOffset /* use +1 because Excel starts at 1 */);

                var dataStyle = GetDataStyle(mapping.PropertyInfo.Name, dataItem);

                if (dataStyle.RowHeight.HasValue &&  row.Height < dataStyle.RowHeight)
                {
                    row.Height = dataStyle.RowHeight.Value;
                }

                var dataValue = mapping.PropertyInfo.GetValue(dataItem);
                if (mapping.Format != null)
                {
                    if (dataValue is IFormattable formattable)
                    {
                        cell.SetValue(formattable.ToString(mapping.Format, CultureInfo.InvariantCulture));
                    }
                }
                cell.SetValue(XLCellValue.FromObject(dataValue));
                
                cell.Style.ApplyStyle(dataStyle);
            }

            rowIndex++;
        }
        
        return workbook;
    }

    /// <summary>
    /// Reads a workbook to convert it into the given model
    /// </summary>
    /// <param name="sheetName">name of the sheet to read from</param>
    /// <param name="workbook"></param>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    public Task<List<TModel>> ReadWorkbookAsync(string sheetName, XLWorkbook workbook)
    {
        var propertyData = GetModelMetaData();

        var sheet = workbook.Worksheet(sheetName);
        var usedArea = sheet.RangeUsed();

        //find header names based on TModel
        var headerNames = new HashSet<string>(propertyData.Where(x => !string.IsNullOrWhiteSpace(x.Name)).Select(x => x.Name)!);

        //find header row (so we can skip comments, etc)
        var headerRowIndex = usedArea
            .Rows(x => x.Cells()
                .All(c => headerNames.Contains(c.Value.ToString())))
            .FirstOrDefault()
            ?.RowNumber() ?? -1;

        var remainingRows = usedArea.Rows(headerRowIndex + 1, usedArea.RowCount()).ToList();

        var output = new List<TModel>();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = new TModel();

            for (var columnIndex = 0; columnIndex < propertyData.Count; columnIndex++)
            {
                var columnData = propertyData[columnIndex];
                var cell = row.Cell(columnIndex + 1 /* use +1 because Excel starts at 1 */);
                var dataValue = 
                    //fp types
                    TrySetValue<double>(columnData, cell) ??
                    TrySetValue<float>(columnData, cell) ??
                    TrySetValue<decimal>(columnData, cell) ??
                    
                    //integer types
                    TrySetValue<int>(columnData, cell) ??
                    TrySetValue<uint>(columnData, cell) ??
                    TrySetValue<byte>(columnData, cell) ??
                    TrySetValue<short>(columnData, cell) ??
                    TrySetValue<ushort>(columnData, cell) ??
                    
                    TrySetValue<long>(columnData, cell) ??
                    TrySetValue<ulong>(columnData, cell) ??
                    
                    //dates
                    TrySetValue<DateTime>(columnData, cell) ??
                    TrySetValue<DateTimeOffset>(columnData, cell) ??
                    
                    TrySetValue<string>(columnData, cell);

                if (columnData.PropertyInfo.PropertyType == dataValue?.GetType())
                {
                    columnData.PropertyInfo.SetValue(data, dataValue);
                }
            }

            output.Add(data);
        }

        return Task.FromResult(output);
    }

    /// <summary>
    /// Override this method to set cell style for header row cells.
    /// </summary>
    /// <returns></returns>
    public virtual SharpExcelCellStyle GetHeaderStyle()
    {
        return SharpExcelCellStyleConstants.DefaultHeaderStyle;
    }

    /// <summary>
    /// Override this method to set cell style for each data cell.
    /// Record + current column are provided so styling can be different based on conditions given by the user
    /// </summary>
    /// <param name="record">current record being processed</param>
    /// <param name="propertyName">current column being processed</param>
    /// <returns></returns>
    public virtual SharpExcelCellStyle GetDataStyle(string propertyName, TModel record)
    {
        return SharpExcelCellStyleConstants.DefaultDataStyle;
    }

    public async Task<HashSet<string>> GetOptionalPropertiesToExport(
        Func<string, Task<bool>>? optionalColumnFunc = null)
    {
        var output = new List<string>();

        var optionalProperties = GetModelMetaData().Where(x => x.Optional);

        foreach (var optional in optionalProperties)
        {
            if (optionalColumnFunc != null && await optionalColumnFunc(optional.PropertyInfo.Name))
            {
                output.Add(optional.PropertyInfo.Name);
            }
        }
        
        return [..output];
    }

    public virtual HashSet<string> GetOptionalColumns()
    {
        var propertyData = GetModelMetaData();
        var output = propertyData.Where(x => x.Optional)
            .Select(x => x.Name);

        return [..output];
    }

    static object? TrySetValue<TPropertyData>(PropertyData columnData, IXLCell cell)
    {
        if (columnData.PropertyInfo.PropertyType == typeof(TPropertyData))
        {
            if (cell.TryGetValue(out TPropertyData dataValue))
            {
                return dataValue;
            }
        }

        return null;
    }

    /// <summary>
    /// Reads model attributes and converts to column metadata
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <returns></returns>
    private List<PropertyData> GetModelMetaData()
    {
        var dataType = typeof(TModel);
        var propertyMappings = new List<PropertyData>();
        for (var columnIndex = 0; columnIndex < dataType.GetProperties().Length; columnIndex++)
        {

            var property = dataType.GetProperties()[columnIndex];
            
            var rowIdentifierAttribute = property.GetCustomAttribute<ExcelRowIdentifierAttribute>();
            if (rowIdentifierAttribute != null)
            {
                continue;
            }

            var attribute = property.GetCustomAttribute<ExcelColumnDefinitionAttribute>();
            var columnName = property.Name;
            if (!string.IsNullOrWhiteSpace(attribute?.DisplayName))
            {
                columnName = attribute?.DisplayName;
            }

            propertyMappings.Add(new PropertyData()
            {
                Name = columnName,
                PropertyInfo = property,
                Format = attribute?.Format,
                Optional = attribute?.IsOptional ?? false,
                ColumnWidth = attribute?.ColumnWidth ?? -1
            });
        }

        return propertyMappings;
    }

    private struct PropertyData
    {
        public string? Name { get; set; }

        public string? Format { get; set; }

        public int ColumnWidth { get; set; }

        public bool Optional { get; set; }

        public PropertyInfo PropertyInfo { get; set; }
    }
}