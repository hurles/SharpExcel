using System.ComponentModel.DataAnnotations;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using SharpExcel.Abstraction;
using SharpExcel.Extensions;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling;
using SharpExcel.Models.Styling.Rules;

namespace SharpExcel.Exporters;

/// <summary>
/// Base class for creating excel workbooks
/// </summary>
public partial class BaseExcelExporter<TModel> : IExcelExporter<TModel>
    where TModel : class, new()
{
    private readonly IOptions<ExporterOptions<TModel>> _options;

    public BaseExcelExporter(IOptions<ExporterOptions<TModel>> options)
    {
        _options = options;
    }
    /// <inheritdoc />
    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(SharpExcelArguments arguments, XLWorkbook workbook)
    {
        var parsedWorkbook = await ReadWorkbookAsync(arguments, workbook);
        ExporterHelpers.ApplyCellValidation(arguments.SheetName!, workbook, parsedWorkbook);
        return workbook;
    }


    /// <inheritdoc />
    public virtual async Task<XLWorkbook> GenerateWorkbookAsync(SharpExcelArguments arguments, IEnumerable<TModel> data)
    {
        var workbook = new XLWorkbook();

        var run = new SharpExcelWriterInstanceData<TModel>()
        {
            DataStyle = _options.Value.Styling.DefaultDataStyle,
            HeaderStyle = _options.Value.Styling.DefaultHeaderStyle,
            ErrorStyle = _options.Value.Styling.DefaultErrorStyle,
            Properties = TypeMapper.GetModelMetaData<TModel>(),
            StylingRuleLookup = _options.Value.Styling.ToStylingRuleLookup(),
            MainWorksheet = workbook.AddWorksheet(arguments.SheetName),
            DropdownSourceWorksheet = workbook.AddWorksheet(ExcelParseHelper.GetDropdownDataSheetName()).Hide(),
            CultureInfo = arguments.CultureInfo
        };
        
        if (run.HeaderStyle.RowHeight.HasValue)
        {
            run.MainWorksheet.Rows().Height = run.HeaderStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var dropdownDataMappings = EnumExporter.AddEnumDropdownMappingsToSheet(run);

        ExporterHelpers.WriteHeaderRow(run, rowIndex);

        //go to next row to start inserting data
        rowIndex++;
        
        foreach (var dataItem in data)
        {
            ExporterHelpers.WriteDataRow(run, dataItem, rowIndex, dropdownDataMappings);
            rowIndex++;
        }
        
        return await Task.FromResult(workbook);
    }
    
    /// <inheritdoc />
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(SharpExcelArguments arguments, XLWorkbook workbook)
    {
        var output = new ExcelReadResult<TModel>();
        var propertyData = TypeMapper.GetModelMetaData<TModel>();

        var sheet = workbook.Worksheet(arguments.SheetName);
        var usedArea = sheet.RangeUsed();

        var headerRowIndex = FindAndMapHeaderRow(usedArea, propertyData, sheet);

        var remainingRows = usedArea.Rows(headerRowIndex + 1, usedArea.RowCount()).ToList();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = new TModel();

            Dictionary<ExcelAddress, List<ValidationResult>> validationResults = new();

            foreach (var columnData in  propertyData.PropertyMappings.OrderBy(x => x.ColumnData.ColumnIndex))
            {
                var cell = row.Cell(columnData.ColumnData.ColumnIndex);

                var excelAddress = new ExcelAddress()
                {
                    RowNumber = row.RowNumber(),
                    ColumnId = cell.Address.ColumnNumber,
                    ColumnName = cell.Address.ColumnLetter,
                    HeaderName = columnData.Name
                };
                
                var dataValue = ExporterHelpers.TrySetCellValue(propertyData.EnumMappings, columnData, cell, arguments.CultureInfo ?? CultureInfo.CurrentCulture);

                if (columnData.PropertyInfo.PropertyType == dataValue?.GetType())
                {
                    columnData.PropertyInfo.SetValue(data, dataValue);
                    
                    ValidationContext context = new ValidationContext(data) { 
                        MemberName = columnData.PropertyInfo.Name, 
                        DisplayName = columnData.Name ?? columnData.PropertyInfo.Name};
                    var validations = new List<ValidationResult>();
                    if (!Validator.TryValidateProperty(dataValue, context, validations))
                    {
                        validationResults.Add(excelAddress, validations);
                    }
                }
            }
            
            output.Records.Add(data);
            if (validationResults.Any())
            {
                foreach (var validationResult in validationResults)
                {
                    output.ValidationResults.Add(data, new ExcelCellValidationResult()
                    {
                        Address = validationResult.Key,
                        ValidationResults = validationResult.Value
                    });
                }
            }
            
        }

        return Task.FromResult(output);
    }

    private static int FindAndMapHeaderRow(IXLRange usedArea, PropertyDataCollection propertyData,
        IXLWorksheet sheet)
    {
        var headerNames = new HashSet<string>(propertyData.PropertyMappings.Where(
                    x => !string.IsNullOrWhiteSpace(x.NormalizedName))
                .Select(x => x.NormalizedName?.ToLowerInvariant())!
        );
        //find header row
        var headerRowIndex = usedArea
            .Rows(x => x.Cells()
                .Any(c => headerNames.Contains(c.Value.ToString().ToLowerInvariant())))
            .FirstOrDefault()
            ?.RowNumber() ?? -1;

        var propertiesByColumnName = propertyData.PropertyMappings.ToDictionary(x => x.NormalizedName);

        foreach (var cell in sheet.Row(headerRowIndex).Cells())
        {
            if (!cell.TryGetValue(out string cellValue))
                continue;

            cellValue = cellValue.Trim().ToLowerInvariant();
            
            if (!headerNames.Contains(cellValue)) 
                continue;
            
            if (propertiesByColumnName.ContainsKey(cellValue))
            {
                propertiesByColumnName[cellValue].ColumnData = new()
                {
                    ColumnName = propertiesByColumnName[cellValue].Name ?? string.Empty,
                    ColumnIndex = cell.Address.ColumnNumber
                };
            }
        }

        return headerRowIndex;
    }
}