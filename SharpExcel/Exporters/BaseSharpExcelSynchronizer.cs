using System.ComponentModel.DataAnnotations;
using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using SharpExcel.Abstraction;
using SharpExcel.Extensions;
using SharpExcel.Models.Arguments;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Results;

namespace SharpExcel.Exporters;

/// <summary>
/// Base class for creating excel workbooks
/// </summary>
public class BaseSharpExcelSynchronizer<TModel> : ISharpExcelSynchronizer<TModel>
    where TModel : class, new()
{
    private readonly IOptions<ExporterOptions<TModel>> _options;

    public BaseSharpExcelSynchronizer(IOptions<ExporterOptions<TModel>> options)
    {
        _options = options;
    }
    /// <inheritdoc />
    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(ExcelArguments arguments, XLWorkbook workbook)
    {
        var parsedWorkbook = await ReadWorkbookAsync(arguments, workbook);
        ExporterHelpers.ApplyCellValidation(arguments.SheetName!, workbook, parsedWorkbook);
        return workbook;
    }


    /// <inheritdoc />
    public virtual async Task<XLWorkbook> GenerateWorkbookAsync(ExcelArguments arguments, IEnumerable<TModel> data)
    {
        var workbook = new XLWorkbook();
        
        var instanceData = CreateWriteInstanceData(arguments, workbook);

        if (instanceData.HeaderStyle.RowHeight.HasValue)
        {
            instanceData.MainWorksheet.Rows().Height = instanceData.HeaderStyle.RowHeight.Value;
        }

        //start at Row 1 because Excel starts at 1
        var rowIndex = 1;

        var dropdownDataMappings = EnumExporter.AddEnumDropdownMappingsToSheet(instanceData);

        ExporterHelpers.WriteHeaderRow(instanceData, rowIndex);

        //go to next row to start inserting data
        rowIndex++;
        
        foreach (var dataItem in data)
        {
            ExporterHelpers.WriteDataRow(instanceData, dataItem, rowIndex, dropdownDataMappings);
            rowIndex++;
        }
        
        return await Task.FromResult(workbook);
    }

    /// <inheritdoc />
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(ExcelArguments arguments, XLWorkbook workbook)
    {
        var instanceData = CreateReadInstanceData(arguments, workbook);
        
        var output = new ExcelReadResult<TModel>();

        var usedArea = instanceData.MainWorksheet.RangeUsed();
        var headerRowIndex = FindAndMapHeaderRow(instanceData, usedArea);
        var remainingRows = usedArea.Rows(headerRowIndex + 1, usedArea.RowCount()).ToList();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = ReadRow(instanceData, row, out var validationResults);

            if (data == null)
            {
                //skip to next record if we can read record
                continue;
            }

            output.Records.Add(data);
            
            //add validation results
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

    /// <summary>
    /// Reads a row and tries to convert it to the given model
    /// </summary>
    /// <param name="instance">instance data</param>
    /// <param name="row">row to read</param>
    /// <param name="validationResults">A dictionary containing validatio nresults of previous rows</param>
    /// <returns></returns>
    private static TModel? ReadRow(
        SharpExcelWriterInstanceData<TModel> instance,
        IXLRangeRow row,
        out Dictionary<ExcelAddress, List<ValidationResult>> validationResults)
    {
        var data = new TModel();

        validationResults = new();

        foreach (var columnData in instance.Properties.PropertyMappings.OrderBy(x => x.ColumnData.ColumnIndex))
        {
            var cell = row.Cell(columnData.ColumnData.ColumnIndex);

            var excelAddress = new ExcelAddress()
            {
                RowNumber = row.RowNumber(),
                ColumnId = cell.Address.ColumnNumber,
                ColumnName = cell.Address.ColumnLetter,
                HeaderName = columnData.Name
            };

            var dataValue = ExporterHelpers.TrySetCellValue(instance.Properties.EnumMappings, columnData, cell,
                instance.CultureInfo ?? CultureInfo.CurrentCulture);

            if (columnData.PropertyInfo.PropertyType == dataValue?.GetType())
            {
                columnData.PropertyInfo.SetValue(data, dataValue);

                ValidationContext context = new ValidationContext(data)
                {
                    MemberName = columnData.PropertyInfo.Name,
                    DisplayName = columnData.Name ?? columnData.PropertyInfo.Name
                };
                var validations = new List<ValidationResult>();
                if (!Validator.TryValidateProperty(dataValue, context, validations))
                {
                    validationResults.Add(excelAddress, validations);
                }
            }
        }

        return data;
    }

    /// <summary>
    /// Finds header row and maps the column order so we can fill them later
    /// </summary>
    /// <param name="instance">instance data</param>
    /// <param name="usedArea">total used area of the workbook</param>
    /// <param name="sheet"></param>
    /// <returns></returns>
    private static int FindAndMapHeaderRow(
        SharpExcelWriterInstanceData<TModel> instance,
        IXLRange usedArea)
    {
        var headerNames = new HashSet<string>(instance.Properties.PropertyMappings.Where(
                    x => !string.IsNullOrWhiteSpace(x.NormalizedName))
                .Select(x => x.NormalizedName?.ToLowerInvariant())!
        );
        //find header row
        var headerRowIndex = usedArea
            .Rows(x => x.Cells()
                .Any(c => headerNames.Contains(c.Value.ToString().ToLowerInvariant())))
            .FirstOrDefault()
            ?.RowNumber() ?? -1;

        var propertiesByColumnName = instance.Properties.PropertyMappings.ToDictionary(x => x.NormalizedName);

        foreach (var cell in instance.MainWorksheet.Row(headerRowIndex).Cells())
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
    
    /// <summary>
    /// Creates lookup data for the current export run
    /// </summary>
    /// <param name="arguments">arguments to use</param>
    /// <param name="workbook">workbook to use</param>
    /// <returns></returns>
    private SharpExcelWriterInstanceData<TModel> CreateWriteInstanceData(ExcelArguments arguments, XLWorkbook workbook)
    {
        var random = new Random();
        var randomNumber = random.Next(0, 1000000);

        var run = new SharpExcelWriterInstanceData<TModel>()
        {
            DataStyle = _options.Value.Styling.DefaultDataStyle,
            HeaderStyle = _options.Value.Styling.DefaultHeaderStyle,
            ErrorStyle = _options.Value.Styling.DefaultErrorStyle,
            Properties = TypeMapper.GetModelMetaData<TModel>(),
            StylingRuleLookup = _options.Value.Styling.ToStylingRuleLookup(),
            MainWorksheet = workbook.AddWorksheet(arguments.SheetName),
            DropdownSourceWorksheet = workbook.AddWorksheet("Dropdowns_" + randomNumber.ToString("000000")).Hide(),
            CultureInfo = arguments.CultureInfo
        };
        
        return run;
    }
    
    /// <summary>
    /// Creates lookup data for the current import run
    /// </summary>
    /// <param name="arguments">arguments to use</param>
    /// <param name="workbook">workbook to use</param>
    /// <returns></returns>
    private SharpExcelWriterInstanceData<TModel> CreateReadInstanceData(ExcelArguments arguments, XLWorkbook workbook)
    {
       return new SharpExcelWriterInstanceData<TModel>()
        {
            DataStyle = _options.Value.Styling.DefaultDataStyle,
            HeaderStyle = _options.Value.Styling.DefaultHeaderStyle,
            ErrorStyle = _options.Value.Styling.DefaultErrorStyle,
            Properties = TypeMapper.GetModelMetaData<TModel>(),
            StylingRuleLookup = _options.Value.Styling.ToStylingRuleLookup(),
            MainWorksheet = workbook.Worksheet(arguments.SheetName),
            CultureInfo = arguments.CultureInfo
        };
    }
}