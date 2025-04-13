using System.ComponentModel.DataAnnotations;
using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Options;
using SharpExcel.Abstraction;
using SharpExcel.Extensions;
using SharpExcel.Models.Configuration;
using SharpExcel.Models.Data;
using SharpExcel.Models.Results;
using SharpExcel.Models.Styling.Constants;

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
    public async Task<XLWorkbook> ValidateAndAnnotateWorkbookAsync(CultureInfo cultureInfo, XLWorkbook workbook)
    {
        var parsedWorkbook = await ReadWorkbookAsync(cultureInfo, workbook);
        ExporterHelpers.ApplyCellValidation(workbook, parsedWorkbook);
        return workbook;
    }
    
    public Task<ExcelReadResult<TModel>> ReadWorkbookAsync(CultureInfo cultureInfo, XLWorkbook workbook)
    {
        var output = new ExcelReadResult<TModel>();
        var instanceData = CreateReadInstanceData(cultureInfo, workbook);

        var rules = _options.Value.Targeting.Rules.GroupBy(rule => rule.SheetName);

        foreach (var ruleGroup in rules)
        {
            if (!instanceData.Workbook.Worksheets.TryGetWorksheet(ruleGroup.Key, out var worksheet))
            {
                continue;
            }

            foreach (var rule in ruleGroup)
            {
                ReadSheetAsync(rule, output, instanceData, worksheet);
            }
        }

        return Task.FromResult(output);
    }

    public virtual async Task<XLWorkbook> GenerateWorkbookAsync(CultureInfo cultureInfo, ICollection<TModel> data)
    {
        var workbook = new XLWorkbook();

        if (!_options.Value.Targeting.Rules.Any())
        {
            _options.Value.Targeting.Rules = [ExcelTargetingConstants<TModel>.DefaultTargetingRule];
        }

        Dictionary<TargetingRule<TModel>, IEnumerable<TModel>> dataGroupedByTargetingRule = new();

        foreach (var targetingRule in _options.Value.Targeting.Rules)
        {
            dataGroupedByTargetingRule.Add(targetingRule, data.Where(x => targetingRule.RulePredicate != null && targetingRule.RulePredicate(x)).ToList());
            if (!workbook.Worksheets.TryGetWorksheet(targetingRule.SheetName, out var _))
            {
                workbook.Worksheets.Add(targetingRule.SheetName);
            }
        }
        
        var instanceData = CreateWriteInstanceData(cultureInfo, workbook);
        EnumExporter.AddEnumDropdownMappingsToSheet(instanceData);

        foreach (var targetingRuleData in dataGroupedByTargetingRule)
        {
            await GenerateSheetAsync(targetingRuleData.Key, instanceData, targetingRuleData.Value);
        }


        return workbook;
    }


    /// <inheritdoc />
    internal virtual Task GenerateSheetAsync(TargetingRule<TModel> targetingRule, SharpExcelWriterInstanceData<TModel> instanceData, IEnumerable<TModel> data)
    {
        if (!instanceData.Workbook.Worksheets.TryGetWorksheet(targetingRule.SheetName, out var _))
        {
            instanceData.Workbook.Worksheets.Add(targetingRule.SheetName);
        }

        //start at Row 1 if not defined because Excel starts at 1
        var rowIndex = targetingRule.Row ?? 1;

        ExporterHelpers.WriteHeaderRow(targetingRule, instanceData, rowIndex, targetingRule.Column);

        //go to next row to start inserting data
        rowIndex++;
        
        foreach (var dataItem in data)
        {
            ExporterHelpers.WriteDataRow(targetingRule, instanceData, dataItem, rowIndex, targetingRule.Column);
            rowIndex++;
        }

        return Task.CompletedTask;
    }

    /// <inheritdoc />
    internal void ReadSheetAsync(TargetingRule<TModel> rule, ExcelReadResult<TModel> result, SharpExcelWriterInstanceData<TModel> instanceData, IXLWorksheet worksheet)
    {
        var usedArea = worksheet.RangeUsed();
        if (usedArea is null)
        {
            return;
        }
        
        var headerRowIndex = FindAndMapHeaderRow(rule, instanceData, usedArea);
        var remainingRows = usedArea.Rows(headerRowIndex, usedArea.RowCount()).ToList();

        //parse remaining data rows
        foreach (var row in remainingRows)
        {
            var data = ReadRow(worksheet, instanceData, row.WorksheetRow(), out var validationResults);

            if (data == null)
            {
                //skip to next record if we can't read record
                continue;
            }

            result.Records.Add(data);
            
            //add validation results
            if (validationResults.Any())
            {
                foreach (var validationResult in validationResults)
                {
                    result.ValidationResults.Add(data, new ExcelCellValidationResult()
                    {
                        Address = validationResult.Key,
                        ValidationResults = validationResult.Value
                    });
                }
            }
            
        }
    }

    /// <summary>
    /// Reads a row and tries to convert it to the given model
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="instance">instance data</param>
    /// <param name="row">row to read</param>
    /// <param name="validationResults">A dictionary containing validation results of previous rows</param>
    /// <returns></returns>
    private static TModel? ReadRow(
        IXLWorksheet sheet,
        SharpExcelWriterInstanceData<TModel> instance,
        IXLRow row,
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
                HeaderName = columnData.Name,
                SheetName = sheet.Name
            };

            var dataValue = ExporterHelpers.TryGetCellValue(instance.Properties.EnumMappings, columnData, cell,
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
        TargetingRule<TModel> rule,
        SharpExcelWriterInstanceData<TModel> instance,
        IXLRange usedArea)
    {
        var headerNames = new HashSet<string>(instance.Properties.PropertyMappings.Where(
                    x => !string.IsNullOrWhiteSpace(x.NormalizedName))
                .Select(x => x.NormalizedName?.ToLowerInvariant())!
        );
        
        //find header row
        var headerRow = usedArea
            .Rows(x => x.Cells()
                .Any(c => headerNames.Contains(c.Value.ToString().ToLowerInvariant())))
            .FirstOrDefault()?.WorksheetRow();

        var propertiesByColumnName = instance.Properties.PropertyMappings.ToDictionary(x => x.NormalizedName);

        var startIndex = usedArea.FirstCell().WorksheetColumn().ColumnNumber();
        
        
        
        if (rule.Column != null && rule.Column > startIndex)
            startIndex = rule.Column ?? 1;

        for (int i = startIndex; i <= usedArea.ColumnCount(); i++)
        {
            var cell = headerRow!.Cell(i);
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

        return headerRow!.RowNumber() - 1;
    }
    
    /// <summary>
    /// Creates lookup data for the current export run
    /// </summary>
    /// <param name="arguments">arguments to use</param>
    /// <param name="workbook">workbook to use</param>
    /// <returns></returns>
    private SharpExcelWriterInstanceData<TModel> CreateWriteInstanceData(CultureInfo cultureInfo, XLWorkbook workbook)
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
            Workbook = workbook,
            DropdownSourceWorksheet = workbook.AddWorksheet("Dropdowns_" + randomNumber.ToString("000000")).Hide(),
            CultureInfo = cultureInfo
        };
        
        return run;
    }
    
    /// <summary>
    /// Creates lookup data for the current import run
    /// </summary>
    /// <param name="arguments">arguments to use</param>
    /// <param name="workbook">workbook to use</param>
    /// <returns></returns>
    private SharpExcelWriterInstanceData<TModel> CreateReadInstanceData(CultureInfo cultureInfo, XLWorkbook workbook)
    {
       return new SharpExcelWriterInstanceData<TModel>()
        {
            DataStyle = _options.Value.Styling.DefaultDataStyle,
            HeaderStyle = _options.Value.Styling.DefaultHeaderStyle,
            ErrorStyle = _options.Value.Styling.DefaultErrorStyle,
            Properties = TypeMapper.GetModelMetaData<TModel>(),
            StylingRuleLookup = _options.Value.Styling.ToStylingRuleLookup(),
            TargetingRules = _options.Value.Targeting.Rules,
            Workbook = workbook,
            CultureInfo = cultureInfo
        };
    }
}