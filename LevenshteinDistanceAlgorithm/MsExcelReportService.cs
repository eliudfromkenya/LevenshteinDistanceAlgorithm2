using OfficeOpenXml;
using System.Drawing;
using System.IO;

namespace LevenshteinDistanceAlgorithm;

public class MsExcelReportService
{
    private int currentRow = 3;

    public string GenerateMatchReport(ExcelPackage excelPackage, List<ItemCodeMatch> itemsMatched, string mainFolder, string branch)
    {
        currentRow = 3;
        var matchCodes = $"{branch} Matches";
        var sheet = excelPackage.Workbook.Worksheets.Add(matchCodes);

        sheet.Cells["A1:G1"].Merge = true;
        sheet.Cells["A1"].Value = matchCodes;
        sheet.Row(1).Height = 20;
        sheet.Row(1).Style.Font.Size = 20;
        sheet.Row(1).Style.Font.Color.SetColor(Color.Purple);
        sheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(1).Style.Font.Bold = true;

        sheet.Row(currentRow).Height = 20;
        sheet.Row(currentRow).Style.Font.Size = 12;
        sheet.Row(currentRow).Style.Font.Color.SetColor(Color.DarkGray);
        sheet.Row(currentRow).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(currentRow).Style.Font.Bold = true;

        var range = sheet.Cells[currentRow, 1, currentRow, 3];
        range.Style.Font.Color.SetColor(Color.RebeccaPurple);
        range.Style.Font.UnderLine = true;
        range.Style.Font.Size = 14;
        range.Merge = true;
        range.Value = $"{branch} Item Codes";

        range = sheet.Cells[currentRow, 4, currentRow, 7];
        range.Style.Font.Color.SetColor(Color.RebeccaPurple);
        range.Style.Font.UnderLine = true;
        range.Style.Font.Size = 14;
        range.Merge = true;
        range.Value = "Matched Item Codes";
        currentRow++;

        sheet.Cells[currentRow, 1].Value = "Item Code";
        sheet.Cells[currentRow, 2].Value = "Description";
        sheet.Cells[currentRow, 3].Value = "Quantity";

        sheet.Cells[currentRow, 4].Value = "Item Code";
        sheet.Cells[currentRow, 5].Value = "Description";
        sheet.Cells[currentRow, 6].Value = "Match Strength";
        sheet.Cells[currentRow, 7].Value = "Already Verified";

        sheet.Cells[currentRow, 4, currentRow, 7].Style.Font.Color.SetColor(Color.RebeccaPurple);
        sheet.Cells[currentRow, 1, currentRow, 7].Style.Font.UnderLine = true;
        currentRow++;

        itemsMatched.OrderBy(c => c.MatchStrength)
            .ToList().ForEach(col =>
        {
            if (!col?.MatchedCode?.IsVerified ?? false)
                sheet.Cells[currentRow, 1, currentRow, 7].Style.Fill.SetBackground(Color.Lavender);
            sheet.Cells[currentRow, 1].Value = col?.OriginalCode?.Code;
            sheet.Cells[currentRow, 2].Value = col?.OriginalCode?.Name;
            sheet.Cells[currentRow, 3].Value = col?.OriginalCode?.Quantity;

            if (col?.MatchedCode != null)
            {
                sheet.Cells[currentRow, 4].Value = col?.MatchedCode?.Code;
                sheet.Cells[currentRow, 5].Value = col?.MatchedCode?.Name;
                sheet.Cells[currentRow, 6].Value = col?.MatchStrength;
                sheet.Cells[currentRow, 7].Value = (col?.MatchedCode.IsVerified ?? false) ? "Yes" : "No";
                sheet.Cells[currentRow, 4].Style.Font.Bold = true;
            }
            range = sheet.Cells[currentRow, 4, currentRow, 6];
            range.Style.Font.Color.SetColor(Color.DarkRed);
            currentRow++;
        });

        var file = Path.Combine(mainFolder, $"{branch}.xlsx");
        //excelPackage.SaveAs(new FileInfo(file), "Zanas2022");
        excelPackage.SaveAs(new FileInfo(file));
        return file;
    }



    public string GenerateGroupsCodeReport(ExcelPackage excelPackage, List<ItemCode> items, string mainFolder, string branch)
    {
        currentRow = 3;
        var matchCodes = $"Master Item Code Groups";
        var sheet = excelPackage.Workbook.Worksheets.Add(matchCodes);

        sheet.Cells["A1:G1"].Merge = true;
        sheet.Cells["A1"].Value = matchCodes;
        sheet.Row(1).Height = 20;
        sheet.Row(1).Style.Font.Size = 20;
        sheet.Row(1).Style.Font.Color.SetColor(Color.Purple);
        sheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(1).Style.Font.Bold = true;

        sheet.Row(currentRow).Height = 20;
        sheet.Row(currentRow).Style.Font.Size = 12;
        sheet.Row(currentRow).Style.Font.Color.SetColor(Color.DarkGray);
        sheet.Row(currentRow).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(currentRow).Style.Font.Bold = true;

        var range = sheet.Cells[currentRow, 1, currentRow, 6];
        range.Style.Font.Color.SetColor(Color.RebeccaPurple);
        range.Style.Font.UnderLine = true;
        range.Style.Font.Size = 14;
        range.Merge = true;
        range.Value = matchCodes;
        currentRow++;

        sheet.Cells[currentRow, 1].Value = "Item Code";
        sheet.Cells[currentRow, 2].Value = "Description";
        sheet.Cells[currentRow, 3].Value = "Quantity";
        sheet.Cells[currentRow, 4].Value = "Unit Of Sale";
        sheet.Cells[currentRow, 5].Value = "Item Group Name";
        sheet.Cells[currentRow, 6].Value = "Is Duplicate";

        sheet.Cells[currentRow, 4, currentRow, 6].Style.Font.Color.SetColor(Color.RebeccaPurple);
        sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.UnderLine = true;
        currentRow++;

        items
            .GroupBy(v => v.HarmonizedGroupName)
            .Where(c => c.Count() > 1)            
            .OrderBy(c => c.FirstOrDefault()?.GroupName)
            .ToList().ForEach(grp =>
            {
                range = sheet.Cells[currentRow, 1, currentRow, 6];
                range.Merge = true;
                range.Value = grp.First().GroupName;
                range.Style.Font.Color.SetColor(Color.DarkRed);
                currentRow++;

                List<string> doneCols = new();

                foreach (var col in grp
                   .OrderBy(c => c.Quantity)
                   .ThenBy(c => c.MeasureUnit)
                   .ThenBy(x => x.Name))
                {
                    if (!col?.IsVerified ?? false)
                        sheet.Cells[currentRow, 1, currentRow, 5].Style.Fill.SetBackground(Color.Lavender);
                    sheet.Cells[currentRow, 1].Value = col?.Code;
                    sheet.Cells[currentRow, 2].Value = col?.Name;
                    sheet.Cells[currentRow, 3].Value = col?.Quantity;
                    sheet.Cells[currentRow, 4].Value = col?.MeasureUnit;
                    sheet.Cells[currentRow, 5].Value = col?.GroupName;
                    sheet.Cells[currentRow, 6].Value = doneCols.Contains(col?.Name ?? "") ? "Yes" : "No";
                    doneCols.Add(col?.Name ?? "");

                    currentRow++;
                }
                currentRow += 3;
            });

        var file = Path.Combine(mainFolder, $"{branch}.xlsx");
        //excelPackage.SaveAs(new FileInfo(file), "Zanas2022");
        excelPackage.SaveAs(new FileInfo(file));
        return file;
    }


    public string GenerateDulplicatedItemCodesReport(ExcelPackage excelPackage, List<ItemCode> items, string mainFolder, string branch)
    {
        currentRow = 3;
        var matchCodes = $"Duplicated Item Codes";
        var sheet = excelPackage.Workbook.Worksheets.Add(matchCodes);

        sheet.Cells["A1:G1"].Merge = true;
        sheet.Cells["A1"].Value = matchCodes;
        sheet.Row(1).Height = 20;
        sheet.Row(1).Style.Font.Size = 20;
        sheet.Row(1).Style.Font.Color.SetColor(Color.Purple);
        sheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(1).Style.Font.Bold = true;

        sheet.Row(currentRow).Height = 20;
        sheet.Row(currentRow).Style.Font.Size = 12;
        sheet.Row(currentRow).Style.Font.Color.SetColor(Color.DarkGray);
        sheet.Row(currentRow).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(currentRow).Style.Font.Bold = true;

        var range = sheet.Cells[currentRow, 1, currentRow, 5];
        range.Style.Font.Color.SetColor(Color.RebeccaPurple);
        range.Style.Font.UnderLine = true;
        range.Style.Font.Size = 14;
        range.Merge = true;
        range.Value = matchCodes;
        currentRow++;

        sheet.Cells[currentRow, 1].Value = "Item Code";
        sheet.Cells[currentRow, 2].Value = "Description";
        sheet.Cells[currentRow, 3].Value = "Quantity";
        sheet.Cells[currentRow, 4].Value = "Unit Of Sale";
        sheet.Cells[currentRow, 5].Value = "Is Duplicate";

        sheet.Cells[currentRow, 4, currentRow, 5].Style.Font.Color.SetColor(Color.RebeccaPurple);
        sheet.Cells[currentRow, 1, currentRow, 5].Style.Font.UnderLine = true;
        currentRow++;

        items
            .GroupBy(v => v.HarmonizedName)
            .Where(x => x.Count() > 1)
            .OrderByDescending(c => c?.Count())
            .ToList().ForEach(grp =>
            {
                List<string> doneCols = new();
                foreach (var col in grp
                   .OrderByDescending(c => c.Quantity)
                   .ThenBy(x => x.MeasureUnit)
                   .ThenBy(x => x.Name))
                {
                    var isDuplicate = doneCols.Contains((col?.HarmonizedName ?? ""));
                     doneCols.Add((col?.HarmonizedName ?? ""));

                    if (!col?.IsVerified ?? false)
                        sheet.Cells[currentRow, 1, currentRow, 5].Style.Fill.SetBackground(Color.Lavender);
                    sheet.Cells[currentRow, 1].Value = col?.Code;
                    sheet.Cells[currentRow, 2].Value = col?.Name;
                    sheet.Cells[currentRow, 3].Value = col?.Quantity;
                    sheet.Cells[currentRow, 4].Value = col?.MeasureUnit;
                    sheet.Cells[currentRow, 5].Value = isDuplicate ? "Yes" : "No";
                   
                    if (isDuplicate)
                    {
                        range = sheet.Cells[currentRow, 1, currentRow, 5];
                        range.Style.Font.Color.SetColor(Color.DarkRed);
                    }
                    currentRow++;
                }
                currentRow += 1;
            });

        var file = Path.Combine(mainFolder, $"{branch}.xlsx");
        excelPackage.SaveAs(new FileInfo(file));
        return file;
    }

    public string GenerateDulplicatedItemCodesOnlyReport(ExcelPackage excelPackage, List<ItemCode> items, string mainFolder, string branch)
    {
        currentRow = 3;
        var matchCodes = $"Duplicated Item Codes Only";
        var sheet = excelPackage.Workbook.Worksheets.Add(matchCodes);

        sheet.Cells["A1:G1"].Merge = true;
        sheet.Cells["A1"].Value = matchCodes;
        sheet.Row(1).Height = 20;
        sheet.Row(1).Style.Font.Size = 20;
        sheet.Row(1).Style.Font.Color.SetColor(Color.Purple);
        sheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(1).Style.Font.Bold = true;

        sheet.Row(currentRow).Height = 20;
        sheet.Row(currentRow).Style.Font.Size = 12;
        sheet.Row(currentRow).Style.Font.Color.SetColor(Color.DarkGray);
        sheet.Row(currentRow).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        sheet.Row(currentRow).Style.Font.Bold = true;

        var range = sheet.Cells[currentRow, 1, currentRow, 5];
        range.Style.Font.Color.SetColor(Color.RebeccaPurple);
        range.Style.Font.UnderLine = true;
        range.Style.Font.Size = 14;
        range.Merge = true;
        range.Value = matchCodes;
        currentRow++;

        sheet.Cells[currentRow, 1].Value = "Item Code";
        sheet.Cells[currentRow, 2].Value = "Description";
        sheet.Cells[currentRow, 3].Value = "Quantity";
        sheet.Cells[currentRow, 4].Value = "Unit Of Sale";
        sheet.Cells[currentRow, 5].Value = "Is Duplicate";

        sheet.Cells[currentRow, 4, currentRow, 5].Style.Font.Color.SetColor(Color.RebeccaPurple);
        sheet.Cells[currentRow, 1, currentRow, 5].Style.Font.UnderLine = true;
        currentRow++;

        items
            .GroupBy(v => v.HarmonizedName)
            .OrderBy(c => c.FirstOrDefault()?.HarmonizedName)
            .ToList().ForEach(grp =>
            {
                List<string> doneCols = new();
                foreach (var col in grp
                   .OrderBy(c => c.MeasureUnit)
                   .ThenBy(x => x.Quantity)
                   .ThenBy(x => x.HarmonizedName))
                {
                    var isDuplicate = doneCols.Contains(col?.Name ?? "");

                    if (isDuplicate)
                    {
                        if (!col?.IsVerified ?? false)
                            sheet.Cells[currentRow, 1, currentRow, 5].Style.Fill.SetBackground(Color.Lavender);
                        sheet.Cells[currentRow, 1].Value = col?.Code;
                        sheet.Cells[currentRow, 2].Value = col?.Name;
                        sheet.Cells[currentRow, 3].Value = col?.Quantity;
                        sheet.Cells[currentRow, 4].Value = col?.MeasureUnit;
                        sheet.Cells[currentRow, 5].Value = isDuplicate ? "Yes" : "No";
                        currentRow++;
                    }
                    doneCols.Add(col?.Name ?? "");                    
                }
            });

        var file = Path.Combine(mainFolder, $"{branch}.xlsx");
        excelPackage.SaveAs(new FileInfo(file));
        return file;
    }
}