using OfficeOpenXml;
using System.Drawing;
using System.IO;

namespace LevenshteinDistanceAlgorithm;

public class MsExcelReportService
{
    private int currentRow = 3;

    public string GenerateMatchReport(List<ItemCodeMatch> itemsMatched, string mainFolder, string branch)
    {
        using var excelPackage = new ExcelPackage();
        excelPackage.Workbook.Properties.Author = "KFA / Primesoft Implimentation Team";
        excelPackage.Workbook.Properties.Title = "Zanas to Maliplus Item Master Matching Report";
        excelPackage.Workbook.Properties.Subject = "Automated Data Matching";
        excelPackage.Workbook.Properties.Created = DateTime.Now;
        excelPackage.Workbook.Properties.Company = "KFA LTD";

        var title = $"{branch} Item Codes Match Information";

        currentRow = 3;
        var sheet = excelPackage.Workbook.Worksheets.Add(title); 

        sheet.Cells["A1:G1"].Merge = true;
        sheet.Cells["A1"].Value = title;
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
        // sheet.Row(currentRow).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightTrellis;

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
                sheet.Cells[currentRow, 7].Value = (col?.MatchedCode.IsVerified??false)?"Yes":"No";
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
}