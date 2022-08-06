using LevenshteinDistanceAlgorithm;
using OfficeOpenXml;
using System.Linq;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
string mainFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Excel Working Files");

using var allStocksMaliplusFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "item list kfa with stocks copy.xlsx")));
using var nyahururuDifferenceFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "item difference Nyahururu Branch.xlsx")));
using var oldItemCodesFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "matching rows.xlsx")));
using var unClarifiedItemCodesFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "not matching.xlsx")));

using var allStocksMaliplus = allStocksMaliplusFile.Workbook.Worksheets[0];
using var nyahururuDifference = nyahururuDifferenceFile.Workbook.Worksheets[0];
using var oldItemCodes = oldItemCodesFile.Workbook.Worksheets[0];
using var unClarifiedItemCodes = unClarifiedItemCodesFile.Workbook.Worksheets[0];

List<ItemCode> allItemsCodes = new();
for (int i = oldItemCodes.Dimension.Start.Row; i < oldItemCodes.Dimension.End.Row; i++)
{
	try
	{
        if(CustomValidations.IsValidItemCode(oldItemCodes.Cells[i, 1].Value?.ToString() ?? ""))
		{
			allItemsCodes.Add(new()
			{
				Code = oldItemCodes.Cells[i, 1].Value?.ToString(),
				Name = oldItemCodes.Cells[i, 2].Value?.ToString(),
				Distributor = oldItemCodes.Cells[i, 3].Value?.ToString(),
				IsVerified = true
			});
        }
    }
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
	}
}





//using var objs = package.Workbook.Worksheets[0];


Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("WeEd", "weed"));


Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wed", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wedfsdf", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wasded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wsded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wsded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wded", "weed"));


