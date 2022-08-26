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

List<ItemCode> allItemsCodes = new(), unCleanItemCodes = new();
for (int i = oldItemCodes.Dimension.Start.Row; i < oldItemCodes.Dimension.End.Row; i++)
{
	try
	{
        if(CustomValidations.IsValidItemCode(oldItemCodes.Cells[i, 1].Value?.ToString() ?? ""))
		{
			allItemsCodes.Add(new()
			{
				Code = oldItemCodes.Cells[i, 1].Value?.ToString()?.Trim(),
				Name = oldItemCodes.Cells[i, 2].Value?.ToString(),
				Distributor = oldItemCodes.Cells[i, 3].Value?.ToString(),
				IsVerified = true
			});
        }
    }
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}


for (int i = allStocksMaliplus.Dimension.Start.Row; i < allStocksMaliplus.Dimension.End.Row; i++)
{
	try
	{
		var itemCode =
			allStocksMaliplus.Cells[i, 1].Value?.ToString() ?? "";

		decimal qty = 0;
		try
		{
			if (decimal.TryParse(allStocksMaliplus.Cells[i, 3].Value?.ToString() ?? "", out decimal num))
				qty = num;
		}
		catch { }
		try
		{
			if (qty == 0 && decimal.TryParse(allStocksMaliplus.Cells[i, 4].Value?.ToString() ?? "", out decimal num))
				qty = num;
		}
		catch { }

		if (CustomValidations.IsValidItemCode(itemCode))
		{
			var item = allItemsCodes.FirstOrDefault(x => x.Code == itemCode);
			if (item == null)
			{
				allItemsCodes.Add(new()
				{
					Quantity = qty,
					Code = allStocksMaliplus.Cells[i, 1].Value?.ToString()?.Trim(),
					Name = allStocksMaliplus.Cells[i, 2].Value?.ToString(),
					Distributor = allStocksMaliplus.Cells[i, 3].Value?.ToString(),
					IsVerified = false
				});
			}
			else
			{
				item.Quantity = qty;
			}
		}
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}

//using var objs = package.Workbook.Worksheets[0];

for (int i = unClarifiedItemCodes.Dimension.Start.Row; i < unClarifiedItemCodes.Dimension.End.Row; i++)
{
	try
	{
		if (CustomValidations.IsValidItemCode(unClarifiedItemCodes.Cells[i, 1].Value?.ToString() ?? ""))
		{
			unCleanItemCodes.Add(new()
			{
				Code = unClarifiedItemCodes.Cells[i, 1].Value?.ToString()?.Trim(),
				Name = unClarifiedItemCodes.Cells[i, 2].Value?.ToString(),
				Distributor = unClarifiedItemCodes.Cells[i, 3].Value?.ToString(),
				IsVerified = false
			});
		}
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}
Matcher.CheckCodes(ref allItemsCodes);
Matcher.CheckCodes(ref unCleanItemCodes);

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("WeEd", "weed"));


Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wed", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wedfsdf", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wasded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wsded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wsded", "weed"));

Console.WriteLine(Matcher.LaveteshinDistanceAlgorithm("Wded", "weed"));


