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

List<ItemCode> allItemsCodes = new(), unCleanItemCodes = new(), nyahururuItemCodes = new();
for (int i = oldItemCodes.Dimension.Start.Row; i < oldItemCodes.Dimension.End.Row; i++)
{
	try
	{
        if(CustomValidations.IsValidItemCode(oldItemCodes.Cells[i, 1].Value?.ToString()?.Trim() ?? ""))
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
			allStocksMaliplus.Cells[i, 1].Value?.ToString()?.Trim() ?? "";

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
					Name = allStocksMaliplus.Cells[i, 2].Value?.ToString()?.Trim().Replace("  "," "),
                    Distributor = allStocksMaliplus.Cells[i, 3].Value?.ToString()?.Trim(),
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

// using var objs = package.Workbook.Worksheets[0];

for (int i = unClarifiedItemCodes.Dimension.Start.Row; i < unClarifiedItemCodes.Dimension.End.Row; i++)
{
	try
	{
		if (CustomValidations.IsValidItemCode(unClarifiedItemCodes.Cells[i, 1].Value?.ToString()?.Trim() ?? ""))
		{
			ItemCode obj = new()
			{
				Code = unClarifiedItemCodes.Cells[i, 1].Value?.ToString()?.Trim(),
				Name = unClarifiedItemCodes.Cells[i, 2].Value?.ToString()?.Trim().Replace("  ", " "),
				Distributor = unClarifiedItemCodes.Cells[i, 3].Value?.ToString()?.Trim(),
				IsVerified = false
			};

			if (!string.IsNullOrWhiteSpace(obj.Name))
				unCleanItemCodes.Add(obj);
		}
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}


for (int i = nyahururuDifference.Dimension.Start.Row; i < nyahururuDifference.Dimension.End.Row; i++)
{
    try
    {
        if (CustomValidations.IsValidItemCode(nyahururuDifference.Cells[i, 3].Value?.ToString()?.Trim() ?? ""))
        {
			nyahururuItemCodes.Add(new()
			{
				Code = nyahururuDifference.Cells[i, 3].Value?.ToString()?.Trim(),
				Name = nyahururuDifference.Cells[i, 4].Value?.ToString()?.Trim().Replace("  ", " "),
				Quantity = decimal.TryParse(nyahururuDifference.Cells[i, 5].Value?.ToString()?.Trim(), out decimal mx) ? mx : 0,
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
Matcher.CheckCodes(ref nyahururuItemCodes);

var data = Matcher.MatchItemCode(nyahururuItemCodes, allItemsCodes, unCleanItemCodes).OrderBy(v => v.OriginalCode).ToList();
new MsExcelReportService().GenerateMatchReport(data, mainFolder, "Nyahururu Branch");
