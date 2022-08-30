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

	List<ItemCode> allItemsCodes = new(), /*unCleanItemCodes = new(), */nyahururuItemCodes = new();
	for (int i = oldItemCodes.Dimension.Start.Row; i < oldItemCodes.Dimension.End.Row; i++)
	{
		try
		{
			if (CustomValidations.IsValidItemCode(oldItemCodes.Cells[i, 1].Value?.ToString()?.Trim() ?? ""))
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
				if (decimal.TryParse(allStocksMaliplus.Cells[i, 4].Value?.ToString() ?? "", out decimal num))
					qty = num;
			}
			catch { }
			try
			{
				if (qty == 0 && decimal.TryParse(allStocksMaliplus.Cells[i, 3].Value?.ToString() ?? "", out decimal num))
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
						Name = allStocksMaliplus.Cells[i, 2].Value?.ToString()?.Trim().Replace("  ", " "),
						Distributor = allStocksMaliplus.Cells[i, 3].Value?.ToString()?.Trim(),
						IsVerified = true
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
			var itemCode = unClarifiedItemCodes.Cells[i, 1].Value?.ToString()?.Trim() ?? "";

			if (CustomValidations.IsValidItemCode(itemCode))
			{
				var item = allItemsCodes.FirstOrDefault(x => x.Code == itemCode);
				if (item == null)
				{
					ItemCode obj = new()
					{
						Quantity = 0,
						Code = unClarifiedItemCodes.Cells[i, 1].Value?.ToString()?.Trim(),
						Name = unClarifiedItemCodes.Cells[i, 2].Value?.ToString()?.Trim().Replace("  ", " "),
						Distributor = unClarifiedItemCodes.Cells[i, 3].Value?.ToString()?.Trim(),
						IsVerified = false
					};
					if (!string.IsNullOrWhiteSpace(obj.Name))
						allItemsCodes.Add(obj);
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
    //Matcher.CheckCodes(ref unCleanItemCodes);
    Matcher.CheckCodes(ref nyahururuItemCodes);

    var data = Matcher.MatchItemCode(nyahururuItemCodes, allItemsCodes)
        .OrderByDescending(v => v.OriginalCode?.IsVerified)
        .ThenBy(v => v.OriginalCode?.Code)
        .ToList();

void GenerateBranchExcels()
{
	try
	{
		using var excelPackage = new ExcelPackage();
		excelPackage.Workbook.Properties.Author = "KFA / Primesoft Implimentation Team";
		excelPackage.Workbook.Properties.Title = "Zanas to Maliplus Item Master Matching Report";
		excelPackage.Workbook.Properties.Subject = "Automated Data Matching";
		excelPackage.Workbook.Properties.Created = DateTime.Now;
		excelPackage.Workbook.Properties.Company = "KFA LTD";

		var branch = "Nyahururu Branch";
		var title = $"{branch} Item Codes Match Information";

		var service = new MsExcelReportService();
		service.GenerateMatchReport(excelPackage, data, mainFolder, branch);
		service.GenerateGroupsCodeReport(excelPackage, allItemsCodes, mainFolder, branch);
		service.GenerateDulplicatedItemCodesReport(excelPackage, allItemsCodes, mainFolder, branch);
		service.GenerateDulplicatedItemCodesOnlyReport(excelPackage, allItemsCodes, mainFolder, branch);

		Console.WriteLine("Done");
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
	//Console.WriteLine("Done");
}


try
{
	do
	{
		const string value = @"Please Select value
   1. Search Free Item Code Forward.
   2. Search Free Item Code Backward.
   3. Search item by name.
   4. Process Excel.
   Q. Quit.
";
		Console.WriteLine(value);

        var key = Console.ReadKey();
		if (key.Key == ConsoleKey.D1)
			ItemChecker.SearchItemForward(allItemsCodes);
		else if (key.Key == ConsoleKey.D2)
			ItemChecker.SearchItemBackward(allItemsCodes);
        else if (key.Key == ConsoleKey.D3)
            ItemChecker.SearchItemByName(allItemsCodes);
        else if (key.Key == ConsoleKey.D4)
			GenerateBranchExcels();
        else if (key.Key == ConsoleKey.Q)
			Environment.Exit(0);
    } while (true);	
}
catch (Exception ex)
{
    Console.BackgroundColor = ConsoleColor.DarkRed;
    Console.WriteLine(ex);
    Console.BackgroundColor = ConsoleColor.Black;
}