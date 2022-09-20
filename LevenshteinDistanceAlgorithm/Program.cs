using LevenshteinDistanceAlgorithm;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Linq;


	ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
	string mainFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Excel Working Files");

	using var allStocksMaliplusFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "item_master.xlsx")));
	//using var nyahururuDifferenceFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "item difference Nyahururu Branch.xlsx")));
    using var nyahururuDifferenceFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "NYAHURURU ITEMS SORT.xlsx")));
	using var oldItemCodesFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "matching rows.xlsx")));
	using var unClarifiedItemCodesFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "not matching.xlsx")));

	using var allStocksMaliplus = allStocksMaliplusFile.Workbook.Worksheets[0];
	using var nyahururuDifference = nyahururuDifferenceFile.Workbook.Worksheets["Sheet2"];
	using var oldItemCodes = oldItemCodesFile.Workbook.Worksheets[0];
	using var unClarifiedItemCodes = unClarifiedItemCodesFile.Workbook.Worksheets[0];

	List<ItemCode> allItemsCodes = new(), /*unCleanItemCodes = new(), */nyahururuItemCodes = new();


//var groupData = JsonConvert.DeserializeObject<Dictionary<string, string>>(JsonData.Groups);

for (int i = oldItemCodes.Dimension.Start.Row; i < oldItemCodes.Dimension.End.Row; i++)
{
	try
	{
		var itemCode = oldItemCodes.Cells[i, 1].Value?.ToString()?.Trim() ?? "";

		if (allItemsCodes.Any(n => n.Code == itemCode))
			continue;

		if (CustomValidations.IsValidItemCode(itemCode))
		{
			allItemsCodes.Add(new()
			{
				Code = itemCode,
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
			//try
			//{
			//	if (decimal.TryParse(allStocksMaliplus.Cells[i, 4].Value?.ToString() ?? "", out decimal num))
			//		qty = num;
			//}
			//catch { }
			//try
			//{
			//	if (qty == 0 && decimal.TryParse(allStocksMaliplus.Cells[i, 3].Value?.ToString() ?? "", out decimal num))
			//		qty = num;
			//}
			//catch { }

			if (CustomValidations.IsValidItemCode(itemCode))
			{
				var item = allItemsCodes.FirstOrDefault(x => x.Code == itemCode);
				if (item == null)
				{
					allItemsCodes.Add(new()
					{
						Quantity = qty,
						Code = allStocksMaliplus.Cells[i, 1].Value?.ToString()?.Trim(),
						Name = allStocksMaliplus.Cells[i, 3].Value?.ToString()?.Trim().Replace("  ", " "),
						//Distributor = allStocksMaliplus.Cells[i, 13].Value?.ToString()?.Trim(),
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


//for (int i = nyahururuDifference.Dimension.Start.Row; i < nyahururuDifference.Dimension.End.Row; i++)
//{
//	try
//	{
//		if (CustomValidations.IsValidItemCode(nyahururuDifference.Cells[i, 3].Value?.ToString()?.Trim() ?? ""))
//		{
//			nyahururuItemCodes.Add(new()
//			{
//				Code = nyahururuDifference.Cells[i, 3].Value?.ToString()?.Trim(),
//				Name = nyahururuDifference.Cells[i, 4].Value?.ToString()?.Trim().Replace("  ", " "),
//				Quantity = decimal.TryParse(nyahururuDifference.Cells[i, 5].Value?.ToString()?.Trim(), out decimal mx) ? mx : 0,
//				IsVerified = false
//			});
//		}
//	}
//	catch (Exception ex)
//	{
//		Console.BackgroundColor = ConsoleColor.DarkRed;
//		Console.WriteLine(ex);
//		Console.BackgroundColor = ConsoleColor.Black;
//	}
//}


for (int i = nyahururuDifference.Dimension.Start.Row; i < nyahururuDifference.Dimension.End.Row; i++)
{
	try
	{
		var cells = nyahururuDifference.Cells;
		var itemCode = cells[i, 1].Value?.ToString()?.Trim() ?? "";
		if (int.TryParse(itemCode, out int itmNumber))
			itemCode = itmNumber.ToString("000000");

		if (CustomValidations.IsValidItemCode(itemCode))
		{
			
			nyahururuItemCodes.Add(new()
			{
				Code = itemCode, 
				Narration = String.Join(",", Enumerable.Range(3,4).Select(v => cells[i, v].Value?.ToString()?.Trim())),
				Name = cells[i, 2].Value?.ToString()?.Trim()
				.Replace("  ", " "),
				Quantity = decimal.TryParse(cells[i, 6].Value?.ToString()?.Trim(), out decimal mx) ? mx : 0,
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

void RefreshCodes()
{

	using var con = new ConnectionObject().MaliplusConnection;
	con.Open();
	using var cmd = con.CreateCommand();
	cmd.CommandText = "SELECT ITEM_CODE, ITEM_NAME, BUY_PRICE, SALE_PRICE FROM ITEM_MASTER";

	using var reader = cmd.ExecuteReader();
	while (reader.Read())
	{
		try
		{
			var itemCode = reader.GetString(0);
			if (CustomValidations.IsValidItemCode(itemCode))
			{
				var item = allItemsCodes.FirstOrDefault(x => x.Code == itemCode);
				_ = decimal.TryParse(reader.GetString(2), out decimal res);
				var qty = res;
				if (item == null)
				{
					allItemsCodes.Add(new()
					{
						Quantity = qty,
						Code = itemCode,
						Name = reader.GetString(1),
						IsVerified = true
					});
				}
				else
				{
					item.Quantity = qty;
					item.Name = reader.GetString(1);
					item.IsVerified = true;
				}
			}
			var grops = JsonData.Groups;
			foreach (var item in allItemsCodes)
			{
				try
				{
					item.ItemGroup = grops.FirstOrDefault(c => c.id == item?.Code?[..2]).name;
				}
				catch { }
			}
            Console.WriteLine("Done");
		}
		catch (Exception ex)
		{
			Console.BackgroundColor = ConsoleColor.DarkRed;
			Console.WriteLine(ex);
			Console.BackgroundColor = ConsoleColor.Black;
		}
	}
}

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


void GenerateFinalExcel()
{
	try
	{
		var excelPackage = allStocksMaliplusFile;
		excelPackage.Workbook.Properties.Author = "KFA / Primesoft Implimentation Team";
		excelPackage.Workbook.Properties.Title = "Zanas to Maliplus Item Master Matching Report";
		excelPackage.Workbook.Properties.Subject = "Automated Data Matching";
		excelPackage.Workbook.Properties.Created = DateTime.Now;
		excelPackage.Workbook.Properties.Company = "KFA LTD";

		var branch = "Finalized Item Master";
		var title = $"{branch} Item Codes Match Information";

		var service = new MsExcelReportService();
		//service.GenerateMatchReport(excelPackage, data, mainFolder, branch);
		service.GenerateGroupsCodeReport(excelPackage, allItemsCodes, mainFolder, branch);
		service.GenerateDulplicatedItemCodes2Report(excelPackage, allItemsCodes, mainFolder, branch);
		//service.GenerateDulplicatedItemCodesOnlyReport(excelPackage, allItemsCodes, mainFolder, branch);

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

//RefreshCodes();
try
{
	do
	{
		const string value = @"Please Select value
   1. Search Free Item Code Forward.
   2. Search Free Item Code Backward.
   3. Search item by name.
   4. Process Excel.
   5. Update Item Code. 
   6. Generate Final Master.
   7. Refresh Item Codes.
   Q. Quit.
";
		Console.WriteLine(value);

		//var objs = from cell in nyahururuDifference.Cells["C:C"]// a:a is the column a
		//		   where cell?.Value?.ToString()?.Equals(x)  // x is the input userid
		//		   select sheet.Cells[cell.Start.Row, 2]; // 2 is column b, Email Address
		var objs = (from cell in nyahururuDifference.Cells["C:C"]
					where CustomValidations
					      .IsValidItemCode(cell?.Value?.ToString() ?? "")
					select new { Data = cell.Value.ToString(), cell.Start.Row }
					)?.OrderBy(c => c.Row)?.Select(v => v.Data)?.ToList();

		var key = Console.ReadKey();
		if (key.Key == ConsoleKey.D1)
			ItemChecker.SearchItemForward(allItemsCodes);
		else if (key.Key == ConsoleKey.D2)
			ItemChecker.SearchItemBackward(allItemsCodes);
		else if (key.Key == ConsoleKey.D3)
			ItemChecker.SearchItemByName(allItemsCodes);
		else if (key.Key == ConsoleKey.D4)
			GenerateBranchExcels();
		else if (key.Key == ConsoleKey.D5)
			ItemChecker.UpdateItemByName(allItemsCodes, nyahururuItemCodes, objs ?? new List<string>(), mainFolder);
		else if (key.Key == ConsoleKey.D6)
			GenerateFinalExcel();
		else if (key.Key == ConsoleKey.D7)
			RefreshCodes();
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