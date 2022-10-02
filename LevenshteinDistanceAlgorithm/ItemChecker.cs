using LevenshteinDistanceAlgorithm;
using OfficeOpenXml;
using System.Text;

internal class ItemChecker
{
    internal static void SearchItemBackward(List<LevenshteinDistanceAlgorithm.ItemCode> allItemsCodes)
    {
        do
        {
            try
            {
                const string value = @"Please enter item code to check next free or Q to quit.";
                Console.WriteLine(value);
                var ans = Console.ReadLine();

                if (ans?.ToUpper()?.StartsWith("Q") ?? false) return;

                if (int.TryParse(ans, out int obj))
                {
                    for (int i = obj; i < 460000; i++)
                    {
                        var code = i.ToString("000000");
                        if (!allItemsCodes.Any(n => n.Code == code))
                        {
                            Console.WriteLine(code);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(ex);
                Console.BackgroundColor = ConsoleColor.Black;
            }
        }while (true);
    }

    internal static void SearchItemForward(List<LevenshteinDistanceAlgorithm.ItemCode> allItemsCodes)
    {
        do
        {
            try
            {
                const string value = @"Please enter item code to check next free or Q to quit.";
                Console.WriteLine(value);
                var ans = Console.ReadLine();

                if (ans?.ToUpper()?.StartsWith("Q") ?? false) return;

                if (int.TryParse(ans, out int obj))
                {
                    for (int i = obj - 1; i >= 10001; i--)
                    {
                        var code = i.ToString("000000");
                        if (!allItemsCodes.Any(n => n.Code == code))
                        {
                            Console.WriteLine(code);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(ex);
                Console.BackgroundColor = ConsoleColor.Black;
            }
        } while (true);
    }

    internal static void SearchItemByName(List<LevenshteinDistanceAlgorithm.ItemCode> allItemsCodes)
    {
        do
        {
            try
            {
                const string value = @"Please enter item name to search or Q to quit.";
                Console.WriteLine(value);
                var ans = Console.ReadLine();
                ans = ans?.ToUpper();

                if (ans?.StartsWith("Q") ?? false) return;

                int xCount = 1;
                var body = new StringBuilder();
                allItemsCodes
                    .OrderBy(n => Matcher.LaveteshinDistanceAlgorithmBody(n.HarmonizedName ?? "", ans ?? ""))
                    .Take(20).ToList().ForEach(word =>
                    {
                        try
                        {
                            var line = $"{xCount++}. {word.Code} - {word.Name} => ({word.ItemGroup})";
                            body.AppendLine(line);
                        }
                        catch { }
                    });
                Console.WriteLine(body.ToString());
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(ex);
                Console.BackgroundColor = ConsoleColor.Black;
            }
        } while (true);
    }

    internal static void UpdateItemByName(List<ItemCode> allItemsCodes, List<ItemCode> nyahururuItems, List<string> nyahururuItemCodes, string mainFolder)
    {
        int currentIndex = 0;
        using var excelPackage = new ExcelPackage();
        var path = Path.Combine(mainFolder, "BranchReplacements.xlsx");
        using var sheet = excelPackage.Workbook.Worksheets.Add("Branch Code Replacements");
        var matches = Matcher.MatchItemCode(nyahururuItems, allItemsCodes);
        do
        {
            try
            {
                var code = nyahururuItemCodes[currentIndex];
                var xCode = matches.FirstOrDefault(x => x.OriginalCode?.Code == code?.Trim());

                if (xCode == null)
                {
                    currentIndex++;
                    continue;
                }

                if (xCode.MatchStrength == 0)
                {
                    try
                    {
                        sheet.Cells[currentIndex + 1, 1].Value = xCode?.OriginalCode?.Code;
                        sheet.Cells[currentIndex + 1, 2].Value = xCode?.OriginalCode?.Name;
                        sheet.Cells[currentIndex + 1, 3].Value = xCode?.MatchedCode?.Code;
                        sheet.Cells[currentIndex + 1, 4].Value = xCode?.MatchedCode?.Name;
                        sheet.Cells[currentIndex + 1, 5].Value = xCode?.MatchStrength;
                        excelPackage.SaveAs(path);
                    }
                    catch { }

                      currentIndex++;
                    continue;
                }

                string value = @$"Checking item code: {xCode?.OriginalCode?.Code} - {xCode?.OriginalCode?.Name}";
                Console.WriteLine(value);


                var ans = Console.ReadLine();
                ans = ans?.ToUpper();

                if (ans?.StartsWith("Q") ?? false) return;
                if (ans?.StartsWith("P") ?? false)
                {
                    if (currentIndex == 0)
                        throw new Exception("Can't move back because that is the first record already");
                    else
                        currentIndex--;
                    continue;
                }
                if (ans?.StartsWith("N") ?? false)
                {
                    if (nyahururuItemCodes.Count >= currentIndex)
                        throw new Exception("Can't move forward because that is the last record already");
                    else
                        currentIndex++;
                    continue;
                }

                var matchedCodes = Matcher.MatchCodes(xCode?.OriginalCode, allItemsCodes, 20);

                int xCount = 1;
                var body = new StringBuilder();
                matchedCodes
                    .OrderBy(n => n.Measure)
                    .Take(15).ToList().ForEach(word =>
                    {
                        try
                        {
                            var line = $"{xCount++}. {word.Code.Code} - {word.Code.Name}";
                            body.AppendLine(line);
                        }
                        catch { }
                    });

                body.AppendLine($@"
Select an item code from the list e.g 1, 2 or 3 or enter a valid item code. ({currentIndex} / {nyahururuItemCodes.Count})
");
                
                Console.WriteLine(body.ToString());
                var newItemCode = Console.ReadLine();
                var newItemName = xCode.OriginalCode.Name;
                short match = -1;
                if (int.TryParse(newItemCode, out int mxs) && mxs < 100)
                {
                    newItemCode = matchedCodes[mxs - 1].Code?.Code;
                    newItemName = matchedCodes[mxs - 1].Code?.Name;
                    match=matchedCodes[mxs - 1].Measure;
                }
                else
                {
                    var objFound = allItemsCodes.FirstOrDefault(x => x.Code == newItemCode);
                    if (objFound != null)
                    {
                        newItemCode = objFound?.Code;
                        newItemName = objFound?.Name;
                        match = 1000;
                    }
                }              

                for (int i = 0; i < 10; i++)
                {
                    if (CustomValidations.IsValidItemCode(newItemCode ?? ""))
                    {
                        try
                        {
                            sheet.Cells[currentIndex + 1, 1].Value = xCode?.OriginalCode?.Code;
                            sheet.Cells[currentIndex + 1, 2].Value = xCode?.OriginalCode?.Name;
                            sheet.Cells[currentIndex + 1, 3].Value = newItemCode;
                            sheet.Cells[currentIndex + 1, 4].Value = newItemName;
                            sheet.Cells[currentIndex + 1, 5].Value = match;
                          
                            excelPackage.SaveAs(path);
                        }
                        catch { }
                         currentIndex++;
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Please re-enter the item code.");
                        newItemCode = Console.ReadLine();
                    }
                }
                //excelPackage.Save(path);
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(ex);
                Console.BackgroundColor = ConsoleColor.Black;
            }
        } while (true);
        try
        {
            var objs = from cell in sheet.Cells["A:A"]
                       where !allItemsCodes.Any(m => m.Code == cell?.Value?.ToString())
                       select cell.Start.Row;

            if (objs.Any())
            {
                foreach (var row in objs)
                {
                    var cell = sheet.Cells[$"B{row}"];
                    cell.Style.Fill.SetBackground(System.Drawing.Color.DarkOrange);
                }
            }
        }
        catch (Exception)
        {
        }


        excelPackage.Save(path);
    }
}