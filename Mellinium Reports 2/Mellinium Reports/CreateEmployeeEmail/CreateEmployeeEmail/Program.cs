// See https://aka.ms/new-console-template for more information
using MoreLinq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text.RegularExpressions;

Console.WriteLine("Hello, World!");

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//string mainFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Mellinium Reports");
var mainFolder = @"C:\Users\Eliud\Documents\GitHub\LevenshteinDistanceAlgorithm2\Mellinium Reports 2\Mellinium Reports";
using var unionisableFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "middle.xlsx")));

using var middleMgtFile = new ExcelPackage(new FileInfo(Path.Combine(mainFolder, "unionss_xls.xlsx")));


using var unionisableSheet = unionisableFile.Workbook.Worksheets[0];
using var middleMgtSheet = middleMgtFile.Workbook.Worksheets[0];

Dictionary<string,string?> departments = new();
List<(string? number, string? name, string? department, string? type, string? email)> staff = new();

string GetEmail(string name)
{
	var names = name?.ToLower()?.Split(' ').Where(v => v?.Length > 2).ToList();
	if(names?.Count < 2)
        names = name?.ToLower()?.Split(' ').ToList();
	var dd = names?[1][0] + names?.First();
	
	var email = dd + "@kenyafarmersassociation.co.ke";
    if(staff.Any(v => v.email == email))
    {
       email = $"{names?.First()}@kenyafarmersassociation.co.ke";
        if (staff.Any(v => v.email == email))
        {
            email = $"{names?.First()}{string.Join("",names?.Skip(1).Select(n => n.First()))}@kenyafarmersassociation.co.ke";
            if (staff.Any(v => v.email == email))
            {
                email = $"{string.Join("", names)}@kenyafarmersassociation.co.ke";
            }
        }
    }
    return email;
}


for (int i = middleMgtSheet.Dimension.Start.Row; i < middleMgtSheet.Dimension.End.Row; i++)
{
	try
	{
		var code = middleMgtSheet.Cells[i, 1].Value?.ToString()?.Trim() ?? "";
		if(Regex.IsMatch(code, "^[0-9]{4} *[0-9]{4}"))
		{
			var val = Regex.Match(code, "^[0-9]{4} *[0-9]{4}").Value?.Replace("  ", " ")?.Split(' ');

			staff.Add((val?.LastOrDefault(), middleMgtSheet.Cells[i, 2].Value?.ToString(), val?.FirstOrDefault(), "Management", GetEmail(middleMgtSheet.Cells[i, 2].Value?.ToString())));
		}
		else if(Regex.IsMatch(code, "^[0-9]{4} *[A-Z]"))
		{
            var val = Regex.Match(code, "^[0-9]{4} *[A-Z]").Value?.Replace("  ", " ")?.Split(' ');
			departments[val?.FirstOrDefault() ?? ""] = middleMgtSheet.Cells[i, 1].Value?.ToString()?[5..].Trim();
		}
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}




for (int i = unionisableSheet.Dimension.Start.Row; i < unionisableSheet.Dimension.End.Row; i++)
{
	try
	{
		var code = unionisableSheet.Cells[i, 1].Value?.ToString()?.Trim() ?? "";
		if (Regex.IsMatch(code, "^[0-9]{4} *[0-9]{4}"))
		{
			var val = Regex.Match(code, "^[0-9]{4} *[0-9]{4}").Value?.Replace("  ", " ")?.Split(' ');

			staff.Add((val?.LastOrDefault(), unionisableSheet.Cells[i, 2].Value?.ToString(), val?.FirstOrDefault(), "Union",GetEmail(unionisableSheet.Cells[i, 2].Value?.ToString())));
		}
		else if (Regex.IsMatch(code, "^[0-9]{4} *[A-Z]"))
		{
			var val = Regex.Match(code, "^[0-9]{4} *[A-Z]").Value?.Replace("  ", " ")?.Split(' ');
			departments[val?.FirstOrDefault() ?? ""] = unionisableSheet.Cells[i, 1].Value?.ToString()?[5..].Trim();
		}
	}
	catch (Exception ex)
	{
		Console.BackgroundColor = ConsoleColor.DarkRed;
		Console.WriteLine(ex);
		Console.BackgroundColor = ConsoleColor.Black;
	}
}

var mm = staff.ToList();
var dd = departments.ToList();

int xCount = 1;
string AddMsExcelReport(ExcelPackage excelPackage, string title)
{
    var currentRow = 3;
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

    var range = sheet.Cells[currentRow, 1, currentRow, 6];
    range.Style.Font.Color.SetColor(Color.RebeccaPurple);
    range.Style.Font.UnderLine = true;
    range.Style.Font.Size = 14;
    range.Merge = true;
    range.Value = $"Staff Details";

    currentRow++;

    sheet.Cells[currentRow, 1].Value = "Num";
    sheet.Cells[currentRow, 2].Value = "Staff Number";
    sheet.Cells[currentRow, 3].Value = "Name";
    sheet.Cells[currentRow, 4].Value = "Category";
    sheet.Cells[currentRow, 5].Value = "Phone Number";
    sheet.Cells[currentRow, 6].Value = "Email Address";

    sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.Color.SetColor(Color.RebeccaPurple);
    sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.UnderLine = true;
    currentRow++;


    staff.GroupBy(c => c.department).OrderBy(m => m.Key).ForEach(m =>
    {
        currentRow += 2;
        var dept = departments.FirstOrDefault(c => c.Key == m.Key);
        sheet.Cells[currentRow, 1, currentRow, 6].Value = $"{m.Key} - {dept.Value}";
        sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.Color.SetColor(Color.DarkOrchid);
        sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.UnderLine = true;
        sheet.Cells[currentRow, 1, currentRow, 6].Style.Font.Bold = true;
        sheet.Cells[currentRow, 1, currentRow, 6].Merge = true;
        currentRow++;

        var stf = m.OrderBy(mm => mm.number).ToList();
        var isMultiple = stf.Select(v => v.type).Distinct().Count() > 1;

        m.GroupBy(m => m.type).ToList().ForEach(cc =>
        {
            cc.OrderBy(mm => mm.number).ToList().ForEach(vv =>
            {
                sheet.Cells[currentRow, 1].Value = xCount++;
                sheet.Cells[currentRow, 2].Value = vv.number;
                sheet.Cells[currentRow, 3].Value = vv.name;
                sheet.Cells[currentRow, 4].Value = vv.type;
                sheet.Cells[currentRow, 5].Value = null;
                sheet.Cells[currentRow, 6].Value = vv.email;

                var range = sheet.Cells[currentRow, 1, currentRow, 6];
                range.Style.Font.Color.SetColor(vv.type == "Management"?  Color.Black : Color.DarkGreen);     

                if (isMultiple && vv.type != "Management")
                {
                    isMultiple = false;
                    range.Style.Border.Top.Style = ExcelBorderStyle.MediumDashDotDot;
                }
                 currentRow++;
            });
        });


    });    

    var file = Path.Combine(mainFolder, $"userlist 345.xlsx");
    //excelPackage.SaveAs(new FileInfo(file), "Zanas2022");
    excelPackage.SaveAs(new FileInfo(file));
    return file;
}

try
{
	using var excelPackage = new ExcelPackage();
	excelPackage.Workbook.Properties.Author = "KFA / Mellinium Implimentation Team";
	excelPackage.Workbook.Properties.Title = "Users Stafflist Report";
	excelPackage.Workbook.Properties.Subject = "System Users";
	excelPackage.Workbook.Properties.Created = DateTime.Now;
	excelPackage.Workbook.Properties.Company = "KFA LTD";

	var title = $"KFA Staff List";

	AddMsExcelReport(excelPackage, title);

	Console.WriteLine("Done");
}
catch (Exception ex)
{
	Console.BackgroundColor = ConsoleColor.DarkRed;
	Console.WriteLine(ex);
	Console.BackgroundColor = ConsoleColor.Black;
}