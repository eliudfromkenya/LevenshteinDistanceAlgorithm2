// using MoreLinq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace LevenshteinDistanceAlgorithm
{
    static class Matcher
    {
        public static void CheckCodes(ref List<ItemCode> codes)
        {
            codes.ForEach(col =>
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(col.Distributor))
                        col.Name = $"{col.Distributor} {col.Name}";
                }
                catch (Exception ex) { Console.WriteLine("Error 1 " + ex.ToString()); }

                try
                {
                    var pattern = @"\(.*[A-Z]+.*\)";
                    if (Regex.Match(col.Name ?? "", pattern).Success)
                    {
                        var value = Regex.Match(col.Name ?? "", pattern).Value;
                        col.Name = $"{value[1..^1]} {col.Name?.Replace(value, "")}";
                    }
                }
                catch (Exception ex) { Console.WriteLine("Error 2A " + ex.ToString()); }
                col.Name = col.Name?.ToUpper();
                try
                {
                    var pattern = @"\"".*[A-Z]+.*\""";
                    if (Regex.Match(col.Name ?? "", pattern).Success)
                    {
                        var value = Regex.Match(col.Name ?? "", pattern).Value;
                        col.Name = $"{value[1..^1]} {col.Name?.Replace(value, "")}";
                    }
                }
                catch (Exception ex) { Console.WriteLine("Error 2B " + ex.ToString()); }
                try
                {
                    col.Name = CheckName(col.Name);
                }
                catch (Exception ex) { Console.WriteLine("Error 2B " + ex.ToString()); }
                try
                {
                    const string pattern = @"([\d]+ *[a-zA-Z]{1,5} *$)|([\d]* *x *[\d]+ *[a-zA-Z]{1,5} *$)|([\d]* *\* *[\d]+ *[a-zA-Z]{1,5} *$)|([\d]+ *[a-zA-Z]{1,5} *x *[\d]* *$)|([\d]+ *[a-zA-Z]{1,5} *\* *[\d]* *$)";

                    if (Regex.Match(col.Name ?? "", pattern).Success)
                        col.MeasureUnit = Regex.Match(col.Name ?? "", pattern).Value?.Replace("  "," ")?.Trim();
                    col.GroupName = col.Name?.Trim();
                    if (col.MeasureUnit?.Length > 1)
                        col.GroupName = col.Name?.Replace(col.MeasureUnit ?? "", "")?.Trim();
                }
                catch (Exception ex){ Console.WriteLine("Error 3 " + ex.ToString()); }
                try
                {
                    var replaces = new Dictionary<string, string[]>
                    {
                        { "KG",new[]{"KGS", "KILO", "KILOGRAM"} },
                        {"GM", new[]{ "GRAM","GS","G","GRM" } },
                        {"MM",new []{"MILLIMETER"} },
                        {"M", new[]{"METER", "MT","MTR"} },
                        {"ML",new []{"MILLITER","MLT"} },
                        {"L", new[]{"LITER", "LT","LTR"} }
                    };
                    //fert, cabb, D/ , S/Loaf, EGG P/BLACK BEAUTY,EGG PLANT B/B , EGG PLANT BEAUTY,L/MASH,L/STOCK,M/SALVE,M SALVE,H/PROS,H/PHOS,H/PHO, L/STOCK,D/LICK,S/LICK,HI PHOS,R/CREOLE,S/D LICK,M/MAKER,S  D/LICK,S/D LICK,S/LICK,D/MEAL,SUGARLOAF,D/LICK, D.LICK, D/LICK,M/MAKER,M. MAKER,TOMATOE,HIGHPHOS,H/PHOSPOROUS,HI-PHOS,S/BABY,H/PHOSPHOROUS,
                    //start with "
                    //replace name double space
                    // LTD. from distributor before matching
                    foreach (var item in replaces)
                    {
                        var objs = item.Value.SelectMany(n => new[] { n, $"{n}S" }).ToList();
                        var pattern = $"[a-zA-Z]+";
                        if (Regex.Match(col.MeasureUnit ?? "", pattern).Success)
                        {
                            var value = Regex.Match(col.MeasureUnit ?? "", pattern).Value;
                            var replacer = objs.FirstOrDefault(m => m == value.ToUpper());
                            if (!string.IsNullOrWhiteSpace(replacer) && !(value == item.Key))
                                col.MeasureUnit = col.MeasureUnit?.Replace(value, item.Key);
                            if ($"{item.Key}S" == value.ToUpper() && !(value == item.Key))
                                col.MeasureUnit = col.MeasureUnit?.Replace(value, item.Key);
                        }
                    }
                }
                catch (Exception ex) { Console.WriteLine("Error 4 " + ex.ToString()); }
                try
                {
                    col.HarmonizedName = $@"{string.Join(" ", Regex.Matches(col.GroupName ?? "", "[0-9a-zA-Z]+")
                   .Select(v => v.Value?.ToUpper()?.Trim())
                   .Distinct().OrderBy(c => c))} {col.MeasureUnit?.Trim()}".Replace(" ","");
                }
                catch (Exception ex) { Console.WriteLine("Error 5 " + ex.ToString()); }
                try
                {
                    col.HarmonizedGroupName = string.Join(" ", Regex.Matches(col.GroupName ?? "", "[0-9a-zA-Z]+")
                   .Select(v => v.Value?.ToUpper()?.Trim())
                   .Distinct().OrderBy(c => c)).Replace(" ", "");
                }
                catch (Exception ex) { Console.WriteLine("Error 5 " + ex.ToString()); }
            });
        }

        private static string? CheckName(string? name)
        {
            //D/LICK,S/LICK,R/CREOLE,S/D LICK,M/MAKER,S  D/LICK,S/D LICK,S/LICK,D/MEAL,SUGARLOAF,D/LICK, D.LICK, D/LICK,M/MAKER,M. MAKER,TOMATOE, S/BABY,H/PHOSPHOROUS
            //start with "
            //replace name double space
            // LTD. from distributor before matching
            if (name.Contains("CABB.") && !name.Contains("CABBAGE"))
                name = name.Replace("CABB.", "CABBAGE");
            if (name.Contains("CABB") && !name.Contains("CABBAGE"))
                name = name.Replace("CABB", "CABBAGE");
            if (name.Contains("FERT.") && !name.Contains("FERTILIZER"))
                name = name.Replace("FERT.", "FERTILIZER");
            if (name.Contains("FERT") && !name.Contains("FERTILIZER"))
                name = name.Replace("FERT", "FERTILIZER");
            if (name.Contains("D/") && !name.Contains("DAIRY "))
                name = name.Replace("D/", "DAIRY ");
            if (name.Contains("S/LOAF") && !name.Contains("SUGAR LOAF"))
                name = name.Replace("S/LOAF", "SUGAR LOAF");
            if (name.Contains("SUGARLOAF") && !name.Contains("SUGAR LOAF"))
                name = name.Replace("SUGARLOAF", "SUGAR LOAF");
            if (name.Contains("EGG P/BLACK BEAUTY"))
                name = name.Replace("EGG P/BLACK BEAUTY", "EGG PLANT BLACK BEAUTY");
            if (name.Contains("EGG PLANT B/B"))
                name = name.Replace("EGG PLANT B/B", "EGG PLANT BLACK BEAUTY");
            if (name.Contains("EGG PLANT BEAUTY"))
                name = name.Replace("EGG PLANT BEAUTY", "EGG PLANT BLACK BEAUTY");
            if (name.Contains("L/MASH") && !name.Contains("LAYERS MASH"))
                name = name.Replace("L/MASH", "LAYERS MASH");
            if (name.Contains("L/STOCK") && !name.Contains("LIVESTOCK"))
                name = name.Replace("L/STOCK", "LIVESTOCK");
            if (name.Contains("L/STOCK") && !name.Contains("LIVESTOCK"))
                name = name.Replace("L/STOCK", "LIVESTOCK");
            if (name.Contains("M/SALVE") && !name.Contains("MILKING SALVE"))
                name = name.Replace("M/SALVE", "MILKING SALVE");
            if (name.Contains("M SALVE") && !name.Contains("MILKING SALVE"))
                name = name.Replace("M SALVE", "MILKING SALVE");
            if (name.Contains("H/PROS") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("H/PROS", "HIGH PHOSPHORUS");
            if (name.Contains("H/PHO") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("H/PHO", "HIGH PHOSPHORUS");
            if (name.Contains("HI PHOS") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("HI PHOS", "HIGH PHOSPHORUS");
            if (name.Contains("HIGHPHOS") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("HIGHPHOS", "HIGH PHOSPHORUS");
            if (name.Contains("H/PHOSPOROUS") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("H/PHOSPOROUS", "HIGH PHOSPHORUS");
            if (name.Contains("HI-PHOS") && !name.Contains("HIGH PHOSPHORUS"))
                name = name.Replace("HI-PHOS", "HIGH PHOSPHORUS");
        }

        public static short LaveteshinDistanceAlgorithm(ItemCode code, ItemCode code2)
        {
            short level = LaveteshinDistanceAlgorithmBody(code.HarmonizedName ?? "", code2.HarmonizedName ?? "");
            if (string.IsNullOrWhiteSpace(code.MeasureUnit)
                || string.IsNullOrWhiteSpace(code.MeasureUnit))
                level += 2;
            else if (code.MeasureUnit != code2.MeasureUnit)
                level += 5;
            return level;
        }
        public static short LaveteshinDistanceAlgorithmBody(string s, string t)
        {
            s = s.ToUpper();
            t = t.ToUpper();

            int n = s.Length, m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            if (n == 0)
                return (short)m;
            if (m == 0)
                return (short)n;

            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 0; j <= m; d[0, j] = j++) ;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }
            return (short)(d[n, m] + 1);
        }

        internal static List<ItemCodeMatch> MatchItemCode(List<ItemCode> nyahururuItemCodes, List<ItemCode> allItemsCodes, List<ItemCode> unCleanItemCodes)
        {
            List<ItemCodeMatch> itemCodeMatches = new();
            nyahururuItemCodes.ForEach(tt =>
            {
                ItemCodeMatch match = new()
                {
                    OriginalCode = tt
                };
                itemCodeMatches.Add(match);

                var obj = allItemsCodes.FirstOrDefault(m => m.HarmonizedName == tt.HarmonizedName);

                if(obj !=null)
                {
                    match.MatchedCode=obj;
                    match.MatchStrength = 0;
                    return;
                }

                obj = allItemsCodes.MinBy(v => LaveteshinDistanceAlgorithm(v, tt));
                match.MatchedCode = obj;
                match.MatchStrength = (short)LaveteshinDistanceAlgorithm(obj, tt);
                if (match.MatchStrength < 4)
                    return;

                var obj2 = unCleanItemCodes.MinBy(v => LaveteshinDistanceAlgorithm(v, tt));
                var strength = (short)LaveteshinDistanceAlgorithm(obj2, tt);
                if(match.MatchStrength < strength + 2)
                {
                    match.MatchedCode = obj2;
                    match.MatchStrength = strength;
                }
            });
            return itemCodeMatches;
        }
    }
}
