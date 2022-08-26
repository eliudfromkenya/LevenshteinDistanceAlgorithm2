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
                        col.Name = $"{col.Distributor}{col.Name}";
                }
                catch { }

                try
                {
                    var pattern = @"\(.*\)";
                    if (Regex.Match(col.Name ?? "", pattern).Success)
                    {
                        var value = Regex.Match(col.Name ?? "", pattern).Value;
                        col.Name = value[1..^1] + col.Name?.Replace(value,"");
                    }
                }
                catch { }

                try
                {
                    const string pattern = @"([\d]+ *[a-zA-Z]{1,5} *$)|([\d]* *x *[\d]+ *[a-zA-Z]{1,5} *$)|([\d]* *\* *[\d]+ *[a-zA-Z]{1,5} *$)|([\d]+ *[a-zA-Z]{1,5} *x *[\d]* *$)|([\d]+ *[a-zA-Z]{1,5} *\* *[\d]* *$)";

                    if (Regex.Match(col.Name ?? "", pattern).Success)
                    {
                        col.MeasureUnit = Regex.Match(col.Name ?? "", pattern).Value;
                        col.GroupName = col.Name?.Replace(col.MeasureUnit, "");
                    }

                    col.HarmonizedName = $@"{string.Join(" ", Regex.Matches(col.GroupName??"", "[0-9a-zA-Z]+")
                    .Select(v => v.Value?.ToUpper())
                    .Distinct().OrderBy(c => c))} {col.MeasureUnit}";
                }
                catch { }
            });
        }

        public static int LaveteshinDistanceAlgorithm(string s, string t)
        {
            s = s.ToUpper();
            t = t.ToUpper();

            int n = s.Length, m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            if (n == 0)
                return m;
            if (m == 0)
                return n;

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
            return d[n, m];
        }
    }
}
