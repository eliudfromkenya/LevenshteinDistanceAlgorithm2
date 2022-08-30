using LevenshteinDistanceAlgorithm;
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
                            var line = $"{xCount++}. {word.Code} - {word.Name}";
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
}