using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class NestingList
{
    public static void Gen(List<Gen> gens)
    {
        var items = new List<string[]>
        {
            new string[]{},
            new []{"", "Перечень карт раскроя"},
            new []
            {
                "№ п/п",
                "№ карты",
                "Типоразмер",
                "Марка матер.",
                "Кратность",
                "Коэф. раскроя"
            }
        };

        for (var i = 0; i < gens.Count; i++)
        {
            var elem = gens[i];

            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.NestName,
                $"{elem.RawThickness:G}x{elem.RawLength:G}x{elem.RawWidth:G}",
                elem.Quality,
                (elem.QuantityNormal+elem.QuantityMirrored).ToString(CultureInfo.InvariantCulture),
                elem.NestingPercent.ToString("F3", CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень карт раскроя.xlsx", items);
    }
}