using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class NestingList
{
    public static void Gen(List<Gen> gens)
    {
        var items = new List<string[]>
        {
            new []
            {
                "№ п/п",
                "Карта",
                "Толщина",
                "Марка",
                "Габариты",
                "Кол-во деталей",
                "Коэф. раскроя",
                "Длина реза",
                "Длина ХХ",
                "Кол-во пробивок",
                "Масса деталей",
                "Масса отхода",
                "Дата"
            }
        };

        for (var i = 0; i < gens.Count; i++)
        {
            var elem = gens[i];

            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.NestName,
                elem.RawThickness.ToString("F1"),
                elem.Quality,
                $"{elem.RawLength:G}x{elem.RawWidth:G}",
                elem.NoOfParts.ToString(),
                elem.NestingPercent.ToString(CultureInfo.InvariantCulture),
                elem.TotalBurning.ToString(CultureInfo.InvariantCulture),
                elem.TotalIdle.ToString(CultureInfo.InvariantCulture),
                elem.NoOfBurningStarts.ToString(),
                elem.PartsWeight.ToString(CultureInfo.InvariantCulture),
                elem.RemnantWeight.ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень карт раскроя.xlsx", items);
    }
}