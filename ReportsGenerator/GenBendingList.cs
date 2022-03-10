using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class BendingList
{
    public static void Gen(Dictionary<string,Wcog> parts)
    {
        var items = new List<string[]>
        {
            new string[]{},
            new []{"", "Перечень гнутых деталей"},
            new[]
            {
                "№ п/п",
                "№ чертежа",
                "№ дет.",
                "Наименование",
                "Марка мат.",
                "Толщина",
                "Кол-во",
                "Масса"
            }
        };

        var list = parts.Keys.ToList();
        list.Sort((x, y) =>
        {
            if(int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x, y, StringComparison.Ordinal);
        });

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = parts[key];

            items.Add(new[]
            {
                (i + 1).ToString(),
                Settings.Default.Drawing,
                elem.PosNo,
                elem.GetName(),
                elem.Quality,
                elem.GetThickness().ToString(CultureInfo.InvariantCulture),
                elem.Quantity.ToString(CultureInfo.InvariantCulture),
                elem.Weight.ToString(CultureInfo.InvariantCulture)
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень гнутых деталей.xlsx", items);
    }
}