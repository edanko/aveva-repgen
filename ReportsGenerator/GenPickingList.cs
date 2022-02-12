using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class PickingList
{
    public static void Gen(Dictionary<string, Wcog> wcog)
    {
        var items = new List<string[]>
        {
            new []
            {
                "Позиция",
                "Наименование",
                "Марка",
                "Кол-во",
                "№ карты раскроя",
                "Кол-во в КР"
            }
        };

        var list = wcog.Keys.ToList();
        list.Sort((x, y) =>
        {
            if (int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x, y, StringComparison.Ordinal);
        });

        foreach (var key in list)
        {
            var elem = wcog[key];

            items.Add(new[]
            {
                elem.PosNo,
                $"{elem.Shape} s{elem.GetThickness()}", 
                elem.Quality,
                elem.Quantity.ToString(),
                elem.NestedOn,
                "nested quantity"
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень листовых деталей.xlsx", items);
    }
}
