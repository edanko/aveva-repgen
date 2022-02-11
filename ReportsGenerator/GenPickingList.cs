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
                "№ п/п",
                "Секция",
                "Позиция",
                "Кол-во",
                "Толщина",
                "Марка",
                "Карта раскроя",
                "Shape",
                "Dimension",
                "Total Length",
                "Moulded Length",
                "Масса",
                "Маршрут",
                "Сборка",
                "АРЭ"
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

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = wcog[key];

            items.Add(new[]
            {
                (i+1).ToString(),
                elem.Block,
                elem.PosNo,
                elem.Quantity.ToString(),
                elem.GetThickness().ToString(CultureInfo.InvariantCulture),
                elem.Quality,
                elem.NestedOn,
                elem.Shape,
                elem.Dimension,
                elem.TotalLength.ToString(CultureInfo.InvariantCulture),
                elem.MouldedLength.ToString(CultureInfo.InvariantCulture),
                elem.Weight.ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень деталей.xlsx", items);
    }
}
