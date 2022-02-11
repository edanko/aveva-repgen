using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class BendingList
{
    private static string GetNameAndDimentionString(Wcog elem)
    {
        var t = elem.GetThickness().ToString("F1");
        var dim = elem.Dimension.Split("*");

        return elem.Shape switch
        {
            "PP" => $"Полособульб s{t} {t} x {dim[1]} x {elem.TotalLength}",
            "FB" => $"Полоса s{t} {t} x {dim[1]} x {elem.TotalLength}",
            "Tube" => $"Труба D{elem.Dimension} x {elem.TotalLength}",
            "RBAR" => $"Пруток {elem.Dimension} x {elem.TotalLength}",
            _ => $"Лист s{elem.GetThickness():F1} {elem.GetThickness():F1} x {elem.CircLength} x {elem.CircWidth}"
        };
    }

    public static void Gen(Dictionary<string,Wcog> parts)
    {
        var items = new List<string[]>
        {
            new[]
            {
                "№ п/п",
                "Номер чертежа",
                "Позиция",
                "Кол-во",
                "Наименование и основные размеры",
                "Карта раскроя",
                "Шифр операции",
                "Оборудование",
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
                elem.Quantity.ToString(),
                GetNameAndDimentionString(elem),
                elem.NestedOn,
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень гнутых деталей.xlsx", items);
    }
}