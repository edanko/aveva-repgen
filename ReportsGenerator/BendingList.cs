using DocumentFormat.OpenXml.Packaging;
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
        list.Sort();

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = parts[key];
            var row = i + 2;

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

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Ведомость гибки.xlsx", items);
    }
}