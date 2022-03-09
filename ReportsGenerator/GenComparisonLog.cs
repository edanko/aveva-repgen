using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class ComparisonLog
{
    public static void Gen(Dictionary<string, Wcog> wcog, Dictionary<string, Docx> docx)
    {
        var items = new List<string[]>
        {
            new[]
            {
                "Позиция",
                "Тип ошибки",
                "WCOG",
                "Спец."
            }
        };

        var wcogKeys = wcog.Keys;
        var docxKeys = docx.Keys;

        var keysOnlyInDocx = docxKeys.Except(wcogKeys).ToList();
        keysOnlyInDocx.Sort((x, y) =>
        {
            if (int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x, y, StringComparison.Ordinal);
        });
        foreach (var k in keysOnlyInDocx)
        {
            items.Add(new[]
            {
                k,
                "Отсутствует в WCOG",
            });
        }

        var keysOnlyInWcog = wcogKeys.Except(docxKeys).ToList();
        keysOnlyInWcog.Sort((x, y) =>
        {
            if (int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x, y, StringComparison.Ordinal);
        });
        foreach (var k in keysOnlyInWcog)
        {
            items.Add(new[]
            {
                k,
                "Отсутствует в спецификации",
            });
        }

        var common = wcogKeys.Intersect(docxKeys).ToList();
        common.Sort((x, y) =>
        {
            if (int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x, y, StringComparison.Ordinal);
        });
        foreach (var k in common)
        {
            if (wcog[k].Quality != docx[k].Quality)
            {
                items.Add(new[]
                {
                    k,
                    "Конфликт материалов",
                    wcog[k].Quality,
                    docx[k].Quality
                });
            }

            if (wcog[k].IsProfile)
            {
                if (wcog[k].Dimension != docx[k].Dimension)
                {
                    items.Add(new[]
                    {
                        k,
                        "Конфликт типоразмеров",
                        wcog[k].Dimension,
                        docx[k].Dimension
                    });
                }
            }
            else
            {
                if (Math.Abs(wcog[k].GetThickness() - docx[k].Thickness) > 0.001)
                {
                    items.Add(new[]
                    {
                        k,
                        "Конфликт толщин",
                        wcog[k].GetThickness().ToString(CultureInfo.InvariantCulture),
                        docx[k].Thickness.ToString(CultureInfo.InvariantCulture)
                    });
                }
            }
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Лог сравнения WCOG и спецификации.xlsx", items);
    }
}