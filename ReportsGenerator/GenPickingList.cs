using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class PickingList
{
    public static void Gen(Dictionary<string, Wcog> wcog, List<Gen> gens)
    {
        var items = new List<string[]>
        {
            new string[]{},
            new []{"", "Перечень листовых деталей"},
            new []
            {
                "Позиция",
                "Наименование",
                "Марка",
                "Кол-во",
                "№ карты раскроя",
                "Кол-во в КР",
                "Масса"
            }
        };

        var list = wcog.Keys.ToList();
        list.Sort((x, y) =>
        {
            if (int.TryParse(x, out var a) && int.TryParse(y, out var b))
            {
                return a.CompareTo(b);
            }
            return string.Compare(x, y, StringComparison.Ordinal);
        });

        foreach (var key in list)
        {
            var elem = wcog[key];
            if (elem.IsProfile)
            {
                continue;
            }

            var ncs = gens.FindAll(x => x.Parts.ContainsKey(elem.PosNo));
            if (ncs.Any())
            {
                for (var i = 0; i < ncs.Count; i++)
                {
                    var nc = ncs[i];
                    if (i == 0)
                    {
                        items.Add(new[]
                        {
                            elem.PosNo,
                            elem.GetName(),
                            elem.Quality,
                            elem.Quantity.ToString(),
                            nc.NestName,
                            nc.Parts[elem.PosNo].ToString(CultureInfo.InvariantCulture),
                            (elem.Weight * nc.Parts[elem.PosNo]).ToString(CultureInfo.InvariantCulture)
                        });
                    }
                    else
                    {
                        items.Add(new[]
                        {
                            "",
                            "",
                            "",
                            "",
                            nc.NestName,
                            nc.Parts[elem.PosNo].ToString(CultureInfo.InvariantCulture),
                            (elem.Weight * nc.Parts[elem.PosNo]).ToString(CultureInfo.InvariantCulture)
                        });
                    }
                }
            }
            else
            {
                items.Add(new[]
                {
                    elem.PosNo,
                    elem.GetName(), 
                    elem.Quality,
                    elem.Quantity.ToString(),
                    elem.NestedOn,
                    elem.Quantity.ToString(),
                    (elem.Weight * elem.Quantity).ToString(CultureInfo.InvariantCulture)
                });
            }
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Перечень листовых деталей.xlsx", items);
    }
}
