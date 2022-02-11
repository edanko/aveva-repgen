using System.Globalization;
using System.Text.RegularExpressions;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public class PlatePivot
{
    private double RawThickness { get; set; }
    private string Quality { get; set; }
    private double RawLength { get; set; }
    private double RawWidth { get; set; }
    private int Quantity { get; set; }
    private double TotalBurning { get; set; }
    private double TotalIdle { get; set; }

    private static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var result = "";
        var matchCollection = regex.Matches(s);
        var num = 0;

        var num2 = matchCollection.Count - 1;
        for (var i = num; i <= num2; i++) result = matchCollection[i].Value;
        return result;
    }

    public static void Gen(List<Gen> gens)
    {
        var platePivot = new List<PlatePivot>();
        foreach (var g in gens)
        {
            var p = platePivot.Find(x =>
                Math.Abs(x.RawThickness - g.RawThickness) < 0.0001 && x.Quality == g.Quality && Math.Abs(x.RawLength - g.RawLength) < 0.0001 &&
                Math.Abs(x.RawWidth - g.RawWidth) < 0.0001);

            if (p == null)
            {
                p = new PlatePivot
                {
                    RawLength = g.RawLength,
                    RawWidth = g.RawWidth,
                    Quality = g.Quality,
                    RawThickness = g.RawThickness,
                    Quantity = 1
                };

                platePivot.Add(p);
            }
            else
            {
                p.Quantity++;
            }

            p.TotalBurning += g.TotalBurning;
            p.TotalIdle += g.TotalIdle;
        }
        platePivot.Sort((x, y) => x.RawThickness.CompareTo(y.RawThickness));

        var items = new List<string[]>
        {
            new []
            {
                "№ п/п",
                "Марка",
                "Толщина",
                "Длина",
                "Ширина",
                "Кол-во",
                "Длина реза",
                "Длина ХХ",
            }
        };

        for (var i = 0; i < platePivot.Count; i++)
        {
            var elem = platePivot[i];

            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.Quality,
                elem.RawThickness.ToString(CultureInfo.InvariantCulture),
                elem.RawLength.ToString(CultureInfo.InvariantCulture),
                elem.RawWidth.ToString(CultureInfo.InvariantCulture),
                elem.Quantity.ToString(),
                elem.TotalBurning.ToString(CultureInfo.InvariantCulture),
                elem.TotalIdle.ToString(CultureInfo.InvariantCulture),

            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Сводная по листам.xlsx", items);
    }
}