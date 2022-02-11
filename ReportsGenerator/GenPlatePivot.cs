using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public class PlatePivot
{
    private double RawThickness { get; set; }
    private string Quality { get; set; }
    private double RawLength { get; set; }
    private double RawWidth { get; set; }
    private int Quantity { get; set; }

    public static void Gen(List<Gen> gens, Dictionary<string, double> qualityList)
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
        }
        platePivot.Sort((x, y) => x.RawThickness.CompareTo(y.RawThickness));

        var items = new List<string[]>
        {
            new []
            {
                "№ п/п",
                "Толщина, мм",
                "Ширина, мм",
                "Длина, мм",
                "Марка материала",
                "Кол-во листов, шт.",
                "Вес 1-го листа, кг",
                "Общий вес, кг",
            }
        };

        for (var i = 0; i < platePivot.Count; i++)
        {
            var elem = platePivot[i];

            var plateWeight = elem.RawThickness * elem.RawWidth * elem.RawLength * qualityList[elem.Quality];
            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.RawThickness.ToString(CultureInfo.InvariantCulture),
                elem.RawWidth.ToString(CultureInfo.InvariantCulture),
                elem.RawLength.ToString(CultureInfo.InvariantCulture),
                elem.Quality,
                elem.Quantity.ToString(),
                plateWeight.ToString(CultureInfo.InvariantCulture),
                (plateWeight*elem.Quantity).ToString(CultureInfo.InvariantCulture)
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Сводная по листам.xlsx", items);
    }
}