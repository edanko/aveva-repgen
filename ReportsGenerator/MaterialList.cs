using System.Globalization;
using System.Text.RegularExpressions;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

class PlatePivot
{
    public double RawThickness { get; set; }
    public string Quality { get; set; }
    public double RawLength { get; set; }
    public double RawWidth { get; set; }
    public int Quantity { get; set; }
    public double TotalBurning { get; set; }
    public double TotalIdle { get; set; }
}

class ProfilePivot
{
    public string Quality { get; set; }
    public string Type { get; set; }
    public double Length { get; set; }
}

public static class MaterialList
{
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

    public static void Gen(Dictionary<string, Wcog> wcog, List<Gen> gens)
    {
        // TODO: split plate and profile to different files
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

        var startRow = 2;

        for (var i = 0; i < platePivot.Count; i++)
        {
            var elem = platePivot[i];
            var row = i + startRow;

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
        
        var profiles = wcog.Where(x => x.Value.IsProfile).ToDictionary(x => x.Key, x => x.Value);
        var profilePivot = new List<ProfilePivot>();
        foreach (var prof in profiles)
        {
            var p = profilePivot.Find(x =>
                x.Quality == prof.Value.Quality &&
                x.Type == prof.Value.Shape+prof.Value.Dimension);

            if (p == null)
            {
                p = new ProfilePivot
                {
                    Quality = prof.Value.Quality,
                    Type = prof.Value.Shape + prof.Value.Dimension
                };

                profilePivot.Add(p);
            }

            p.Length += prof.Value.TotalLength;
        }

        var nextRow = startRow + platePivot.Count + 2;

        items.Add(new [] {""});
        items.Add(new[] {"Сводная по профилю"});

        nextRow += 2;

        for (var i = 0; i < profilePivot.Count; i++)
        {
            var elem = profilePivot[i];
            var row = i + nextRow;

            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.Type,
                elem.Quality,
                elem.Length.ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkingDir}\\{Settings.Default.Drawing} - Сводная по материалам.xlsx", items);
    }
}