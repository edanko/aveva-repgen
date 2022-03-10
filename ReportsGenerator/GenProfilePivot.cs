using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public class ProfilePivot
{
    private string Quality { get; set; }
    private string Dimension { get; set; }
    private double Length { get; set; }

    private string GetName()
    {
        if (Dimension.Contains("PP"))
        {
            return $"Полособульб {Dimension.Replace("PP", "HP")}";
        }

        return Dimension;
    }
    
    public static void Gen(Dictionary<string, Wcog> wcog, List<Profile> dict)
    {
        var profiles = wcog.Where(x => x.Value.IsProfile).ToDictionary(x => x.Key, x => x.Value);
        var profilePivot = new List<ProfilePivot>();
        foreach (var prof in profiles)
        {
            var p = profilePivot.Find(x =>
                x.Quality == prof.Value.Quality &&
                x.Dimension == prof.Value.Dimension);

            if (p == null)
            {
                p = new ProfilePivot
                {
                    Quality = prof.Value.Quality,
                    Dimension = prof.Value.Dimension
                };

                profilePivot.Add(p);
            }

            p.Length += prof.Value.TotalLength;
        }

        profilePivot.Sort((x, y) =>
        {
            if(int.TryParse(x.Dimension.Split("*")[^1], out var a) && int.TryParse(y.Dimension.Split("*")[^1], out var b))
            {
                return a.CompareTo(b);
            }
            return String.Compare(x.Dimension.Split("*")[^1], y.Dimension.Split("*")[^1], StringComparison.Ordinal);
        });
        
        var items = new List<string[]>
        {
            new string[]{},
            new []{"", "Сводная ведомость материалов (профиль)"},
            new []
            {
                "№ п/п",
                "Типоразмер",
                "Марка",
                "Длина заготовки, мм",
                "Кол-во заготовок",
                "Масса заготовки, кг",
                "Норма расхода, кг"
            }
        };

        for (var i = 0; i < profilePivot.Count; i++)
        {
            var elem = profilePivot[i];

            var profileData = dict.Find(x => x.Normalized == elem.Dimension);
            if (profileData == null)
            {
                MessageBox.Show($"Типоразмер \"{elem.Dimension}\" не найден в profiles.csv");
                continue;
            }
            
            var totalLength = elem.Length;
            var barLength = profileData.BarLength;
            var numBars = Math.Ceiling(totalLength / barLength);
            var barWeight = barLength / 1000 * profileData.Weight;
            
            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.GetName(),
                elem.Quality,
                barLength.ToString(CultureInfo.InvariantCulture),
                numBars.ToString(CultureInfo.InvariantCulture),
                barWeight.ToString(CultureInfo.InvariantCulture),
                (barWeight*numBars).ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Сводная по профилю.xlsx", items);
    }
}