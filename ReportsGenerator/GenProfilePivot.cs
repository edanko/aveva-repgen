using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public class ProfilePivot
{
    private string Quality { get; set; }
    private string Type { get; set; }
    private double Length { get; set; }

    public static void Gen(Dictionary<string, Wcog> wcog)
    {
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

        var items = new List<string[]>
        {
            new []
            {
                "№ п/п",
                "Типоразмер",
                "Марка",
                "Длина",
            }
        };

        for (var i = 0; i < profilePivot.Count; i++)
        {
            var elem = profilePivot[i];

            items.Add(new[]
            {
                (i + 1).ToString(),
                elem.Type,
                elem.Quality,
                elem.Length.ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Сводная по профилю.xlsx", items);
    }
}