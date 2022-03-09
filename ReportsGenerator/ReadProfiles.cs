using System.Globalization;

namespace ReportsGenerator;

public class Profile
{
    public string Name { get; set; }
    public string Normalized { get; set; }
    public double Weight { get; set; }
    public double BarLength { get; set; }
}

public static class Profiles
{
    public static List<Profile> Read(string file)
    {
        var res = new List<Profile>();
        
        var lines = File.ReadAllLines(file);
        foreach (var l in lines)
        {
            if (string.IsNullOrWhiteSpace(l))
            {
                continue;
            }

            var s = l.Split(";", StringSplitOptions.RemoveEmptyEntries);
            if (s.Length != 4)
            {
                continue;
            }
            
            var profile = new Profile()
            {
                Name = s[0],
                Normalized = s[1],
                Weight = double.Parse(s[2], NumberStyles.Any, CultureInfo.InvariantCulture),
                BarLength = double.Parse(s[3], NumberStyles.Any, CultureInfo.InvariantCulture)
            };
            
            res.Add(profile);
        }
        return res;
    }
}