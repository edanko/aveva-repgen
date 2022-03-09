using System.Globalization;

namespace ReportsGenerator;

public class Wcog
{
    public string PosNo { get; private set; }
    public int Quantity { get; private set; }
    public double Weight { get; private set; }
    public string Block { get; private set; }
    public string Quality { get; private set; }
    private double Thickness { get; set; }
    public string Shape { get; private set; }
    public string Dimension { get; private set; }
    public double TotalLength { get; private set; }
    public double MouldedLength { get; private set; }
    public double CircLength { get; private set; }
    public double CircWidth { get; private set; }
    public string NestedOn { get; private set; }
    public bool IsProfile { get; private set; }
    public bool IsBent { get; private set; }

    public string GetName()
    {
        return Shape switch
        {
            "PP" => $"Полособульб {Dimension.Replace("PP", "HP")}",
            "FB" => $"Полоса s{GetThickness():G}",
            "Tube" => $"Труба D{Dimension}",
            "RBAR" => $"Пруток {Dimension}",
            _ => $"Лист s{GetThickness():G}"
        };
    }

    public double GetThickness()
    {
        if (Dimension.Length > 0)
        {
            var spl = Dimension.Split('*');
            return double.Parse(spl.Length > 1 ? spl[1] : Dimension, CultureInfo.InvariantCulture);
        }

        return Thickness;
    }

    public static Dictionary<string, Wcog> Read(string file)
    {
        var wcog = new Dictionary<string, Wcog>();

        var wcogFileLines = File.ReadAllLines(file);

        for (var i = 2; i < wcogFileLines.Length; i++)
        {
            var c = wcogFileLines[i].Split(',');
            if (c.Length == 1)
            {
                c = wcogFileLines[i].Split(';');
            }

            if (c.Length < 27)
            {
                continue;
            }

            var pos = GetPos(c[0]);

            if (wcog.ContainsKey(pos))
            {
                wcog[pos].Quantity++;
                continue;
            }

            var shape = c[23].Trim();
            var l = new Wcog
            {
                Block = c[6],
                Quality = c[11],
                NestedOn = c[18],
                Shape = shape,
                Dimension = shape+c[24],
                TotalLength = double.TryParse(c[25], NumberStyles.Any, CultureInfo.InvariantCulture, out var val2)
                    ? val2
                    : 0.0,
                MouldedLength = double.TryParse(c[26], NumberStyles.Any, CultureInfo.InvariantCulture, out var val3)
                    ? val3
                    : 0.0,
                Thickness = double.TryParse(c[22], NumberStyles.Any, CultureInfo.InvariantCulture, out var val4)
                    ? val4
                    : 0.0,
                Weight = double.TryParse(c[1], NumberStyles.Any, CultureInfo.InvariantCulture, out var val5)
                    ? val5
                    : 0.0,
                CircLength = double.TryParse(c[20], NumberStyles.Any, CultureInfo.InvariantCulture, out var val6)
                    ? val6
                    : 0.0,
                CircWidth = double.TryParse(c[21], NumberStyles.Any, CultureInfo.InvariantCulture, out var val7)
                    ? val7
                    : 0.0,
                Quantity = 1,
                PosNo = pos
            };

            if (c[8].Contains("CURVED") || c[8].Contains("BENT") || c[8].Contains("FOLDED") ||
                c[8].Contains("KNUCKLED"))
            {
                l.IsBent = true;
            }

            // Treat flat bar profile as plate part.
            if (c[8].Contains("PROFILE") && l.Shape != "FB")
            {
                l.IsProfile = true;
            }

            wcog[pos] = l;
        }

        return wcog;
    }

    private static string GetPos(string s)
    {
        var pos = s.Split('-')[^1].Replace("P", "").Replace("S", "").Replace("B", "").Replace("C", "");
        if (string.IsNullOrWhiteSpace(pos))
        {
            pos = s.Trim('-');
        }
        return pos;
    }
}