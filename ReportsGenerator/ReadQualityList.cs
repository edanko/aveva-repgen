﻿using System.Globalization;

namespace ReportsGenerator;

public static class QualityList
{
    public static Dictionary<string, double> Read(string file)
    {
        var res = new Dictionary<string, double>();

        double lastDensity = 0.0;
        
        var lines = File.ReadAllLines(file);
        foreach (var l in lines)
        {
            if (string.IsNullOrWhiteSpace(l))
            {
                continue;
            }

            var s = l.Split(" ", StringSplitOptions.RemoveEmptyEntries);
            if (s[2] == "*")
            {
                res.Add(s[0], lastDensity);
                continue;
            }

            var val = double.Parse(s[2], CultureInfo.InvariantCulture);
            res.Add(s[0], val);
            lastDensity = val;
        }
        return res;
    }
}