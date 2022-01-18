using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ReportsGenerator;

public static class QualityList
{
    public static Dictionary<string, double> Read(string file)
    {
        var res = new Dictionary<string, double>();
        
        var lines = File.ReadAllLines(file);
        foreach (var l in lines)
        {
            if (string.IsNullOrWhiteSpace(l))
            {
                continue;
            }

            var s = l.Split(" ", StringSplitOptions.RemoveEmptyEntries);
            res.Add(s[0], double.Parse(s[2], CultureInfo.InvariantCulture));
        }
        return res;
    }
}