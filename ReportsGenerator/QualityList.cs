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
            var density = Regexp(s[2], "[0-9]+(?:\\.[0-9]*)?(?=E-)");
            res.Add(s[0], double.Parse(s[2], CultureInfo.InvariantCulture));
        }
        return res;
    }

    private static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var result = "";
        var matchCollection = regex.Matches(s);
        var num = 0;

        var num2 = matchCollection.Count - 1;
        for (var i = num; i <= num2; i++)
        {
            result = matchCollection[i].Value;
        }
        return result;
    }
}