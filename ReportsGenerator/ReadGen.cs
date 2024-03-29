﻿using System.Globalization;

namespace ReportsGenerator;

public class Gen
{
    public string NestName { get; private set; }
    public double RawThickness { get; private set; }
    public string Quality { get; private set; }
    public double RawLength { get; private set; }
    public double RawWidth { get; private set; }
    public int NoOfParts { get; private set; }
    public double RawArea { get; private set; }
    public double PartsArea { get; private set; }
    public double TotalBurning { get; private set; }
    public double TotalIdle { get; private set; }
    public int NoOfBurningStarts { get; private set; }
    public double RawWeight { get; private set; }
    public double PartsWeight { get; private set; }
    public double RemnantWeight { get; private set; }
    public double NestingPercent { get; private set; }
    public int QuantityNormal { get; private set; }
    public int QuantityMirrored { get; private set; }
    public Dictionary<string,uint> Parts { get; private set; }
    public static List<Gen> Read(List<string> files, Dictionary<string, double> densityList)
    {
        var res = new List<Gen>();

        foreach (var file in files)
        {
            var g = new Gen
            {
                Parts = new Dictionary<string, uint>()
            };

            var lines = File.ReadAllLines(file);
            for (var i = 0; i < lines.Length; i++)
            {
                var l = lines[i];
                if (l.Contains("NEST_NAME"))
                {
                    g.NestName = l.Split('=')[1];
                }
                else if (l.Contains("RAW_THICKNESS"))
                {
                    g.RawThickness = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("QUALITY"))
                {
                    g.Quality = l.Split('=')[1];
                }
                else if (l.Contains("RAW_LENGTH"))
                {
                    g.RawLength = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("RAW_WIDTH"))
                {
                    g.RawWidth = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("NO_OF_PARTS"))
                {
                    g.NoOfParts = int.Parse(l.Split('=')[1]);
                }
                else if (l.Contains("RAW_AREA"))
                {
                    g.RawArea = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("PART_AREA"))
                {
                    g.PartsArea += double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("TOTAL_BURNING"))
                {
                    g.TotalBurning = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("TOTAL_IDLE"))
                {
                    g.TotalIdle = double.Parse(l.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (l.Contains("NO_OF_BURNING_STARTS"))
                {
                    g.NoOfBurningStarts = int.Parse(l.Split('=')[1]);
                }
                else if (l.Contains("QUANTITY_NORMAL"))
                {
                    g.QuantityNormal = int.Parse(l.Split('=')[1]);
                }
                else if (l.Contains("QUANTITY_MIRRORED"))
                {
                    g.QuantityMirrored = int.Parse(l.Split('=')[1]);
                }

                else if (l == "PART_DATA")
                {
                    var posno = lines[i + 2].Split('=')[1];
                    if (g.Parts.ContainsKey(posno))
                    {
                        g.Parts[posno]++;
                    }
                    else
                    {
                        g.Parts[posno] = 1;
                    }

                    i += 2;
                }
            }

            g.RawWeight = g.RawThickness * g.RawLength * g.RawWidth * densityList[g.Quality];
            g.PartsWeight = g.PartsArea * g.RawThickness * densityList[g.Quality];
            g.RemnantWeight = g.RawWeight - g.PartsWeight;
            g.NestingPercent = 1 - g.RemnantWeight / g.RawWeight;

            res.Add(g);
        }

        return res;
    }
}