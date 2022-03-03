namespace ReportsGenerator;

public static class Materials
{
    public static Dictionary<string, string> Read(string file)
    {
        var res = new Dictionary<string, string>();
        
        var lines = File.ReadAllLines(file);
        foreach (var l in lines)
        {
            if (string.IsNullOrWhiteSpace(l))
            {
                continue;
            }

            var s = l.Split(";", StringSplitOptions.RemoveEmptyEntries);
            if (s.Length != 2)
            {
                continue;
            }
            
            res.Add(s[0].ToUpper(), s[1].ToUpper());
        }
        return res;
    }
}