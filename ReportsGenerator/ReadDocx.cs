using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ReportsGenerator;

public class Docx
{
    public string PosNo { get; private set; }
    public int Quantity { get; private set; }
    public string Dimension { get; private set; }
    public string Quality { get; private set; }
    public double Weight { get; private set; }
    public string MaterialCode { get; private set; }
    public string MaterialListCode { get; private set; }
    public string Shape { get; private set; }
    public string Assembly { get; private set; }
    public bool IsProfile { get; private set; }

    private static string NormalizeProfileDimension(string dimension)
    {
        // TODO: add other profile dimensions
        var dict = new Dictionary<string, string>
        {
            {"5", "50*4.0"},
            {"5.5", "55*4.5"},
            {"6", "60*5.0"},
            {"7", "70*5.0"},
            {"8", "80*5.0"},
            {"9", "90*5.5"},
            {"10", "100*6.0"},
            {"12", "120*6.5"},
            {"14а", "140*7.0"},
            {"14б", "140*9.0"},
            {"16а", "160*8.0"},
            {"16б", "160*10.0"},
            {"18а", "180*9.0"},
            {"18б", "180*11.0"},
            {"20а", "200*10.0"},
            {"20б", "200*12.0"},
            {"22а", "220*11.0"},
            {"22б", "220*13.0"},
            {"24а", "240*12.0"},
            {"24б", "240*14.0"},
        };

        if (dict.ContainsKey(dimension))
        {
            return dict[dimension];
        }
        return dimension;
    }
    
    public static Dictionary<string, Docx> Read(string drawName, Dictionary<string, string> materials)
    {
        WordprocessingDocument document;
        try
        {
            document = WordprocessingDocument.Open(drawName, false);
        }
        catch (ArgumentNullException)
        {
            return null;
        }

        if (document.MainDocumentPart?.Document.Body == null)
        {
            return null;
        }

        var tables = document.MainDocumentPart.Document.Body.Elements<Table>().ToList();
        var result = Kind1(tables);
        if (result.Count == 0)
        { 
            result = Kind2(tables);
        }

        if (result.Count == 0)
        {
            result = Kind3(tables);
        }
        
        document.Close();

        foreach (var part in result)
        {
            if (materials.ContainsKey(part.Value.Quality))
            {
                result[part.Key].Quality = materials[part.Value.Quality];
            }
        }
        
        return result;
    }

    private static Dictionary<string, Docx> Kind1(IEnumerable<Table> tables)
    {
        // TODO: Handle shapes: Полособульб, Лист, Полоса, Пруток?
        var result = new Dictionary<string, Docx>();

        foreach (var table in tables)
        {
            var columnsCount = table.Elements<TableGrid>().First().ChildElements.Count;

            if (columnsCount != 20)
            {
                continue;
            }

            var lastName = "";
            foreach (var row in table.Elements<TableRow>())
            {
                var columns = row.Descendants<TableCell>().ToArray();

                if (string.IsNullOrWhiteSpace(columns[0].InnerText) && string.IsNullOrWhiteSpace(columns[1].InnerText))
                {
                    continue;
                }
                
                if (columns[1].InnerText.ToUpper().Contains("СВОДНЫЕ ДАННЫЕ"))
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(columns[0].InnerText) && columns[1].InnerText != "")
                {
                    lastName = columns[1].InnerText.Trim();
                    continue;
                }
                
                if (string.IsNullOrWhiteSpace(columns[1].InnerText))
                {
                    continue;
                }
                
                if (columns[3].InnerText != "796")
                {
                    continue;
                }

                var current = new Docx();
                current.PosNo = columns[0].InnerText;
                current.Quantity = int.Parse(columns[4].InnerText, CultureInfo.InvariantCulture);
                current.MaterialCode = columns[2].InnerText;
                current.MaterialListCode = columns[9].InnerText;
                
                var nameParts = ProcessKind1Name(columns[1].InnerText);
                current.Shape = nameParts[0];
                if (nameParts[0] != "Лист")
                {
                    current.IsProfile = true;
                }
                current.Dimension = nameParts[1];
                current.Assembly = lastName;
                current.Quality = columns[11].InnerText.ToUpper().Trim();

                var weightColumn = 6;
                if (current.Quantity > 1)
                {
                    weightColumn = 5;
                }
                current.Weight = double.Parse(columns[weightColumn].InnerText, CultureInfo.InvariantCulture);

                result.Add(current.PosNo, current);
            }
        }

        return result;
    }
    
    private static Dictionary<string, Docx> Kind2(IEnumerable<Table> tables)
    {
        // TODO: Handle shapes: Лист, Полособульб, Стенка, Поясок
        var result = new Dictionary<string, Docx>();

        foreach (var table in tables)
        {
            var columnsCount = table.Elements<TableGrid>().First().ChildElements.Count;

            if (columnsCount != 29)
            {
                continue;
            }

            var lastName = "";
            foreach (var row in table.Elements<TableRow>())
            {
                var columns = row.Descendants<TableCell>().ToArray();
                
                if (columns[3].InnerText.ToUpper().Contains("СВОДНЫЕ ДАННЫЕ"))
                {
                    break;
                }
                
                if (string.IsNullOrWhiteSpace(columns[1].InnerText) || columns.Length != 23)
                {
                    lastName = columns[3].InnerText.Trim();
                    continue;
                }
                
                if (columns[6].InnerText != "796")
                {
                    continue;
                }

                var current = new Docx();
                current.PosNo = columns[1].InnerText;
                current.Quantity = int.Parse(columns[7].InnerText, CultureInfo.InvariantCulture);
                current.MaterialCode = columns[5].InnerText;
                current.MaterialListCode = columns[12].InnerText;

                var nameParts = columns[3].InnerText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (nameParts.Length == 2)
                {
                    current.Shape = nameParts[0];
                    current.Dimension = nameParts[1];
                }
                
                current.Assembly = lastName;
                current.Quality = columns[14].InnerText.ToUpper().Trim();

                var weightColumn = 9;
                if (current.Quantity > 1)
                {
                    weightColumn = 8;
                }
                current.Weight = double.Parse(columns[weightColumn].InnerText, CultureInfo.InvariantCulture);

                result.Add(current.PosNo, current);
            }
        }

        return result;
    }
    
    private static Dictionary<string, Docx> Kind3(IEnumerable<Table> tables)
    {
        var result = new Dictionary<string, Docx>();

        foreach (var table in tables)
        {
            var columnsCount = table.Elements<TableGrid>().First().ChildElements.Count;

            if (columnsCount != 21)
            {
                continue;
            }

            var rows = table.Elements<TableRow>().ToList();
            for (var i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var columns = row.Descendants<TableCell>().ToArray();

                if (columns[2].InnerText.ToUpper().Contains("СВОДНЫЕ ДАННЫЕ"))
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(columns[1].InnerText) || columns.Length != 21)
                {
                    continue;
                }

                if (columns[4].InnerText != "796")
                {
                    continue;
                }

                var current = new Docx();
                current.PosNo = columns[1].InnerText;
                current.MaterialCode = columns[3].InnerText;
                current.Quantity = int.Parse(columns[5].InnerText, CultureInfo.InvariantCulture);

                var nameParts = ProcessKind3Name(rows[i + 1].Descendants<TableCell>().ToArray()[2].InnerText);
                if (nameParts.Length == 3)
                {
                    current.Shape = nameParts[0];
                    if (nameParts[0] != "Лист")
                    {
                        current.IsProfile = true;
                    }

                    current.Dimension = nameParts[1];
                    current.Assembly = nameParts[2];
                }
                
                current.MaterialListCode = columns[10].InnerText;
                current.Quality = columns[12].InnerText.ToUpper().Trim();
                current.Weight = double.Parse(columns[7].InnerText, CultureInfo.InvariantCulture);

                result.Add(current.PosNo, current);
            }
        }

        return result;
    }
    
    private static string[] ProcessKind1Name(string s)
    {
        var regex = new Regex(@"(\S+)\s+(\S+).*");
        var matchCollection = regex.Matches(s);

        var match = matchCollection[0];
        if (match.Groups.Count == 3)
        {
            return new[]
            {
                match.Groups[1].Value,
                match.Groups[2].Value,
            };
        }
        
        return new [] {s};
    }
    
    private static string[] ProcessKind3Name(string s)
    {
        var regex = new Regex(@"(\S+)\s+(\S+)\s{4}(.*)");
        var matchCollection = regex.Matches(s);

        var match = matchCollection[0];
        if (match.Groups.Count == 4)
        {
            return new[]
            {
                match.Groups[1].Value,
                match.Groups[2].Value,
                match.Groups[3].Value
            };
        }
        
        return new [] {s};
    }
}