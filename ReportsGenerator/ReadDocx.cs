using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DataTable = System.Data.DataTable;

namespace ReportsGenerator;

public class Docx
{
    // TODO: add block, route, assembly, is profile
    public string PosNo { get; private set; }
    public int Quantity { get; private set; }
    public string Dimension { get; private set; }
    public string Quality { get; private set; }
    public double Weight { get; private set; }

    public static Dictionary<string, Docx> Read(string drawName)
    {
        // FIXME: wtf is going on here? extract to function?
        const string sortament = "5,50*4.0,2.25	5.5,55*4.5,2.73	6,60*5.0,3.36	7,70*5.0,3.98	8,80*5.0,4.58	9,90*5.5,5.52	10,100*6.0,6.76	12,120*6.5,8.75	14а,140*7.0,11.05	14б,140*9.0,13.23	16а,160*8.0,14.08	16б,160*10.0,16.60	18а,180*9.0,17.41	18б,180*11.0,20.24	20а,200*10.0,21.47	20б,200*12.0,24.60	22а,220*11.0,25.75	22б,220*13.0,29.20	24а,240*12.0,30.42	24б,240*14.0,34.18	90*90*6.0,90*6.0,8.33	63*63*6.0,63*6.0,5.72	108*6.0,108*6.0,15.09	133*8.0,133*8.0,24.66	168*9.0,168*9.0,35.29";

        var dataTable = new DataTable("Sortament");
        var array = new DataColumn[2];
        dataTable.Columns.Add("NAME");
        dataTable.Columns.Add("DIM");
        array[0] = dataTable.Columns["NAME"];
        dataTable.PrimaryKey = array;
        var array2 = sortament.Split('\t');
        foreach (var text in array2)
        {
            var array4 = text.Split(',');
            dataTable.Rows.Add(array4[0], array4[1]);
        }

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

        var result = new Dictionary<string, Docx>();

        foreach (var table in document.MainDocumentPart.Document.Body.Elements<Table>())
        {
            var columnsCount = table.Elements<TableGrid>().First().ChildElements.Count;

            if (columnsCount != 29)
            {
                continue;
            }

            foreach (var row in table.Elements<TableRow>())
            {
                var columns = row.Descendants<TableCell>().ToArray();

                if (columns[3].InnerText.ToUpper().Contains("СВОДНЫЕ ДАННЫЕ"))
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(columns[1].InnerText) || columns.Length != 23)
                {
                    continue;
                }

                var current = new Docx();

                current.PosNo = columns[1].InnerText;
                current.Quantity = int.Parse(columns[7].InnerText, CultureInfo.InvariantCulture);
                
                var dimension = DataProcessor.Regexp(" " + columns[3].InnerText + " ",
                        "(?<!,)(?<=[S,s])[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[x|х][0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|(?<=\\s{1})[R,r](?<!,)[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[a,b,а,б]{0,1}(?=\\s{1,})|(?<=[П][р][у][т][о][к][ ])[0-9]{1,}(?=\\s{1,})|[0-9]{1,}[\\*][0-9]{1,3}[\\.|,]{0,1}[0-9]{0,}")
                    .ToLower().Replace(",", ".");

                if (dimension.Contains("полособульб"))
                {
                    dimension = dimension.Trim().Replace("полособульб ", "r");
                }

                /*if (!string.IsNullOrWhiteSpace(dimension))
                {
                    if (dimension.IndexOf("x") < 0)
                    {
                        var array6 = dimension.Split('x');
                        if (array6[0].Contains("."))
                        {
                            array6[0] = $"{array6[0]}.0";
                        }

                        dimension = $"{array6[1]}*{array6[0]}";
                    }
                    else if (dimension.IndexOf("х") > 0)
                    {
                        var array6 = dimension.Split('x');
                        if (array6[0].IndexOf(".") < 1)
                        {
                            array6[0] = $"{array6[0]}.0";
                        }

                        dimension = $"{array6[1]}*{array6[0]}";
                    }
                    else if (dimension.Contains("r") || dimension.Contains("p"))
                    {
                        var value =
                            dataTable.Rows.Find(dimension.Replace("r", "").Replace("р", "").Replace("a", "а"));
                        if (value != null)
                        {
                            dimension = value[1].ToString();
                        }
                    }
                }*/

                current.Dimension = dimension;

                current.Quality = RenameMaterial(columns[14].InnerText);

                if (current.Quantity > 1)
                {
                    current.Weight = double.Parse(columns[8].InnerText, CultureInfo.InvariantCulture);
                }
                else
                {
                    current.Weight = double.Parse(columns[9].InnerText, CultureInfo.InvariantCulture);
                }

                result.Add(current.PosNo, current);
            }
        }
        document.Close();

        return result;
    }

    private static string RenameMaterial(string s)
    {
        s = s.ToUpper();
        return s.Replace("D500W", "DW")
            .Replace("D500CB", "DCB")
            .Replace("E500W", "EW")
            .Replace("E500CB", "ECB")
            .Replace("E500Z-П", "E500W")
            .Replace("45Г17Ю3", "45G")
            .Replace("F500W", "FW")
            .Replace("СП 20", "ST20")
            .Replace("СТ3СП ГОФРИРОВАННАЯ", "SP3PS_125")
            .Replace("СТ3СП", "SP3PS_143")
            .Replace("08Х18Н10Т", "Н10")
            .Replace("E36Z35", "E36Z")
            .Replace("D36Z35", "D36Z")
            .Replace("БЕТОН СЕРПЕНТИНИТОВЫЙ", "BS")
            .Replace("БЕТОН СЕРПЕНТИНИТОВЫЙ С КАРБИДОМ БОРА", "BSB")
            .Replace("А", "A")
            .Replace("Н", "H")
            .Replace(" ", "");
    }
}