using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using ReportsGenerator.My.Resources;

namespace ReportsGenerator
{
    internal static class Doc
    {
        public static object Read(BackgroundWorker bw, string drawName)
        {
            var dataTable = new System.Data.DataTable("Sortament");
            var array = new DataColumn[2];
            dataTable.Columns.Add("NAME");
            dataTable.Columns.Add("DIM");
            array[0] = dataTable.Columns["NAME"];
            dataTable.PrimaryKey = array;
            var array2 = Resources.Sortament.Split(new[]
            {
                '\t'
            });
            foreach (var text in array2)
            {
                var array4 = text.Split(',');
                dataTable.Rows.Add(new object[]
                {
                    array4[0],
                    array4[1]
                });
            }

            var clsid = new Guid("000209FF-0000-0000-C000-000000000046");
            var application =
                (Application) Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
            Document document;
            try
            {
                document = application.Documents.Open(drawName);
            }
            catch (Exception)
            {
                bw.ReportProgress(0, $"Не удается открыть файл {drawName} проверьте, существует ли он!\r\n");
                application.Quit(false);
                return null;
            }
            
            Table table = document.Tables.Cast<Table>().FirstOrDefault(t => t.Columns.Count == 23);

            /*if (document.Tables[1].Rows.Count < 1)
            {
                return null;
            }*/

            //var table = document.Tables[4];

            var range = table?.ConvertToText();
            var array5 = new object[5, 1];
            var tableLines = range?.Text.Split('\r');
            document.Close(false);
            application.Quit(false);
            var num = 0; 
            var startLine = 5; // was 0
            var totalLines = tableLines?.Length;

            for (var j = startLine; j <= totalLines; j++)
            {
                bw.ReportProgress(0, Operators.AddObject("@ ", NewLateBinding.LateIndexGet(tableLines, new object[]
                {
                    j
                }, null)));

                var columns = tableLines[j].Split('-');
                

                if (columns[3].ToUpper().Contains("СВОДНЫЕ ДАННЫЕ"))
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(columns[1]) || columns.Length != 23)
                {
                    continue;
                }
                
                array5 = (object[,]) Utils.CopyArray(array5, new object[5, num + 1]);
                array5[0, num] = columns[1]; // поз
                array5[1, num] = columns[7]; // кол-во
                array5[2, num] = ""; // columns[9]; маршрут
                var thickness = DataProcessor.Regexp(" " + columns[3] + " ",
                        "(?<!,)(?<=[S,s])[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[x|х][0-9]{1,}[\\.|,]{0,1}[0-9]{0,}(?=\\s{1,})|(?<=\\s{1})[R,r](?<!,)[0-9]{1,}[\\.|,]{0,1}[0-9]{0,}[a,b,а,б]{0,1}(?=\\s{1,})|(?<=[П][р][у][т][о][к][ ])[0-9]{1,}(?=\\s{1,})|[0-9]{1,}[\\*][0-9]{1,3}[\\.|,]{0,1}[0-9]{0,}")
                    .ToLower().Replace(",", ".");

                if (thickness.Contains("полособульб"))
                {
                    thickness = thickness.Trim().Replace("полособульб ", "r");
                }

                if (!string.IsNullOrWhiteSpace(thickness))
                {
                    if (Strings.InStr(thickness, "x") > 0)
                    {
                        var array6 = thickness.Split(new[]
                        {
                            'x'
                        });
                        if (Strings.InStr(array6[0], ".") < 1)
                        {
                            array6[0] = $"{array6[0]}.0";
                        }

                        thickness = $"{array6[1]}*{array6[0]}";
                    }
                    else if (Strings.InStr(thickness, "х") > 0)
                    {
                        var array6 = thickness.Split(new[]
                        {
                            'х'
                        });
                        if (Strings.InStr(array6[0], ".") < 1)
                        {
                            array6[0] = $"{array6[0]}.0";
                        }

                        thickness = $"{array6[1]}*{array6[0]}";
                    }
                    else if (Strings.InStr(thickness, "r") > 0 | Strings.InStr(thickness, "р") > 0)
                    {
                        var dataRow =
                            dataTable.Rows.Find(thickness.Replace("r", "").Replace("р", "").Replace("a", "а"));
                        if (!Information.IsNothing(dataRow))
                        {
                            thickness = dataRow[1].ToString();
                        }
                    }
                }

                array5[3, num] = thickness;
                array5[4, num] = columns[14];
                num++;
                
            }
            
            return array5;
        }

        public static object RenameMaterials(BackgroundWorker bw, string s)
        {
            s = s.ToUpper();
            return s.Replace("D500W", "DW").Replace("D500CB", "DCB").Replace("E500W", "EW").Replace("E500CB", "ECB")
                .Replace("E500Z-П", "E500W").Replace("45Г17Ю3", "45G").Replace("F500W", "FW").Replace("СП 20", "ST20")
                .Replace("СТ3СП ГОФРИРОВАННАЯ", "SP3PS_125").Replace("СТ3СП", "SP3PS_143").Replace("08Х18Н10Т", "Н10")
                .Replace("E36Z35", "E36Z").Replace("D36Z35", "D36Z").Replace("БЕТОН СЕРПЕНТИНИТОВЫЙ", "BS")
                .Replace("БЕТОН СЕРПЕНТИНИТОВЫЙ С КАРБИДОМ БОРА", "BSB").Replace("А", "A").Replace("Н", "H")
                .Replace(" ", "");
        }
    }
}