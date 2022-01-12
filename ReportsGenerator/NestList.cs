using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using ReportsGenerator.My;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace ReportsGenerator;

internal sealed class NestList
{
    /*public static bool Genmask(string mask, string filename)
    {
        var result = false;
        checked
        {
            var text = Strings.Right(filename, filename.Length - Strings.InStrRev(filename, "\\"));
            if (Strings.InStr(filename, ".gen") != 0)
            {
                if (Strings.InStrRev(text, "-") < 1)
                {
                    return result;
                }
                text = Strings.Left(text, Strings.InStr(text, "-") - 1);
                if (string.Equals(text, mask))
                {
                    result = true;
                }
            }
            return result;
        }
    }*/

    public static double Text2double2(object arr)
    {
        var text = arr.ToString();

        return double.Parse(text, CultureInfo.InvariantCulture);
    }

    public static object NestlistGen(BackgroundWorker bw, List<string> files)
    {
        var array = new object[files.Count, 12];
        var errorFlag = false;
        var i = 0;
        foreach (var file in files)
        {
            var lines = File.ReadAllLines(file);
            foreach (var line in lines)
                if (line.Contains("NEST_NAME"))
                {
                    array[i, 1] = line.Split('=')[1];
                }
                else if (line.Contains("RAW_THICKNESS"))
                {
                    array[i, 2] = double.Parse(line.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (line.Contains("QUALITY"))
                {
                    array[i, 3] = line.Split('=')[1];
                }
                else if (line.Contains("RAW_LENGTH"))
                {
                    var num3 = Strings.InStr(line, "=");
                    var text2 = Strings.Right(line, Strings.Len(line) - num3);
                    text2 = Strings.Left(text2, Strings.Len(text2) - 3);
                    array[i, 4] = text2;
                }
                else if (line.Contains("RAW_WIDTH"))
                {
                    var num3 = Strings.InStr(line, "=");
                    var text2 = Strings.Right(line, Strings.Len(line) - num3);
                    text2 = Strings.Left(text2, Strings.Len(text2) - 3);
                    array[i, 4] =
                        Operators.ConcatenateObject(Operators.ConcatenateObject(array[i, 4], " x "), text2);
                }
                else if (line.Contains("NO_OF_PARTS"))
                {
                    var partsNum = line.Split('=')[1];

                    array[i, 5] = int.Parse(partsNum);
                }
                else if (line.Contains("RAW_AREA"))
                {
                    array[i, 6] = double.Parse(line.Split('=')[1], CultureInfo.InvariantCulture);
                }
                else if (line.Contains("PART_AREA"))
                {
                    var num3 = Strings.InStr(line, "=");
                    var text2 = Strings.Right(line, Strings.Len(line) - num3);
                    text2 = Conversions.ToString(
                        Conversions.ToDouble(Strings.Replace(text2, ".", ",")));
                    array[i, 7] = Conversions.ToDouble(array[i, 7]) + Conversions.ToDouble(text2);
                    array[i, 11] = Operators.AddObject(array[i, 11], 1);
                }
                else if (line.Contains("TOTAL_BURNING"))
                {
                    var num3 = Strings.InStr(line, "=");
                    var text2 = Strings.Right(line, Strings.Len(line) - num3);
                    array[i, 8] = Conversions.ToLong(Strings.Replace(text2, ".", ","));
                }
                else if (line.Contains("TOTAL_IDLE"))
                {
                    var num3 = Strings.InStr(line, "=");
                    var text2 = Strings.Right(line, Strings.Len(line) - num3);
                    array[i, 9] = Conversions.ToLong(Strings.Replace(text2, ".", ","));
                }
                else if (line.Contains("NO_OF_BURNING_STARTS"))
                {
                    array[i, 10] = line.Split('=')[1];
                }

            array[i, 0] = i;
            if (Operators.ConditionalCompareObjectNotEqual(array[i, 5], array[i, 11], false))
            {
                bw.ReportProgress(0,
                    Operators.ConcatenateObject(
                        Operators.ConcatenateObject(
                            Operators.ConcatenateObject("Конфликт числа деталей в карте #", array[i, 1]), "!"),
                        "\r\n"));
                errorFlag = true;
            }

            i++;
        }

        if (errorFlag)
        {
            var dialogResult = MessageBox.Show(MyProject.Forms.Form1,
                "В картах раскроя обнаружен конфликт числа деталей. Продолжить работу?", "Предупреждение системы",
                MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
            {
                bw.ReportProgress(0,
                    $"Работа прекращена преждевременно в {Conversions.ToString(DateTime.Now)}!\r\n");
                return null;
            }

            bw.ReportProgress(0, "Работа продожена с конфликтом числа деталей в картах раскроя!\r\n");
        }

        var array4 = new object[Information.UBound(array, 2)];
        var flag2 = false;
        while (!flag2)
        {
            flag2 = true;
            var num4 = Information.LBound(array);
            var num5 = Information.UBound(array) - 1;
            for (var l = num4; l <= num5; l++)
                //3
                if (Text2double2(array[l, 2]) > Text2double2(array[l + 1, 2]))
                {
                    var num6 = Information.LBound(array, 2);
                    var num7 = Information.UBound(array, 2) - 1;
                    for (var m = num6; m <= num7; m++)
                    {
                        array4[m] = array[l, m];
                        array[l, m] = array[l + 1, m];
                        array[l + 1, m] = array4[m];
                        flag2 = false;
                    }
                }
            /*else if (Conversions.ToInteger(array[l, 2]) > Conversions.ToInteger(array[l + 1, 2]) & text2double2((array[l, 3])) == text2double2((array[l + 1, 3])))
                {
                    int num8 = Information.LBound(array, 2);
                    int num9 = Information.UBound(array, 2) - 1;
                    for (int n = num8; n <= num9; n++)
                    {
                        array4[n] = (array[l, n]);
                        array[l, n] = (array[l + 1, n]);
                        array[l + 1, n] = (array4[n]);
                        flag2 = false;
                    }
                }*/
        }

        var num10 = Information.LBound(array);
        var num11 = Information.UBound(array);
        for (i = num10; i <= num11; i++) array[i, 0] = i + 1;

        var array5 = new object[Information.UBound(array) + 1, 6];
        var num12 = Information.LBound(array);
        var num13 = Information.UBound(array);
        for (i = num12; i <= num13; i++)
        {
            array5[i, 0] = array[i, 3];
            array5[i, 1] = array[i, 2];
            array5[i, 2] = 1;
            array5[i, 3] = array[i, 8];
            array5[i, 4] = array[i, 9];
            array5[i, 5] = array[i, 4];
        }

        var num14 = 0;
        var num15 = Information.LBound(array5) + 1;
        var num16 = Information.UBound(array5);
        for (i = num15; i <= num16; i++)
        {
            var num17 = 0;
            var num18 = num14;
            var flag3 = false;
            for (var num19 = num17; num19 <= num18; num19++)
            {
                flag3 = true;
                if (Conversions.ToBoolean(Operators.AndObject(
                        Operators.CompareObjectEqual(array[i, 3], array5[num19, 0], false),
                        Operators.CompareObjectEqual(array[i, 2], array5[num19, 1], false))))
                {
                    array5[num19, 2] = Operators.AddObject(array5[num19, 2], 1);
                    array5[num19, 3] = Operators.AddObject(array5[num19, 3], array[i, 8]);
                    array5[num19, 4] = Operators.AddObject(array5[num19, 4], array[i, 9]);
                    flag3 = false;
                    num19 = num14;
                }
            }

            if (flag3)
            {
                num14++;
                array5[num14, 0] = array[i, 3];
                array5[num14, 1] = array[i, 2];
                array5[num14, 2] = 1;
                array5[num14, 3] = array[i, 8];
                array5[num14, 4] = array[i, 9];
                array5[num14, 5] = array[i, 4];
            }
        }

        var array6 = new object[num14 + 1, 6];
        var num20 = Information.LBound(array6);
        var num21 = Information.UBound(array6);
        for (i = num20; i <= num21; i++)
        {
            array6[i, 0] = array5[i, 0];
            array6[i, 1] = array5[i, 1];
            array6[i, 2] = array5[i, 2];
            array6[i, 3] = array5[i, 3];
            array6[i, 4] = array5[i, 4];
            array6[i, 5] = array5[i, 5];
        }

        int num22;
        if (Information.UBound(array) % 22 == 0)
            num22 = (int) Math.Round(Information.UBound(array) / 22.0);
        else
            num22 = (int) Math.Round(Conversion.Int(Information.UBound(array) / 22.0) + 1.0);

        int num23;
        if (Information.UBound(array6) % 12 == 0)
            num23 = (int) Math.Round(Information.UBound(array6) / 12.0);
        else
            num23 = (int) Math.Round(Conversion.Int(Information.UBound(array6) / 12.0) + 1.0);

        // TODO: change this to openxml
        var clsid = new Guid("00024500-0000-0000-C000-000000000046");
        var application =
            (Application) Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
        var workbook = application.Workbooks.Open($"{System.Windows.Forms.Application.StartupPath}\\VED_KR2.xls");
        if (num23 > 1)
        {
            var num24 = 2;
            var num25 = num23;
            for (i = num24; i <= num25; i++)
            {
                NewLateBinding.LateCall(workbook.Sheets[$"Лист{Conversions.ToString(i - 1)}"], null, "Copy",
                    new object[]
                    {
                        Missing.Value,
                        workbook.Sheets[$"Лист{Conversions.ToString(i - 1)}"]
                    }, null, null, null, true);
                NewLateBinding.LateSetComplex(workbook.Sheets[$"Лист{Conversions.ToString(i - 1)} (2)"], null,
                    "name", new object[]
                    {
                        $"Лист{Conversions.ToString(i)}"
                    }, null, null, false, true);
            }
        }

        NewLateBinding.LateSetComplex(workbook.Sheets["List2"], null, "name", new object[]
        {
            $"Лист{Conversions.ToString(num23 + 1)}"
        }, null, null, false, true);
        if (num22 > 1)
        {
            var num26 = num23 + 2;
            var num27 = num23 + num22;
            for (i = num26; i <= num27; i++)
            {
                NewLateBinding.LateCall(workbook.Sheets[$"Лист{Conversions.ToString(i - 1)}"], null, "Copy",
                    new object[]
                    {
                        Missing.Value,
                        workbook.Sheets[$"Лист{Conversions.ToString(i - 1)}"]
                    }, null, null, null, true);
                NewLateBinding.LateSetComplex(workbook.Sheets[$"Лист{Conversions.ToString(i - 1)} (2)"], null,
                    "name", new object[]
                    {
                        $"Лист{Conversions.ToString(i)}"
                    }, null, null, false, true);
            }
        }

        var num28 = num23;
        var num29 = num22;
        var now = DateAndTime.Now;
        var num30 = 0;
        var num31 = 10;
        DataTable dataTable;
        try
        {
            dataTable = (DataTable) GetMaterialDensity(bw);
            if (Information.IsNothing(dataTable))
            {
                bw.ReportProgress(0,
                    "Не удается получить данные о плотностях материалов! Проверьте существование и правильное форматирование файла sbh_quality_list.def. Выбрать файл можно в настройках!\r\n");
                return array;
            }
        }
        catch (Exception)
        {
            bw.ReportProgress(0,
                "Не удается получить данные о плотностях материалов! Проверьте существование и правильное форматирование файла sbh_quality_list.def. Выбрать файл можно в настройках!\r\n");
            return array;
        }

        var num33 = num28 + 1;
        var num34 = num28 + num29;
        for (i = num33; i <= num34; i++)
        {
            var num32 = 7.85;
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B2"
                }, null, null, null), null, "Value", new object[]
            {
                MySettingsProperty.Settings.Project
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B3"
                }, null, null, null), null, "Value", new object[]
            {
                ""
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B4"
                }, null, null, null), null, "Value", new object[]
            {
                ""
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "G5"
                }, null, null, null), null, "Value", new object[]
            {
                now
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "M33"
                }, null, null, null), null, "Value", new object[]
            {
                i
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "M34"
                }, null, null, null), null, "Value", new object[]
            {
                num28 + num29
            }, null, null, false, true);
            var num35 = 1;
            do
            {
                if (num30 <= Information.UBound(array))
                {
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"A{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 0]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"B{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 1]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"C{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 2]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"D{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 3]
                    }, null, null, false, true);
                    foreach (var obj in dataTable.Rows)
                    {
                        var dataRow = (DataRow) obj;
                        if (Operators.ConditionalCompareObjectEqual(array[num30, 3], dataRow[0], false))
                        {
                            num32 = Convert.ToDouble(dataRow[1]);
                            break;
                        }
                    }

                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"E{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 4]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"F{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 5]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"G{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.SubtractObject(1,
                            Operators.DivideObject(
                                Operators.SubtractObject(Operators.MultiplyObject(array[num30, 6], 100),
                                    Operators.DivideObject(array[num30, 7], 10000)),
                                Operators.MultiplyObject(array[num30, 6], 100)))
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"H{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.DivideObject(array[num30, 8], 1000)
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"I{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.DivideObject(array[num30, 9], 1000)
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"J{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array[num30, 10]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"K{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.DivideObject(
                            Operators.MultiplyObject(Operators.MultiplyObject(array[num30, 7], array[num30, 2]),
                                num32), 1000000.0)
                    }, null, null, false, true);
                }

                num30++;
                num31++;
                num35++;
            } while (num35 <= 22);

            num31 = 10;
        }

        num30 = 0;
        num31 = 15;
        var num36 = 0.0;
        var num37 = 0.0;
        var num38 = 1;
        var num39 = num28;
        for (i = num38; i <= num39; i++)
        {
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B2"
                }, null, null, null), null, "Value", new object[]
            {
                MySettingsProperty.Settings.Project
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B3"
                }, null, null, null), null, "Value", new object[]
            {
                ""
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "B4"
                }, null, null, null), null, "Value", new object[]
            {
                ""
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "G5"
                }, null, null, null), null, "Value", new object[]
            {
                now
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "H10"
                }, null, null, null), null, "Value", new object[]
            {
                Information.UBound(array) + 1
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "M33"
                }, null, null, null), null, "Value", new object[]
            {
                i
            }, null, null, false, true);
            NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[$"Лист{Conversions.ToString(i)}"],
                null, "Range", new object[]
                {
                    "M34"
                }, null, null, null), null, "Value", new object[]
            {
                num28 + num29
            }, null, null, false, true);
            if (i == num28)
            {
                var num40 = 0;
                var num41 = Information.UBound(array6);
                for (var num42 = num40; num42 <= num41; num42++)
                {
                    num36 = Conversions.ToDouble(Operators.AddObject(num36,
                        Operators.DivideObject(array6[num42, 3], 1000)));
                    num37 = Conversions.ToDouble(Operators.AddObject(num37,
                        Operators.DivideObject(array6[num42, 4], 1000)));
                }

                NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                    workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                    {
                        "H28"
                    }, null, null, null), null, "Value", new object[]
                {
                    "________________________________________"
                }, null, null, false, true);
                NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                    workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                    {
                        "H29"
                    }, null, null, null), null, "Value", new object[]
                {
                    num36
                }, null, null, false, true);
                NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                    workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                    {
                        "L29"
                    }, null, null, null), null, "Value", new object[]
                {
                    num37
                }, null, null, false, true);
            }

            var num43 = 1;
            do
            {
                if (num30 <= Information.UBound(array6))
                {
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"A{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array6[num30, 0]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"C{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new object[]
                    {
                        "S ="
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"D{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        array6[num30, 1]
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"E{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.ConcatenateObject(Operators.ConcatenateObject("(         ", array6[num30, 2]),
                            " к.р.)")
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"G{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new object[]
                    {
                        "Lрез="
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"H{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.DivideObject(array6[num30, 3], 1000)
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"K{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new object[]
                    {
                        "Lхол="
                    }, null, null, false, true);
                    NewLateBinding.LateSetComplex(NewLateBinding.LateGet(
                        workbook.Sheets[$"Лист{Conversions.ToString(i)}"], null, "Range", new object[]
                        {
                            $"L{Conversions.ToString(num31)}"
                        }, null, null, null), null, "Value", new[]
                    {
                        Operators.DivideObject(array6[num30, 4], 1000)
                    }, null, null, false, true);
                }

                num30++;
                num31++;
                num43++;
            } while (num43 <= 12);

            num31 = 15;
        }

        try
        {
            workbook.SaveAs(
                $"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xls",
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xls cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0,
                $"Не получилось сохранить {MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xls!\r\n");
        }
        finally
        {
            workbook.Close(false, Missing.Value, Missing.Value);
            application.Quit();
        }

        return array;
    }

    public static object GetMaterialDensity(BackgroundWorker bw)
    {
        var dataTable = new DataTable("MaterialDensity");
        var qualityList = MySettingsProperty.Settings.QualityList;
        if (File.Exists(qualityList))
        {
            var array2 = File.ReadAllLines(qualityList);
            dataTable.Columns.Add("Material");
            dataTable.Columns.Add("Density");
            foreach (var text in array2)
                if (!Information.IsNothing(text) && Operators.CompareString(text, "", false) != 0)
                {
                    var array = Strings.Split(text);
                    dataTable.Rows.Add(array[0], Regexp(array[2], "[0-9]+(?:\\.[0-9]*)?(?=E-)"));
                }

            return dataTable;
        }

        return null;
    }

    private static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var result = "";
        var matchCollection = regex.Matches(s);
        var num = 0;
        checked
        {
            var num2 = matchCollection.Count - 1;
            for (var i = num; i <= num2; i++) result = matchCollection[i].Value;
            return result;
        }
    }

    public static void MaterialListGen(BackgroundWorker bw, ArrayList myArray, Array nestMap)
    {
        var array = new object[4, 1];
        var array2 = new object[4, 1];
        var array3 = new object[4, 1];
        object obj = 0;
        object obj2 = 0;
        object obj3 = 0;
        var flag = true;
        var flag2 = true;
        var flag3 = true;
        var array4 = MySettingsProperty.Settings.NestSizeList.Split(',');
        var num = 0;
        var num2 = Information.LBound(nestMap);
        var num3 = Information.UBound(nestMap);

        for (var i = num2; i <= num3; i++)
        {
            var num4 = Information.LBound(array4);
            var num5 = Information.UBound(array4);
            for (var j = num4; j <= num5; j++)
                if (Operators.ConditionalCompareObjectEqual(array4[j], NewLateBinding.LateIndexGet(nestMap,
                        new object[]
                        {
                            i,
                            4
                        }, null), false))
                {
                    num++;
                    break;
                }
        }

        object[,] array5;
        object[,] array6;
        var array7 = new object[1, 6];
        if (num > 0)
        {
            array5 = new object[num - 1 + 1, 6];
            array6 = new object[num - 1 + 1, 13];
            num = 0;
            var num6 = Information.LBound(nestMap);
            var num7 = Information.UBound(nestMap);
            for (var k = num6; k <= num7; k++)
            {
                var num8 = Information.LBound(array4);
                var num9 = Information.UBound(array4);
                for (var l = num8; l <= num9; l++)
                    if (Operators.ConditionalCompareObjectEqual(array4[l], NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                4
                            }, null), false))
                    {
                        array5[num, 0] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                3
                            }, null);
                        array5[num, 1] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                2
                            }, null);
                        array5[num, 2] = 1;
                        array5[num, 3] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                8
                            }, null);
                        array5[num, 4] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                9
                            }, null);
                        array5[num, 5] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                4
                            }, null);
                        array6[num, 0] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                0
                            }, null);
                        array6[num, 1] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                1
                            }, null);
                        array6[num, 2] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                2
                            }, null);
                        array6[num, 3] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                3
                            }, null);
                        array6[num, 4] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                4
                            }, null);
                        array6[num, 5] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                5
                            }, null);
                        array6[num, 6] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                6
                            }, null);
                        array6[num, 7] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                7
                            }, null);
                        array6[num, 8] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                8
                            }, null);
                        array6[num, 9] = NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                k,
                                9
                            }, null);
                        num++;
                        break;
                    }
            }

            var num10 = 0;
            var num11 = Information.LBound(array5) + 1;
            var num12 = Information.UBound(array5);
            for (var m = num11; m <= num12; m++)
            {
                var num13 = 0;
                var num14 = num10;
                var flag4 = false;
                for (var n = num13; n <= num14; n++)
                {
                    flag4 = true;
                    if (Conversions.ToBoolean(Operators.AndObject(
                            Operators.CompareObjectEqual(array6[m, 3], array5[n, 0], false),
                            Operators.CompareObjectEqual(array6[m, 2], array5[n, 1], false))))
                    {
                        array5[n, 2] = Operators.AddObject(array5[n, 2], 1);
                        array5[n, 3] = Operators.AddObject(array5[n, 3], array6[m, 8]);
                        array5[n, 4] = Operators.AddObject(array5[n, 4], array6[m, 9]);
                        flag4 = false;
                        n = num10;
                    }
                }

                if (flag4)
                {
                    num10++;
                    array5[num10, 0] = array6[m, 3];
                    array5[num10, 1] = array6[m, 2];
                    array5[num10, 2] = 1;
                    array5[num10, 3] = array6[m, 8];
                    array5[num10, 4] = array6[m, 9];
                    array5[num10, 5] = array6[m, 4];
                }
            }

            array7 = new object[num10 + 1, 6];
            var num15 = Information.LBound(array7);
            var num16 = Information.UBound(array7);
            for (var num17 = num15; num17 <= num16; num17++)
            {
                array7[num17, 0] = array5[num17, 0];
                array7[num17, 1] = array5[num17, 1];
                array7[num17, 2] = array5[num17, 2];
                array7[num17, 3] = array5[num17, 3];
                array7[num17, 4] = array5[num17, 4];
                array7[num17, 5] = array5[num17, 5];
            }
        }
        else
        {
            num = -1;
        }

        var num18 = 2;
        var num19 = myArray.Count - 1;
        for (var num20 = num18; num20 <= num19; num20++)
            if (!Information.IsNothing(NewLateBinding.LateIndexGet(myArray[num20],
                    new object[]
                    {
                        23
                    }, null)))
            {
                if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                    {
                        23
                    }, null)), "PP") > 0)
                {
                    if (flag)
                    {
                        obj = 0;
                        array[0, Conversions.ToInteger(obj)] = Operators.AddObject(obj, 1);
                        array[1, Conversions.ToInteger(obj)] = NewLateBinding.LateIndexGet(myArray[num20], new object[]
                        {
                            11
                        }, null);
                        array[2, Conversions.ToInteger(obj)] = NewLateBinding.LateIndexGet(myArray[num20], new object[]
                        {
                            24
                        }, null);
                        array[3, Conversions.ToInteger(obj)] = Conversions.ToInteger(NewLateBinding.LateIndexGet(
                            myArray[num20], new object[]
                            {
                                25
                            }, null));
                        flag = false;
                    }
                    else
                    {
                        var flag5 = true;
                        var num21 = 0;
                        var num22 = Information.UBound(array, 2);
                        for (var num23 = num21; num23 <= num22; num23++)
                            if ((string.CompareOrdinal(Conversions.ToString(array[1, num23]), Conversions.ToString(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        11
                                    }, null))) == 0) & (string.CompareOrdinal(Conversions.ToString(array[2, num23]),
                                    Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        24
                                    }, null))) == 0))
                            {
                                array[3, num23] = Operators.AddObject(array[3, num23], Conversions.ToInteger(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        25
                                    }, null)));
                                flag5 = false;
                                break;
                            }

                        if (flag5)
                        {
                            obj = Operators.AddObject(obj, 1);
                            array = (object[,]) Utils.CopyArray(array,
                                new object[4, Conversions.ToInteger(obj) + 1]);
                            array[0, Conversions.ToInteger(obj)] = Operators.AddObject(obj, 1);
                            array[1, Conversions.ToInteger(obj)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    11
                                }, null);
                            array[2, Conversions.ToInteger(obj)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    24
                                }, null);
                            array[3, Conversions.ToInteger(obj)] = Conversions.ToInteger(
                                NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                {
                                    25
                                }, null));
                        }
                    }
                }
                else if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                         {
                             23
                         }, null)), "RBAR") > 0)
                {
                    if (flag2)
                    {
                        obj2 = 0;
                        array2[0, Conversions.ToInteger(obj2)] = Operators.AddObject(obj2, 1);
                        array2[1, Conversions.ToInteger(obj2)] = NewLateBinding.LateIndexGet(myArray[num20],
                            new object[]
                            {
                                11
                            }, null);
                        array2[2, Conversions.ToInteger(obj2)] = NewLateBinding.LateIndexGet(myArray[num20],
                            new object[]
                            {
                                24
                            }, null);
                        array2[3, Conversions.ToInteger(obj2)] = Conversions.ToInteger(NewLateBinding.LateIndexGet(
                            myArray[num20], new object[]
                            {
                                25
                            }, null));
                        flag2 = false;
                    }
                    else
                    {
                        var flag5 = true;
                        var num24 = 0;
                        var num25 = Information.UBound(array2, 2);
                        for (var num26 = num24; num26 <= num25; num26++)
                            if ((string.CompareOrdinal(Conversions.ToString(array2[1, num26]), Conversions.ToString(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        11
                                    }, null))) == 0) & (string.CompareOrdinal(Conversions.ToString(array2[2, num26]),
                                    Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        24
                                    }, null))) == 0))
                            {
                                array2[3, num26] = Operators.AddObject(array2[3, num26], Conversions.ToInteger(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        25
                                    }, null)));
                                flag5 = false;
                                break;
                            }

                        if (flag5)
                        {
                            obj2 = Operators.AddObject(obj2, 1);
                            array2 = (object[,]) Utils.CopyArray(array2,
                                new object[4, Conversions.ToInteger(obj2) + 1]);
                            array2[0, Conversions.ToInteger(obj2)] = Operators.AddObject(obj2, 1);
                            array2[1, Conversions.ToInteger(obj2)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    11
                                }, null);
                            array2[2, Conversions.ToInteger(obj2)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    24
                                }, null);
                            array2[3, Conversions.ToInteger(obj2)] = Conversions.ToInteger(
                                NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                {
                                    25
                                }, null));
                        }
                    }
                }
                else if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                         {
                             23
                         }, null)), "TUBE") > 0)
                {
                    if (flag3)
                    {
                        obj3 = 0;
                        array3[0, Conversions.ToInteger(obj3)] = Operators.AddObject(obj3, 1);
                        array3[1, Conversions.ToInteger(obj3)] = NewLateBinding.LateIndexGet(myArray[num20],
                            new object[]
                            {
                                11
                            }, null);
                        array3[2, Conversions.ToInteger(obj3)] = NewLateBinding.LateIndexGet(myArray[num20],
                            new object[]
                            {
                                24
                            }, null);
                        array3[3, Conversions.ToInteger(obj3)] = Conversions.ToInteger(NewLateBinding.LateIndexGet(
                            myArray[num20], new object[]
                            {
                                25
                            }, null));
                        flag3 = false;
                    }
                    else
                    {
                        var flag5 = true;
                        var num27 = 0;
                        var num28 = Information.UBound(array3, 2);
                        for (var num29 = num27; num29 <= num28; num29++)
                            if ((string.CompareOrdinal(Conversions.ToString(array3[1, num29]), Conversions.ToString(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        11
                                    }, null))) == 0) & (string.CompareOrdinal(Conversions.ToString(array3[2, num29]),
                                    Conversions.ToString(NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        24
                                    }, null))) == 0))
                            {
                                array3[3, num29] = Operators.AddObject(array3[3, num29], Conversions.ToInteger(
                                    NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                    {
                                        25
                                    }, null)));
                                flag5 = false;
                                break;
                            }

                        if (flag5)
                        {
                            obj3 = Operators.AddObject(obj3, 1);
                            array3 = (object[,]) Utils.CopyArray(array3,
                                new object[4, Conversions.ToInteger(obj3) + 1]);
                            array3[0, Conversions.ToInteger(obj3)] = Operators.AddObject(obj3, 1);
                            array3[1, Conversions.ToInteger(obj3)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    11
                                }, null);
                            array3[2, Conversions.ToInteger(obj3)] = NewLateBinding.LateIndexGet(myArray[num20],
                                new object[]
                                {
                                    24
                                }, null);
                            array3[3, Conversions.ToInteger(obj3)] = Conversions.ToInteger(
                                NewLateBinding.LateIndexGet(myArray[num20], new object[]
                                {
                                    25
                                }, null));
                        }
                    }
                }
            }

        if (flag) array = null;

        if (flag2) array2 = null;

        if (flag3) array3 = null;

        var clsid = new Guid("00024500-0000-0000-C000-000000000046");
        var application =
            (Application) Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
        var workbook = application.Workbooks.Add(Missing.Value);
        var worksheet = (Worksheet) workbook.Worksheets[1];
        worksheet.get_Range("A1:B1", Missing.Value).Merge(Missing.Value);
        worksheet.get_Range("A1", Missing.Value).set_Value(Missing.Value, "зак.");
        worksheet.get_Range("C1:D1", Missing.Value).Merge(Missing.Value);
        worksheet.get_Range("C1", Missing.Value).set_Value(Missing.Value,
            $"сек.{MySettingsProperty.Settings.Block}");
        worksheet.get_Range("E1:I1", Missing.Value).Merge(Missing.Value);
        worksheet.get_Range("E1", Missing.Value).set_Value(Missing.Value,
            $"черт.{MySettingsProperty.Settings.Draw}");
        worksheet.get_Range("C2", Missing.Value).set_Value(Missing.Value, "Name");
        worksheet.get_Range("D2", Missing.Value).set_Value(Missing.Value, "Quality");
        worksheet.get_Range("E2", Missing.Value).set_Value(Missing.Value, "Length");
        worksheet.get_Range("F2", Missing.Value).set_Value(Missing.Value, "Width");
        worksheet.get_Range("G2", Missing.Value).set_Value(Missing.Value, "Thickness");
        worksheet.get_Range("H2", Missing.Value).set_Value(Missing.Value, "Quantity");
        worksheet.get_Range("I2", Missing.Value).set_Value(Missing.Value, "N остатка");
        var num30 = 3;
        var flag6 = false;
        var num31 = Information.LBound(nestMap);
        var num32 = Information.UBound(nestMap);
        for (var num33 = num31; num33 <= num32; num33++)
            if (!Information.IsNothing(NewLateBinding.LateIndexGet(nestMap,
                    new object[]
                    {
                        num33,
                        0
                    }, null)))
            {
                worksheet.get_Range($"A{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    NewLateBinding.LateIndexGet(nestMap, new object[]
                    {
                        num33,
                        0
                    }, null));
                worksheet.get_Range($"C{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    NewLateBinding.LateIndexGet(nestMap, new object[]
                    {
                        num33,
                        1
                    }, null));
                worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    NewLateBinding.LateIndexGet(nestMap, new object[]
                    {
                        num33,
                        3
                    }, null));
                var num34 = Conversions.ToInteger(Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(
                    nestMap, new object[]
                    {
                        num33,
                        4
                    }, null)), Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(nestMap, new object[]
                {
                    num33,
                    4
                }, null)), " x ")));
                var num35 = Conversions.ToInteger(Strings.Right(Conversions.ToString(NewLateBinding.LateIndexGet(
                    nestMap, new object[]
                    {
                        num33,
                        4
                    }, null)), Strings.Len(NewLateBinding.LateIndexGet(nestMap,
                    new object[]
                    {
                        num33,
                        4
                    }, null)) - Strings.InStrRev(Conversions.ToString(NewLateBinding.LateIndexGet(nestMap,
                    new object[]
                    {
                        num33,
                        4
                    }, null)), " ")));
                worksheet.get_Range($"E{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, num34);
                worksheet.get_Range($"F{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, num35);
                worksheet.get_Range($"G{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    NewLateBinding.LateIndexGet(nestMap, new object[]
                    {
                        num33,
                        2
                    }, null));
                var num36 = Information.LBound(array4);
                var num37 = Information.UBound(array4);
                for (var num38 = num36; num38 <= num37; num38++)
                    if (Operators.ConditionalCompareObjectEqual(array4[num38], NewLateBinding.LateIndexGet(nestMap,
                            new object[]
                            {
                                num33,
                                4
                            }, null), false))
                    {
                        flag6 = true;
                        break;
                    }

                if (flag6)
                    worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value)
                        .set_Value(Missing.Value, 1);
                else
                    worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value)
                        .set_Value(Missing.Value, "");

                flag6 = false;
                num30++;
            }

        var num39 = 0;
        num30 += 3;
        if (num > -1)
        {
            worksheet.get_Range($"D{Conversions.ToString(num30)}:F{Conversions.ToString(num30)}", Missing.Value)
                .Merge(Missing.Value);
            worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value)
                .set_Value(Missing.Value, "Сводная");
            worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).HorizontalAlignment = 3;
            worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).Font.Bold = TriState.True;
            worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).Font.Size = 14;
            num30++;
            var num40 = 0;
            var num41 = Information.LBound(array7);
            var num42 = Information.UBound(array7);
            for (var num43 = num41; num43 <= num42; num43++)
                if (!Information.IsNothing(array7[num43, 0]))
                {
                    var num34 = Conversions.ToInteger(Strings.Left(Conversions.ToString(array7[num43, 5]),
                        Strings.InStr(Conversions.ToString(array7[num43, 5]), " x ")));
                    var num35 = Conversions.ToInteger(Strings.Right(Conversions.ToString(array7[num43, 5]),
                        Strings.Len(array7[num43, 5]) -
                        Strings.InStrRev(Conversions.ToString(array7[num43, 5]), " ")));
                    worksheet.get_Range($"C{Conversions.ToString(num30)}", Missing.Value)
                        .set_Value(Missing.Value, num40 + 1);
                    worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                        array7[num43, 0]);
                    worksheet.get_Range($"E{Conversions.ToString(num30)}", Missing.Value)
                        .set_Value(Missing.Value, num34);
                    worksheet.get_Range($"F{Conversions.ToString(num30)}", Missing.Value)
                        .set_Value(Missing.Value, num35);
                    worksheet.get_Range($"G{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                        array7[num43, 1]);
                    worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                        array7[num43, 2]);
                    num30++;
                    num40++;
                }

            num30++;
        }

        if (!Information.IsNothing(array))
        {
            var num44 = Information.LBound(array, 2);
            var num45 = Information.UBound(array, 2);
            for (var num46 = num44; num46 <= num45; num46++)
            {
                worksheet.get_Range($"C{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, num46 + 1);
                worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, array[1, num46]);
                worksheet.get_Range($"E{Conversions.ToString(num30)}:G{Conversions.ToString(num30)}", Missing.Value)
                    .Merge(Missing.Value);
                worksheet.get_Range($"E{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject("r ", array[2, num46]));
                worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject(NewLateBinding.LateGet(null, typeof(Math), "Round", new[]
                    {
                        Operators.DivideObject(array[3, num46], 1000),
                        1
                    }, null, null, null), " п.м."));
                num30++;
            }

            num39 += Information.UBound(array, 2);
        }
        else
        {
            num39 += 0;
        }

        if (!Information.IsNothing(array2))
        {
            var num47 = Information.LBound(array2, 2);
            var num48 = Information.UBound(array2, 2);
            for (var num49 = num47; num49 <= num48; num49++)
            {
                worksheet.get_Range($"C{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, num39 + num49 + 1);
                worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    array2[1, num49]);
                worksheet.get_Range($"E{Conversions.ToString(num30)}:G{Conversions.ToString(num30)}", Missing.Value)
                    .Merge(Missing.Value);
                worksheet.get_Range($"E{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject("Пруток d ", array2[2, num49]));
                worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject(NewLateBinding.LateGet(null, typeof(Math), "Round", new[]
                    {
                        Operators.DivideObject(array2[3, num49], 1000),
                        1
                    }, null, null, null), " п.м."));
                num30++;
            }

            num39 += Information.UBound(array2, 2);
        }
        else
        {
            num39 += 0;
        }

        if (!Information.IsNothing(array3))
        {
            var num50 = Information.LBound(array3, 2);
            var num51 = Information.UBound(array3, 2);
            for (var num52 = num50; num52 <= num51; num52++)
            {
                worksheet.get_Range($"C{Conversions.ToString(num30)}", Missing.Value)
                    .set_Value(Missing.Value, num39 + num52 + 1);
                worksheet.get_Range($"D{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    array3[1, num52]);
                worksheet.get_Range($"E{Conversions.ToString(num30)}:G{Conversions.ToString(num30)}", Missing.Value)
                    .Merge(Missing.Value);
                worksheet.get_Range($"E{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject("Труба o ", array3[2, num52]));
                worksheet.get_Range($"H{Conversions.ToString(num30)}", Missing.Value).set_Value(Missing.Value,
                    Operators.ConcatenateObject(NewLateBinding.LateGet(null, typeof(Math), "Round", new[]
                    {
                        Operators.DivideObject(array3[3, num52], 1000),
                        1
                    }, null, null, null), " п.м."));
                num30++;
            }
        }

        worksheet.get_Range($"A1:I{Conversions.ToString(num30)}", Missing.Value).Font.Size = 11;
        worksheet.get_Range($"A1:I{Conversions.ToString(num30)}", Missing.Value).Font.Name = "Calibri";
        worksheet.get_Range($"A1:I{Conversions.ToString(num30)}", Missing.Value).HorizontalAlignment = 3;
        worksheet.get_Range($"A1:I{Conversions.ToString(num30)}", Missing.Value).VerticalAlignment = 3;
        worksheet.get_Range("A1:I1", Missing.Value).Font.Bold = TriState.True;
        worksheet.get_Range("A1:I1", Missing.Value).Font.Size = 14;
        worksheet.get_Range($"A1:I{Conversions.ToString(num30)}", Missing.Value).Borders.Weight = 2;
        try
        {
            workbook.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\M.B._{MySettingsProperty.Settings.Block}.xls");
            bw.ReportProgress(0, $"M.B._{MySettingsProperty.Settings.Block}.xls cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить M.B._{MySettingsProperty.Settings.Block}.xls!\r\n");
        }
        finally
        {
            workbook.Close(false, Missing.Value, Missing.Value);
            application.Quit();
        }
    }
}