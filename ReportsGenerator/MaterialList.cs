using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ReportsGenerator.My;

namespace ReportsGenerator;

class PlatePivot
{
    public double RawThickness { get; set; }
    public string Quality { get; set; }
    public double RawLength { get; set; }
    public double RawWidth { get; set; }
    public int Quantity { get; set; }
    public double TotalBurning { get; set; }
    public double TotalIdle { get; set; }
}

class ProfilePivot
{
    public string Quality { get; set; }
    public string Type { get; set; }
    public double Length { get; set; }
}

public static class MaterialList
{
    private static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var result = "";
        var matchCollection = regex.Matches(s);
        var num = 0;

        var num2 = matchCollection.Count - 1;
        for (var i = num; i <= num2; i++) result = matchCollection[i].Value;
        return result;
    }

    public static void Gen(BackgroundWorker bw, Dictionary<string, Wcog> wcog, List<Gen> gens)
    {
        var doc = SpreadsheetDocument.Open($"{System.Windows.Forms.Application.StartupPath}\\templates\\material_list.xlsx", true);
        var worksheet = ExcelHelper.GetWorksheetPartByName(doc, "Сводная");

        ExcelHelper.UpdateCell(worksheet, gens.Count.ToString(), 1, "F");
        ExcelHelper.UpdateCell(worksheet, $"хз", 2, "F");
        ExcelHelper.UpdateCell(worksheet, $"хз", 3, "F");

        var platePivot = new List<PlatePivot>();
        foreach (var g in gens)
        {
            var p = platePivot.Find(x =>
                Math.Abs(x.RawThickness - g.RawThickness) < 0.0001 && x.Quality == g.Quality && Math.Abs(x.RawLength - g.RawLength) < 0.0001 &&
                Math.Abs(x.RawWidth - g.RawWidth) < 0.0001);

            if (p == null)
            {
                p = new PlatePivot
                {
                    RawLength = g.RawLength,
                    RawWidth = g.RawWidth,
                    Quality = g.Quality,
                    RawThickness = g.RawThickness,
                    Quantity = 1
                };

                platePivot.Add(p);
            }
            else
            {
                p.Quantity++;
            }

            p.TotalBurning += g.TotalBurning;
            p.TotalIdle += g.TotalIdle;
        }

        var startRow = 7;

        for (var i = 0; i < platePivot.Count; i++)
        {
            var elem = platePivot[i];
            var row = i + startRow;

            ExcelHelper.UpdateCell(worksheet, (i + 1).ToString(), row, "A");
            ExcelHelper.UpdateCell(worksheet, elem.Quality, row, "B");
            ExcelHelper.UpdateCell(worksheet, elem.RawThickness.ToString(), row, "C");
            ExcelHelper.UpdateCell(worksheet, elem.RawLength.ToString(), row, "D");
            ExcelHelper.UpdateCell(worksheet, elem.RawWidth.ToString(), row, "E");
            ExcelHelper.UpdateCell(worksheet, elem.Quantity.ToString(), row, "F");
            ExcelHelper.UpdateCell(worksheet, elem.TotalBurning.ToString(), row, "G");
            ExcelHelper.UpdateCell(worksheet, elem.TotalIdle.ToString(), row, "H");
        }
        
        var profiles = wcog.Where(x => x.Value.IsProfile).ToDictionary(x => x.Key, x => x.Value);
        var profilePivot = new List<ProfilePivot>();
        foreach (var prof in profiles)
        {
            var p = profilePivot.Find(x =>
                x.Quality == prof.Value.Quality &&
                x.Type == prof.Value.Shape+prof.Value.Dimension);

            if (p == null)
            {
                p = new ProfilePivot
                {
                    Quality = prof.Value.Quality,
                    Type = prof.Value.Shape + prof.Value.Dimension
                };

                profilePivot.Add(p);
            }

            p.Length += prof.Value.TotalLength;
        }

        var nextRow = startRow + platePivot.Count + 2;

        ExcelHelper.UpdateCell(worksheet, "Сводная по профилю", nextRow, "A");

        nextRow += 2;

        for (var i = 0; i < profilePivot.Count; i++)
        {
            var elem = profilePivot[i];
            var row = i + nextRow;

            ExcelHelper.UpdateCell(worksheet, (i + 1).ToString(), row, "A");
            ExcelHelper.UpdateCell(worksheet, elem.Type, row, "B");
            ExcelHelper.UpdateCell(worksheet, elem.Quality, row, "C");
            ExcelHelper.UpdateCell(worksheet, elem.Length.ToString(), row, "D");

        }

        try
        {
            doc.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Сводная материальная ведомость.xlsx");
            bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Сводная материальная ведомость.xlsx cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить {MySettingsProperty.Settings.Draw} - Сводная материальная ведомость.xlsx\r\n");
        }
        finally
        {
            doc.Close();
        }
        
        /*

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
        }*/
    }
}