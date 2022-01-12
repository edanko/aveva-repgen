using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using ReportsGenerator.My;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ReportsGenerator
{
	internal sealed class PartList
	{
		public static ArrayList PartlistGen(BackgroundWorker bw, string wcogFile, string partlistFile, string docFile)
		{
			var wcogLines = File.ReadAllLines(wcogFile);
			var arrayList = new ArrayList();
			var arrayList2 = new ArrayList();
			checked
			{
				var num = 0;
				int j;
				foreach (var line in wcogLines)
				{
					var elems = line.Split(',');
					if (num > 1 && elems.Length != 0)
					{
						if (string.IsNullOrEmpty(elems[0]))
						{
							Interaction.MsgBox("Пустая строка! Сейчас будет ошибка!");
						}
						j = int.Parse(DataProcessor.Regexp(elems[0], "[0-9]+(?:[0-9]*)"));
					}
					else
					{
						j = 1;
					}
					if (elems.Length != 0 && j != 0)
					{
							var textCopy = line;
						if (Information.UBound(elems) < 26)
						{
							var num3 = Information.UBound(elems);
							for (j = 25; j >= num3; j += -1)
							{
								textCopy += ",";
							}
						}
						arrayList.Add(Strings.Split(textCopy, ","));
						num++;
					}
					else if (!Information.IsNothing(elems))
					{
						bw.ReportProgress(0, $"Нулевой элемент в wcog, деталь {elems[7]}!\r\n");
					}
				}
				var partlistLines = File.ReadAllLines(partlistFile);
				var num4 = 0;
				foreach (var line in partlistLines)
				{
					var elems = line.Split(',');
					if (num4 > 1 & elems.Length != 0)
					{
						j = int.Parse(DataProcessor.Regexp(elems[0], "[0-9]+(?:[0-9]*)"));
					}
					else
					{
						j = 1;
					}
					if (!Information.IsNothing(elems) & j != 0)
					{
							var text2Copy = line;
						if (Information.UBound(elems) < 26)
						{
							var num5 = 25;
							var num6 = Information.UBound(elems);
							for (j = num5; j >= num6; j += -1)
							{
								text2Copy += ",";
							}
						}
						arrayList2.Add(text2Copy.Split(','));
						num4++;
					}
					else if (!Information.IsNothing(elems))
					{
						bw.ReportProgress(0, $"Нулевой элемент в partlist, деталь {elems[7]}!\r\n");
					}
				}
				var flag = false;
				for (j = 2; j <= num - 1; j++)
				{
					NewLateBinding.LateIndexSetComplex(arrayList[j], new object[]
					{
						0,
						DataProcessor.Regexp(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
						{
							0
						}, null)), "[0-9]+(?:[0-9]*)")
					}, null, false, true);
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						0
					}, null), 0, false))
					{
						bw.ReportProgress(0, $"Нулевой элемент в wcog, строка {Conversions.ToString(j)}!\r\n");
						flag = true;
					}
				}
				for (j = 1; j <= num4 - 1; j++)
				{
					NewLateBinding.LateIndexSetComplex(arrayList2[j], new object[]
					{
						0,
						DataProcessor.Regexp(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList2[j], new object[]
						{
							0
						}, null)), "[0-9]+(?:[0-9]*)")
					}, null, false, true);
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(arrayList2[j], new object[]
					{
						0
					}, null), 0, false))
					{
						bw.ReportProgress(0, $"Нулевой элемент в partlist, строка {Conversions.ToString(j)}!\r\n");
						flag = true;
					}
				}
				if (flag)
				{
					return null;
				}
				int l;
				for (j = 2; j <= num - 1; j++)
				{
					l = 1;
					var flag2 = false;
					while (!flag2)
					{
						if (Conversions.ToInteger(NewLateBinding.LateIndexGet(arrayList[j], new object[]
						{
							0
						}, null)) == Conversions.ToInteger(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
						{
							0
						}, null)) & Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
						{
							8
						}, null)), "used") == 0)
						{
							NewLateBinding.LateIndexSetComplex(arrayList[j], new[]
							{
								1,
								(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
								{
									4
								}, null))
							}, null, false, true);
							if (Conversions.ToBoolean(Operators.AndObject(Operators.CompareObjectEqual(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								18
							}, null), "", false), Operators.CompareObjectNotEqual(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
							{
								13
							}, null), "", false))))
							{
								NewLateBinding.LateIndexSetComplex(arrayList[j], new[]
								{
									18,
									(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
									{
										13
									}, null))
								}, null, false, true);
							}
							if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								23
							}, null)), "FB") > 0)
							{
								var value = (int)Math.Round(Math.Round(double.Parse(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
								{
									16
								}, null) as string, CultureInfo.InvariantCulture) + 0.5));
								NewLateBinding.LateIndexSetComplex(arrayList[j], new object[]
								{
									25,
									Conversions.ToString(value)
								}, null, false, true);
								NewLateBinding.LateIndexSetComplex(arrayList[j], new object[]
								{
									26,
									Conversions.ToString(value)
								}, null, false, true);
							}
							if (Conversions.ToInteger(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
							{
								1
							}, null)) > 1)
							{
								NewLateBinding.LateIndexSetComplex(arrayList2[l], new object[]
								{
									1,
									Conversions.ToInteger(NewLateBinding.LateIndexGet(arrayList2[l], new object[]
									{
										1
									}, null)) - 1
								}, null, false, true);
							}
							else
							{
								NewLateBinding.LateIndexSetComplex(arrayList2[l], new object[]
								{
									8,
									"used"
								}, null, false, true);
							}
							flag2 = true;
						}
						l++;
						if (l > arrayList2.Count - 1)
						{
							flag2 = true;
						}
					}
				}

                object obj = new object[12, 1];
				var array6 = new object[Information.UBound((Array)arrayList[0]) + 1];
				var flag3 = false;
				int n;
				while (!flag3)
				{
					flag3 = true;
					var num13 = 2;
					var num14 = num - 2;
					for (var m = num13; m <= num14; m++)
					{
						if (Conversion.Val((NewLateBinding.LateIndexGet(arrayList[m], new object[]
						{
							0
						}, null))) > Conversion.Val((NewLateBinding.LateIndexGet(arrayList[m + 1], new object[]
						{
							0
						}, null))))
						{
							var num15 = Information.LBound((Array)arrayList[0]);
							var num16 = Information.UBound((Array)arrayList[0]);
							for (n = num15; n <= num16; n++)
							{
								array6[n] = (NewLateBinding.LateIndexGet(arrayList[m], new object[]
								{
									n
								}, null));
								NewLateBinding.LateIndexSetComplex(arrayList[m], new[]
								{
									n,
									(NewLateBinding.LateIndexGet(arrayList[m + 1], new object[]
									{
										n
									}, null))
								}, null, false, true);
								NewLateBinding.LateIndexSetComplex(arrayList[m + 1], new[]
								{
									n,
									(array6[n])
								}, null, false, true);
								flag3 = false;
							}
						}
					}
				}
				var num17 = 0;
				var num18 = num - 1;
				for (var m = num17; m <= num18; m++)
				{
					NewLateBinding.LateIndexSetComplex(arrayList[m], new object[]
					{
						2,
						1
					}, null, false, true);
				}
                var arrayList4 = new ArrayList
                {
                    arrayList[0]
                };
                n = 0;
				var flag4 = true;
				var num20 = num - 1;
				for (var m = 1; m <= num20; m++)
				{
					var num21 = 0;
					var num22 = n;
					for (l = num21; l <= num22; l++)
					{
						if (Conversions.ToBoolean(Operators.AndObject(Operators.CompareObjectEqual(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							0
						}, null), NewLateBinding.LateIndexGet(arrayList[m], new object[]
						{
							0
						}, null), false), Operators.CompareObjectEqual(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							18
						}, null), NewLateBinding.LateIndexGet(arrayList[m], new object[]
						{
							18
						}, null), false))))
						{
							var obj2 = arrayList4[l];
							var instance = obj2;
							var array7 = new object[2];
							var array8 = array7;
							var num23 = 0;
							var num24 = 2;
							array8[num23] = num24;
							array7[1] = Operators.AddObject(NewLateBinding.LateIndexGet(obj2, new object[]
							{
								num24
							}, null), 1);
							NewLateBinding.LateIndexSetComplex(instance, array7, null, false, true);
							flag4 = false;
							break;
						}
					}
					if (flag4)
					{
						n++;
						arrayList4.Add(arrayList[m]);
					}
					flag4 = true;
				}
				num = n + 1;
				var docx = File.Exists(docFile) ? Doc.Read(bw, docFile) : null;
				
				var num28 = num - 1;
				for (l = 2; l <= num28; l++)
				{
					var flag5 = false;
					obj = (object[,])Utils.CopyArray((Array)obj, new object[12, l - 2 + 1]);
					NewLateBinding.LateIndexSet(obj, new object[]
					{
						0,
						l - 2,
						MySettingsProperty.Settings.Block
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						1,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							0
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						2,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							2
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						3,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							22
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						4,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							11
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						5,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							18
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						6,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							23
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						7,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							24
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						8,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							25
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						9,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							26
						}, null))
					}, null);
					NewLateBinding.LateIndexSet(obj, new[]
					{
						10,
						l - 2,
						(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
						{
							1
						}, null))
					}, null);
					if (!Information.IsNothing((docx)))
					{
						var num29 = 0;
						var num30 = Information.UBound((Array)docx, 2);
						for (j = num29; j <= num30; j++)
						{
							if (Math.Abs(Conversion.Val((NewLateBinding.LateIndexGet(docx, new object[]
                                {
                                    0,
                                    j
                                }, null))) - Conversion.Val((NewLateBinding.LateIndexGet(arrayList4[l], new object[]
                                {
                                    0
                                }, null)))) < 0.0001)
							{
								NewLateBinding.LateIndexSet(obj, new[]
								{
									11,
									l - 2,
									(NewLateBinding.LateIndexGet(docx, new object[]
									{
										2,
										j
									}, null))
								}, null);
								if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
								{
									11
								}, null), NewLateBinding.LateIndexGet(docx, new object[]
								{
									4,
									j
								}, null), false))
								{
									bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
									{
										0
									}, null), " конфликт материалов (WCOG - Спец.): "), NewLateBinding.LateIndexGet(arrayList4[l], new object[]
									{
										11
									}, null)), " "), NewLateBinding.LateIndexGet(docx, new object[]
									{
										4,
										j
									}, null)), " !"), "\r\n"));
								}
								if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
								{
									24
								}, null), "", false))
								{
									if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
									{
										24
									}, null), NewLateBinding.LateIndexGet(docx, new object[]
									{
										3,
										j
									}, null), false))
									{
										bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
										{
											0
										}, null), " конфликт типоразмеров (WCOG - Спец.): "), NewLateBinding.LateIndexGet(arrayList4[l], new object[]
										{
											24
										}, null)), " "), NewLateBinding.LateIndexGet(docx, new object[]
										{
											3,
											j
										}, null)), " !"), "\r\n"));
									}
								}
								else
								{
									try
									{
										if (decimal.Compare(Convert.ToDecimal((NewLateBinding.LateIndexGet(arrayList4[l], new object[]
										{
											22
										}, null))), Convert.ToDecimal((NewLateBinding.LateIndexGet(docx, new object[]
										{
											3,
											j
										}, null)))) != 0)
										{
											bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
											{
												0
											}, null), " конфликт типоразмеров (WCOG - Спец.): "), NewLateBinding.LateIndexGet(arrayList4[l], new object[]
											{
												22
											}, null)), " "), NewLateBinding.LateIndexGet(docx, new object[]
											{
												3,
												j
											}, null)), " !"), "\r\n"));
										}
									}
									catch (Exception)
									{
										bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
										{
											0
										}, null), " конфликт типоразмеров (WCOG - Спец.): "), NewLateBinding.LateIndexGet(arrayList4[l], new object[]
										{
											22
										}, null)), " "), NewLateBinding.LateIndexGet(docx, new object[]
										{
											3,
											j
										}, null)), " !"), "\r\n"));
									}
								}
								flag5 = true;
								break;
							}
						}
						if (!flag5)
						{
							bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(arrayList4[l], new object[]
							{
								0
							}, null), " отсутствует в .doc!"), "\r\n"));
						}
					}
					else
					{
						NewLateBinding.LateIndexSet(obj, new object[]
						{
							11,
							l - 2,
							""
						}, null);
					}
				}
				if (!Information.IsNothing((docx)))
				{
					var num31 = 0;
					var num32 = Information.UBound((Array)docx, 2);
					for (j = num31; j <= num32; j++)
					{
						var flag5 = false;
						var num33 = 2;
						var num34 = num - 1;
						for (l = num33; l <= num34; l++)
						{
							if (Math.Abs(Conversion.Val((NewLateBinding.LateIndexGet(docx, new object[]
                                {
                                    0,
                                    j
                                }, null))) - Conversion.Val((NewLateBinding.LateIndexGet(arrayList4[l], new object[]
                                {
                                    0
                                }, null)))) < 0.0001)
							{
								flag5 = true;
								break;
							}
						}
						if (!flag5)
						{
							bw.ReportProgress(0, Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(docx, new object[]
							{
								0,
								j
							}, null), " отсутствует в WCOG!"), "\r\n"));
						}
					}
				}
				var clsid = new Guid("00024500-0000-0000-C000-000000000046");
				var application = (Application)Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
				var workbook = application.Workbooks.Open(
					$"{System.Windows.Forms.Application.StartupPath}\\perech_det_shabl22220.xls", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
				var worksheet = (Worksheet)workbook.Sheets["list"];
				worksheet.get_Range("A3", Missing.Value).set_Value(Missing.Value, "зак.");
				worksheet.get_Range("C3", Missing.Value).set_Value(Missing.Value,
					$"сек.{MySettingsProperty.Settings.Block}");
				worksheet.get_Range("E3", Missing.Value).set_Value(Missing.Value,
					$"черт.{MySettingsProperty.Settings.Draw}");
				worksheet.get_Range("C5", Missing.Value).set_Value(Missing.Value, num - 2);
				worksheet.get_Range("L5", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(arrayList4[1], new object[]
				{
					1
				}, null)));
				var num35 = Information.LBound((Array)obj, 2);
				var num36 = Information.UBound((Array)obj, 2);
				for (j = num35; j <= num36; j++)
				{
					l = j + 6;
					worksheet.get_Range($"B{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						0,
						j
					}, null)));
					worksheet.get_Range($"C{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						1,
						j
					}, null)));
					worksheet.get_Range($"D{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						2,
						j
					}, null)));
					worksheet.get_Range($"E{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						3,
						j
					}, null)));
					worksheet.get_Range($"F{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						4,
						j
					}, null)));
					worksheet.get_Range($"G{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						5,
						j
					}, null)));
					worksheet.get_Range($"H{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						6,
						j
					}, null)));
					worksheet.get_Range($"I{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						7,
						j
					}, null)));
					worksheet.get_Range($"J{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						8,
						j
					}, null)));
					worksheet.get_Range($"K{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						9,
						j
					}, null)));
					worksheet.get_Range($"L{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						10,
						j
					}, null)));
					worksheet.get_Range($"M{Conversions.ToString(l)}", Missing.Value).set_Value(Missing.Value, (NewLateBinding.LateIndexGet(obj, new object[]
					{
						11,
						j
					}, null)));
				}
				worksheet.Range[$"A5:O{Conversions.ToString(l)}", Missing.Value].Borders.Weight = 2;
				try
				{
					worksheet.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Перечень деталей.xls");
					bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Перечень деталей.xls cоздан\r\n");
				}
				catch (Exception)
				{
					bw.ReportProgress(0,$"Не получилось сохранить{MySettingsProperty.Settings.Draw} - Перечень деталей.xls!\r\n");
				}
				finally
				{
					workbook.Close(false, Missing.Value, Missing.Value);
					application.Quit();
				}
				//j = 1;
				var num37 = 0;
				var num38 = 2;
				var num39 = num - 1;
				for (j = num38; j <= num39; j++)
				{
					if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						8
					}, null)), "CURVED") != 0 | Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						12
					}, null)), "BENT") != 0 | Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						12
					}, null)), "FOLDED") != 0 | Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						8
					}, null)), "KNUCKLED") != 0)
					{
						num37++;
					}
				}
				if (num37 <= 0)
				{
					bw.ReportProgress(0, "Гнутые детали не обнаружены!\r\n");
					return arrayList;
				}
				bw.ReportProgress(0, $"Гнутых деталей найдено: {Conversions.ToString(num37)}\r\n");
				object obj4 = new object[num37, 7];
				l = 0;
				for (j = 2; j <= num - 1; j++)
                {
					if (Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						8
					}, null)).Contains("CURVED") || Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						12
					}, null)).Contains("BENT") || Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						12
					}, null)).Contains("FOLDED") || Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
					{
						8
					}, null)).Contains("KNUCKLED"))
					{
						NewLateBinding.LateIndexSet(obj4, new object[]
						{
							l,
							0,
							l + 1
						}, null);
						NewLateBinding.LateIndexSet(obj4, new object[]
						{
							l,
							1,
							MySettingsProperty.Settings.Draw
						}, null);
						NewLateBinding.LateIndexSet(obj4, new object[]
						{
							l,
							2,
							Conversions.ToInteger(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								0
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new object[]
						{
							l,
							3,
							1
						}, null);
						if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
						{
							8
						}, null)), "PLATE") != 0)
						{
							var num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								22
							}, null)), ".");
							string text3;
							if (num42 != 0 & Conversions.ToInteger(Strings.Mid(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								22
							}, null)), num42 + 1, 1)) > 0)
							{
								text3 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									22
								}, null)), num42 + 1);
							}
							else if (num42 != 0)
							{
								text3 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									22
								}, null)), num42 - 1);
							}
							else
							{
								text3 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									22
								}, null));
							}
							NewLateBinding.LateIndexSet(obj4, new object[]
							{
								l,
								4,
								$"Лист s{text3}"
							}, null);
							num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								20
							}, null)), ".");
							string text4;
							if (num42 != 0)
							{
								text4 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									20
								}, null)), num42 - 1);
							}
							else
							{
								text4 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									20
								}, null));
							}
							num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								21
							}, null)), ".");
							string text5;
							if (num42 != 0)
							{
								text5 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									21
								}, null)), num42 - 1);
							}
							else
							{
								text5 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									21
								}, null));
							}
							NewLateBinding.LateIndexSet(obj4, new object[]
							{
								l,
								5,
								string.Concat(new[]
								{
									text3,
									" x ",
									text4,
									" x ",
									text5
								})
							}, null);
						}
						else if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
						{
							8
						}, null)), "PROFILE") != 0)
						{
							if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								23
							}, null)), "PP") != 0)
							{
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									4,
									"Полособульб"
								}, null);
							}
							if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								23
							}, null)), "FB") != 0)
							{
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									4,
									"Полоса"
								}, null);
							}
							if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								23
							}, null)), "Tube") != 0)
							{
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									4,
									"Труба"
								}, null);
							}
							if (Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								23
							}, null)), "RBAR") != 0)
							{
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									4,
									"Пруток"
								}, null);
								var text4 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null));
								var num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									25
								}, null)), ".");
								string text3;
								if (num42 != 0)
								{
									text3 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
									{
										25
									}, null)), num42 - 1);
								}
								else
								{
									text3 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
									{
										25
									}, null));
								}
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									5,
									$"D{text4} x {text3}"
								}, null);
							}
							else
							{
								var num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null)), "*");
								var text4 = Strings.Right(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null)), Strings.Len((NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null))) - num42);
								num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null)), "*");
								var text5 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									24
								}, null)), num42 - 1);
								num42 = Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
								{
									25
								}, null)), ".");
								string text3;
								if (num42 != 0)
								{
									text3 = Strings.Left(Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
									{
										25
									}, null)), num42 - 1);
								}
								else
								{
									text3 = Conversions.ToString(NewLateBinding.LateIndexGet(arrayList[j], new object[]
									{
										25
									}, null));
								}
								NewLateBinding.LateIndexSet(obj4, new object[]
								{
									l,
									5,
									string.Concat(new[]
									{
										text4,
										" x ",
										text5,
										" x ",
										text3
									})
								}, null);
							}
						}
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							l,
							6,
							(NewLateBinding.LateIndexGet(arrayList[j], new object[]
							{
								18
							}, null))
						}, null);
						l++;
					}
				}

                var num43 = 0;
				var num44 = Information.LBound((Array)obj4) + 1;
				var num45 = Information.UBound((Array)obj4);
				for (j = num44; j <= num45; j++)
				{
					var num46 = 0;
					var num47 = num43;
					var flag6 = false;
					for (l = num46; l <= num47; l++)
					{
						flag6 = true;
						if (Conversions.ToBoolean(Operators.AndObject(Operators.CompareObjectEqual(NewLateBinding.LateIndexGet(obj4, new object[]
						{
							j,
							2
						}, null), NewLateBinding.LateIndexGet(obj4, new object[]
						{
							l,
							2
						}, null), false), Operators.CompareObjectEqual(NewLateBinding.LateIndexGet(obj4, new object[]
						{
							j,
							5
						}, null), NewLateBinding.LateIndexGet(obj4, new object[]
						{
							l,
							5
						}, null), false))))
						{
							NewLateBinding.LateIndexSet(obj4, new[]
							{
								l,
								3,
								Operators.AddObject(NewLateBinding.LateIndexGet(obj4, new object[]
								{
									l,
									3
								}, null), 1)
							}, null);
							if (Conversions.ToBoolean(Operators.AndObject(Strings.InStr(Conversions.ToString(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								num43,
								6
							}, null)), Conversions.ToString(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								6
							}, null))) == 0, Operators.CompareObjectNotEqual(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								num43,
								6
							}, null), "", false))))
							{
								NewLateBinding.LateIndexSet(obj4, new[]
								{
									num43,
									6,
									Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(obj4, new object[]
									{
										num43,
										6
									}, null), ", "), NewLateBinding.LateIndexGet(obj4, new object[]
									{
										j,
										6
									}, null))
								}, null);
							}
							else if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								num43,
								6
							}, null), "", false))
							{
								NewLateBinding.LateIndexSet(obj4, new[]
								{
									num43,
									6,
									(NewLateBinding.LateIndexGet(obj4, new object[]
									{
										j,
										6
									}, null))
								}, null);
							}
							flag6 = false;
							l = num43;
						}
					}
					if (flag6)
					{
						num43++;
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							0,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								0
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							1,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								1
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							2,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								2
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new object[]
						{
							num43,
							3,
							1
						}, null);
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							4,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								4
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							5,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								5
							}, null))
						}, null);
						NewLateBinding.LateIndexSet(obj4, new[]
						{
							num43,
							6,
							(NewLateBinding.LateIndexGet(obj4, new object[]
							{
								j,
								6
							}, null))
						}, null);
					}
				}
				var array9 = new object[num43+1, 7];
				var num48 = Information.LBound((Array)obj4);
				for (j = num48; j <= num43; j++)
				{
					array9[j, 0] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						0
					}, null));
					array9[j, 1] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						1
					}, null));
					array9[j, 2] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						2
					}, null));
					array9[j, 3] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						3
					}, null));
					array9[j, 4] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						4
					}, null));
					array9[j, 5] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						5
					}, null));
					array9[j, 6] = (NewLateBinding.LateIndexGet(obj4, new object[]
					{
						j,
						6
					}, null));
				}
				/*var flag7 = false;
				while (!flag7)
				{
					flag7 = true;
					var num50 = Information.LBound(array9);
					var num51 = Information.UBound(array9) - 1;
					for (var m = num50; m <= num51; m++)
					{
						if (Text2double(bw, (array9[m, 6])) > Text2double(bw, (array9[m + 1, 6])))
						{
							var num52 = Information.LBound(array9, 2);
							var num53 = Information.UBound(array9, 2) - 1;
							for (n = num52; n <= num53; n++)
							{
								var text6 = Conversions.ToString(array9[m, n]);
								array9[m, n] = (array9[m + 1, n]);
								array9[m + 1, n] = text6;
								flag7 = false;
							}
						}
						else if (Conversions.ToBoolean(Operators.AndObject(Operators.CompareObjectGreater(Conversion.Int((array9[m, 2])), Conversion.Int((array9[m + 1, 2])), false), Text2double(bw, (array9[m, 6])) == Text2double(bw, (array9[m + 1, 6])))))
						{
							var num54 = Information.LBound(array9, 2);
							var num55 = Information.UBound(array9, 2) - 1;
							for (n = num54; n <= num55; n++)
							{
								var text6 = Conversions.ToString(array9[m, n]);
								array9[m, n] = (array9[m + 1, n]);
								array9[m + 1, n] = text6;
								flag7 = false;
							}
						}
					}
				}*/
				var num56 = Information.LBound(array9);
				var num57 = Information.UBound(array9);
				for (j = num56; j <= num57; j++)
				{
					array9[j, 0] = j + 1;
				}
				int num58;
				if (Information.UBound(array9) % 14 == 0)
				{
					num58 = (int)Math.Round(Information.UBound(array9) / 14.0);
				}
				else
				{
					num58 = (int)Math.Round(Conversion.Int(Information.UBound(array9) / 14.0) + 1.0);
				}
				clsid = new Guid("00024500-0000-0000-C000-000000000046");
				application = (Application)Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
				workbook = application.Workbooks.Open($"{System.Windows.Forms.Application.StartupPath}\\VED_GIB2.xls", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
				if (num58 > 1)
				{
					var num59 = 2;
					var num60 = num58;
					for (j = num59; j <= num60; j++)
					{
						NewLateBinding.LateCall(workbook.Sheets[$"Лист{Conversions.ToString(j - 1)}"], null, "Copy", new object[]
						{
							Missing.Value,
							(workbook.Sheets[$"Лист{Conversions.ToString(j - 1)}"])
						}, null, null, null, true);
						NewLateBinding.LateSetComplex(workbook.Sheets[$"Лист{Conversions.ToString(j - 1)} (2)"], null, "name", new object[]
						{
							$"Лист{Conversions.ToString(j)}"
						}, null, null, false, true);
					}
				}
				var now = DateAndTime.Now;
				var num61 = 0;
				var num62 = 8;
				var num63 = 1;
				var num64 = num58;
				for (j = num63; j <= num64; j++)
				{
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"C2"
					}, null, null, null), null, "Value", new object[]
					{
						MySettingsProperty.Settings.Project
					}, null, null, false, true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"C3"
					}, null, null, null), null, "Value", new object[]
					{
						""
					}, null, null, false, true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"C4"
					}, null, null, null), null, "Value", new object[]
					{
						""
					}, null, null, false, true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"G4"
					}, null, null, null), null, "Value", new object[]
					{
						now
					}, null, null, false, true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"K36"
					}, null, null, null), null, "Value", new object[]
					{
						j
					}, null, null, false, true);
					NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
						$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
					{
						"K37"
					}, null, null, null), null, "Value", new object[]
					{
						num58
					}, null, null, false, true);
					var num65 = 1;
					do
					{
						if (num61 <= Information.UBound(array9))
						{
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"A{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 0])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"B{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 1])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"C{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 2])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"D{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 3])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"E{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 4])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"E{Conversions.ToString(num62 + 1)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 5])
							}, null, null, false, true);
							NewLateBinding.LateSetComplex(NewLateBinding.LateGet(workbook.Sheets[
								$"Лист{Conversions.ToString(j)}"], null, "Range", new object[]
							{
								$"H{Conversions.ToString(num62)}"
							}, null, null, null), null, "Value", new[]
							{
								(array9[num61, 6])
							}, null, null, false, true);
						}
						num61++;
						num62 += 2;
						num65++;
					}
					while (num65 <= 14);
					num62 = 8;
				}
				try
				{
					workbook.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Ведомость гибки.xls");
					bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Ведомость гибки.xls cоздана\r\n");
				}
				catch (Exception)
				{
					bw.ReportProgress(0,$"Не получилось сохранить {MySettingsProperty.Settings.Draw} - Ведомость гибки.xls!\r\n");
				}
				finally
				{
					workbook.Close(false, Missing.Value, Missing.Value);
					application.Quit();
				}
				return arrayList;
			}
		}

		public static double Text2double(BackgroundWorker bw, object arr)
		{
			var text = arr.ToString();

            double ret = 0.0;

            try
            {
                ret = double.Parse(text, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e.Message}: {text}");
            }

            return ret;
        }


	}
}
