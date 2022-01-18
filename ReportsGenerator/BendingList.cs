using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ReportsGenerator.My;

namespace ReportsGenerator;

public static class BendingList
{
    public static string GetNameAndDimentionString(Wcog elem)
    {
        var t = elem.GetThickness().ToString("F1");
        var dim = elem.Dimension.Split("*");

        return elem.Shape switch
        {
            "PP" => $"Полособульб s{t} {t} x {dim[1]} x {elem.TotalLength}",
            "FB" => $"Полоса s{t} {t} x {dim[1]} x {elem.TotalLength}",
            "Tube" => $"Труба D{elem.Dimension} x {elem.TotalLength}",
            "RBAR" => $"Пруток {elem.Dimension} x {elem.TotalLength}",
            _ => $"Лист s{elem.GetThickness():F1} {elem.GetThickness():F1} x {elem.CircLength} x {elem.CircWidth}"
        };
    }

    public static void Gen(BackgroundWorker bw, Dictionary<string,Wcog> parts)
    {
        var doc = SpreadsheetDocument.Open($"{System.Windows.Forms.Application.StartupPath}\\templates\\bending_list.xlsx", true);
        var worksheet = ExcelHelper.GetWorksheetPartByName(doc, "list1");
        
        var list = parts.Keys.ToList();
        list.Sort();

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = parts[key];
            var row = i + 2;

            ExcelHelper.UpdateCell(worksheet, (i + 1).ToString(), row, "A");
            ExcelHelper.UpdateCell(worksheet, MySettingsProperty.Settings.Draw, row, "B");
            ExcelHelper.UpdateCell(worksheet, elem.PosNo, row, "C");
            ExcelHelper.UpdateCell(worksheet, elem.Quantity.ToString(), row, "D");
            ExcelHelper.UpdateCell(worksheet, GetNameAndDimentionString(elem), row, "E");
            ExcelHelper.UpdateCell(worksheet, elem.NestedOn, row, "F");
        }
        
        try
        {
            doc.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Ведомость гибки.xlsx");
            bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Ведомость гибки.xlsx cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить {MySettingsProperty.Settings.Draw} - Ведомость гибки.xlsx\r\n");
        }
        finally
        {
            doc.Close();
        }
    }
}