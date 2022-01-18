using System;
using System.Collections.Generic;
using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using ReportsGenerator.My;

namespace ReportsGenerator;

public static class NestingList
{
    public static void Gen(BackgroundWorker bw, List<Gen> gens)
    {
        var doc = SpreadsheetDocument.Open($"{System.Windows.Forms.Application.StartupPath}\\templates\\nesting_list.xlsx", true);

        var worksheet = ExcelHelper.GetWorksheetPartByName(doc, "Ведомость");
        for (var i = 0; i < gens.Count; i++)
        {
            var elem = gens[i];
            var row = i + 2;
            
            ExcelHelper.UpdateCell(worksheet, (i + 1).ToString(), row, "A");
            ExcelHelper.UpdateCell(worksheet, elem.NestName, row, "B");
            ExcelHelper.UpdateCell(worksheet, elem.RawThickness.ToString("F1"), row, "C");
            ExcelHelper.UpdateCell(worksheet, elem.Quality, row, "D");
            ExcelHelper.UpdateCell(worksheet, $"{elem.RawLength:G}x{elem.RawWidth:G}", row, "E");
            ExcelHelper.UpdateCell(worksheet, elem.NoOfParts.ToString(), row, "F");
            ExcelHelper.UpdateCell(worksheet, elem.NestingPercent.ToString(), row, "G");
            ExcelHelper.UpdateCell(worksheet, elem.TotalBurning.ToString(), row, "H");
            ExcelHelper.UpdateCell(worksheet, elem.TotalIdle.ToString(), row, "I");
            ExcelHelper.UpdateCell(worksheet, elem.NoOfBurningStarts.ToString(), row, "J");
            ExcelHelper.UpdateCell(worksheet, elem.PartsWeight.ToString(), row, "K");
            ExcelHelper.UpdateCell(worksheet, elem.RemnantWeight.ToString(), row, "L");
        }

        try
        {
            doc.SaveAs($"{MySettingsProperty.Settings.WorkDir}\\{MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xlsx");
            bw.ReportProgress(0, $"{MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xlsx cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить {MySettingsProperty.Settings.Draw} - Ведомость карт раскроя.xlsx\r\n");
        }
        finally
        {
            doc.Close();
        }
    }
}