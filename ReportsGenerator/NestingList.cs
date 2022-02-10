using System.ComponentModel;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using ReportsGenerator.Properties;

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
            ExcelHelper.UpdateCell(worksheet, elem.NestingPercent.ToString(CultureInfo.InvariantCulture), row, "G");
            ExcelHelper.UpdateCell(worksheet, elem.TotalBurning.ToString(CultureInfo.InvariantCulture), row, "H");
            ExcelHelper.UpdateCell(worksheet, elem.TotalIdle.ToString(CultureInfo.InvariantCulture), row, "I");
            ExcelHelper.UpdateCell(worksheet, elem.NoOfBurningStarts.ToString(), row, "J");
            ExcelHelper.UpdateCell(worksheet, elem.PartsWeight.ToString(CultureInfo.InvariantCulture), row, "K");
            ExcelHelper.UpdateCell(worksheet, elem.RemnantWeight.ToString(CultureInfo.InvariantCulture), row, "L");
        }

        try
        {
            doc.SaveAs($"{Settings.Default.WorkingDir}\\{Settings.Default.Drawing} - Ведомость карт раскроя.xlsx");
            bw.ReportProgress(0, $"{Settings.Default.Drawing} - Ведомость карт раскроя.xlsx cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить {Settings.Default.Drawing} - Ведомость карт раскроя.xlsx\r\n");
        }
        finally
        {
            doc.Close();
        }
    }
}