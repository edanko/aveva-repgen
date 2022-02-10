using System.ComponentModel;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class PickingList
{
    public static void Gen(BackgroundWorker bw, Dictionary<string, Wcog> wcog, Dictionary<string, Docx> docx)
    {
        var wcogKeys = wcog.Keys;
        var docxKeys = docx.Keys;

        var keysOnlyInDocx = docxKeys.Except(wcogKeys);
        foreach (var k in keysOnlyInDocx)
        {
            bw.ReportProgress(0, $"{k} отсутствует в WCOG!\r\n");
        }

        var keysOnlyInWcog = wcogKeys.Except(docxKeys);
        foreach (var k in keysOnlyInWcog)
        {
            bw.ReportProgress(0, $"{k} отсутствует в спецификации!\r\n");
        }

        var common = wcogKeys.Intersect(docxKeys).ToList();
        common.Sort();
        foreach (var k in common)
        {
            if (wcog[k].Quality != docx[k].Quality)
            {
                bw.ReportProgress(0, $"{k} конфликт материалов (WCOG - Спец.): {wcog[k].Quality} {docx[k].Quality}\r\n");
            }

            if (wcog[k].Dimension != docx[k].Dimension)
            {
                bw.ReportProgress(0, $"{k} конфликт типоразмеров (WCOG - Спец.): {wcog[k].Dimension} {docx[k].Dimension}\r\n");
            }
        }

        var doc = SpreadsheetDocument.Open($"{System.Windows.Forms.Application.StartupPath}\\templates\\picking_list.xlsx", true);
        var worksheet = ExcelHelper.GetWorksheetPartByName(doc, "list");
        
        var list = wcogKeys.ToList();
        list.Sort();

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = wcog[key];
            var row = i + 2;

            ExcelHelper.UpdateCell(worksheet, (i+1).ToString(), row, "A");
            ExcelHelper.UpdateCell(worksheet, elem.Block, row, "B");
            ExcelHelper.UpdateCell(worksheet, elem.PosNo, row, "C");
            ExcelHelper.UpdateCell(worksheet, elem.Quantity.ToString(), row, "D");
            ExcelHelper.UpdateCell(worksheet, elem.GetThickness().ToString(CultureInfo.InvariantCulture), row, "E");
            ExcelHelper.UpdateCell(worksheet, elem.Quality, row, "F");
            ExcelHelper.UpdateCell(worksheet, elem.NestedOn, row, "G");
            ExcelHelper.UpdateCell(worksheet, elem.Shape, row, "H");
            ExcelHelper.UpdateCell(worksheet, elem.Dimension, row, "I");
            ExcelHelper.UpdateCell(worksheet, elem.TotalLength.ToString(CultureInfo.InvariantCulture), row, "J");
            ExcelHelper.UpdateCell(worksheet, elem.MouldedLength.ToString(CultureInfo.InvariantCulture), row, "K");
            ExcelHelper.UpdateCell(worksheet, elem.Weight.ToString(CultureInfo.InvariantCulture), row, "L");
        }
        
        try
        {
            doc.SaveAs($"{Settings.Default.WorkingDir}\\{Settings.Default.Drawing} - Перечень деталей.xlsx");
            bw.ReportProgress(0, $"{Settings.Default.Drawing} - Перечень деталей.xlsx cоздан\r\n");
        }
        catch (Exception)
        {
            bw.ReportProgress(0, $"Не получилось сохранить {Settings.Default.Drawing} - Перечень деталей.xlsx\r\n");
        }
        finally
        {
            doc.Close();
        }
    }
}
