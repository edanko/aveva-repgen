using System.ComponentModel;
using System.Globalization;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class PickingList
{
    public static void Gen(BackgroundWorker bw, Dictionary<string, Wcog> wcog, Dictionary<string, Docx> docx)
    {
        // TODO: move checking docx vs wcog to separate file
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
        
        var items = new List<string[]>
        {
            new []
            {
                "№ п/п",
                "Секция",
                "Позиция",
                "Кол-во",
                "Толщина",
                "Карта раскроя",
                "Shape",
                "Dimension",
                "Total Length",
                "Moulded Length",
                "Масса",
                "Маршрут",
                "Сборка",
                "АРЭ"
            }
        };


        var list = wcogKeys.ToList();
        list.Sort();

        for (var i = 0; i < list.Count; i++)
        {
            var key = list[i];
            var elem = wcog[key];

            items.Add(new[]
            {
                (i+1).ToString(),
                elem.Block,
                elem.PosNo,
                elem.Quantity.ToString(),
                elem.GetThickness().ToString(CultureInfo.InvariantCulture),
                elem.Quality,
                elem.NestedOn,
                elem.Shape,
                elem.Dimension,
                elem.TotalLength.ToString(CultureInfo.InvariantCulture),
                elem.MouldedLength.ToString(CultureInfo.InvariantCulture),
                elem.Weight.ToString(CultureInfo.InvariantCulture),
            });
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkingDir}\\{Settings.Default.Drawing} - Перечень деталей.xlsx", items);
    }
}
