using System.ComponentModel;
using System.Text.RegularExpressions;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

internal static class DataProcessor
{
    public static void GenerateAll(BackgroundWorker bw)
    {
        var files = Directory.GetFiles(Settings.Default.WorkFolder);

        if (files.Length == 0)
        {
            bw.ReportProgress(0, "В выбранной папке отсутствуют файлы!\r\n");
            return;
        }

        bw.ReportProgress(0, "Начало работы...\r\n");


        if (!File.Exists(Settings.Default.QualityList))
        {
            bw.ReportProgress(0,
                "Не удается получить данные о плотностях материалов! Проверьте существование и правильное форматирование файла sbh_quality_list.def. Нужно выбрать файл в настройках\r\n");
            return;
        }

        var qualityList = QualityList.Read(Settings.Default.QualityList);
        var wcogFile = "";
        var docxFile = "";
        var genFiles = new List<string>();

        foreach (var f in files)
        {
            var file = Path.GetFileName(f.ToLower());
            if (file.StartsWith("wcog1") && file.EndsWith(".csv"))
            {
                wcogFile = f;
            }
            else if (file.EndsWith(".docx") && !file.StartsWith("~"))
            {
                docxFile = f;
            }
            else if (file.EndsWith(".gen"))
            {
                genFiles.Add(f);
            }
        }

        if (string.IsNullOrEmpty(wcogFile))
        {
            bw.ReportProgress(0, "Файл wcog не обнаружен\r\n");
            return;
        }

        if (string.IsNullOrEmpty(docxFile))
        {
            bw.ReportProgress(0, "Файл спецификации не обнаружен\r\n");
            return;
        }

        var wcog = Wcog.Read(wcogFile);
        if (wcog.Count == 0)
        {
            bw.ReportProgress(0, "Из wcog'а ничего не прочитано\r\n");
            return;
        }

        var docx = Docx.Read(bw, docxFile);
        if (docx.Count == 0)
        {
            bw.ReportProgress(0, "Из спецификации ничего не прочитано\r\n");
            return;
        }

        PickingList.Gen(bw, wcog, docx);

        var bentParts = wcog.Where(x => x.Value.IsBent).ToDictionary(x => x.Key, x => x.Value);
        if (bentParts.Count == 0)
        {
            bw.ReportProgress(0, "Гнутые детали не обнаружены!\r\n");
        }
        else
        { 
            bw.ReportProgress(0, $"Гнутых деталей найдено: {bentParts.Count}\r\n");
            BendingList.Gen(bentParts);
        }

        if (genFiles.Count == 0)
        {
            bw.ReportProgress(0, "Файлы GEN не обнаружены в заданном расположении, генерация ведомости карт раскроя и материальной ведомости невозможна!\r\n");
            return;
        }

        var gens = Gen.Read(genFiles, qualityList);

        NestingList.Gen(gens);
        MaterialList.Gen(wcog, gens);

        bw.ReportProgress(0, "Работа завершена\r\n");
    }

    public static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var matchCollection = regex.Matches(s);
        var result = matchCollection.Count switch
        {
            > 1 => matchCollection[1].Value,
            1 => matchCollection[0].Value,
            _ => s
        };
        return result;
    }
}