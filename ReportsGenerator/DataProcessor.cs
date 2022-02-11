using System.ComponentModel;
using System.Text.RegularExpressions;
using ReportsGenerator.Properties;

namespace ReportsGenerator;

internal static class DataProcessor
{
    public static void GenerateAll()
    {
        var files = Directory.GetFiles(Settings.Default.WorkFolder);

        if (files.Length == 0)
        {
            MessageBox.Show("В выбранной папке отсутствуют файлы!");
            return;
        }


        if (!File.Exists(Settings.Default.QualityList))
        {
            MessageBox.Show("Не удается получить данные о плотностях материалов! Проверьте существование и правильное форматирование файла sbh_quality_list.def");
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
            MessageBox.Show("Файл wcog не обнаружен");
            return;
        }

        if (string.IsNullOrEmpty(docxFile))
        {
            MessageBox.Show("Файл спецификации не обнаружен");
            return;
        }

        var wcog = Wcog.Read(wcogFile);
        if (wcog.Count == 0)
        {
            MessageBox.Show("Из wcog'а ничего не прочитано");
            return;
        }

        var docx = Docx.Read(docxFile);
        if (docx.Count == 0)
        {
            MessageBox.Show("Из спецификации ничего не прочитано");
            return;
        }

        ComparisonLog.Gen(wcog, docx);
        PickingList.Gen(wcog);

        var bentParts = wcog.Where(x => x.Value.IsBent).ToDictionary(x => x.Key, x => x.Value);
        if (bentParts.Count > 0)
        {
            BendingList.Gen(bentParts);
        }

        if (genFiles.Count == 0)
        {
            MessageBox.Show("Файлы GEN не обнаружены в заданном расположении, генерация ведомости карт раскроя и материальной ведомости невозможна!");
            return;
        }

        var gens = Gen.Read(genFiles, qualityList);

        NestingList.Gen(gens);
        PlatePivot.Gen(gens);
        ProfilePivot.Gen(wcog);

        MessageBox.Show("Работа завершена");
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