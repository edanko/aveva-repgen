﻿using ReportsGenerator.Properties;

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

        var densityList = DensityList.Read(Settings.Default.QualityList);
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
        
        var profiles = Profiles.Read("profiles.csv");
        var materials = Materials.Read("materials.csv");
        var docx = Docx.Read(docxFile, materials, profiles);
        if (docx.Count == 0)
        {
            MessageBox.Show("Из спецификации ничего не прочитано");
            return;
        }

        if (genFiles.Count == 0)
        {
            MessageBox.Show("Файлы GEN не обнаружены в заданном расположении, генерация ведомости карт раскроя и материальной ведомости невозможна!");
            return;
        }
        var gens = Gen.Read(genFiles, densityList);
        gens.Sort((a, b) => String.Compare(a.NestName, b.NestName, StringComparison.Ordinal));

        ComparisonLog.Gen(wcog, docx);
        PickingList.Gen(wcog, gens);

        var bentParts = wcog.Where(x => x.Value.IsBent).ToDictionary(x => x.Key, x => x.Value);
        if (bentParts.Count > 0)
        {
            BendingList.Gen(bentParts);
        }

        NestingList.Gen(gens);
        PlatePivot.Gen(gens, densityList);
        ProfilePivot.Gen(wcog, profiles);

        MessageBox.Show("Работа завершена");
    }
}