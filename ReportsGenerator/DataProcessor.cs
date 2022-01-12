using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using ReportsGenerator.My;

namespace ReportsGenerator;

internal static class DataProcessor
{
    public static void GenerateAll(BackgroundWorker bw)
    {
        var files = Directory.GetFiles(MySettingsProperty.Settings.WorkDir);

        if (files.Length == 0)
        {
            bw.ReportProgress(0, "В выбранной папке отсутствуют любые файлы!\r\n");
            return;
        }

        bw.ReportProgress(0, "Начало работы...\r\n");
        var wcogFile = "";
        var partlistFile = "";
        var docFile = "";
        var gens = new List<string>();

        foreach (var f in files)
        {
            var file = Path.GetFileName(f.ToLower());
            if (file.StartsWith("wcog1") && file.EndsWith(".csv"))
                wcogFile = f;
            else if (file.StartsWith("partlist1") && file.EndsWith(".csv"))
                partlistFile = f;
            else if (file.EndsWith(".docx") && !file.StartsWith("~"))
                docFile = f;
            else if (file.EndsWith(".gen")) gens.Add(f);
        }

        var partlist = new ArrayList();
        if (string.IsNullOrEmpty(wcogFile) || string.IsNullOrEmpty(partlistFile))
            bw.ReportProgress(0,
                "Файлы wcog или partlist1 не обнаружены в заданном расположении, генерация перечня деталей и ведомости гибки невозможна!\r\n");
        else
            partlist = PartList.PartlistGen(bw, wcogFile, partlistFile, docFile);

        object nestlist = null;
        if (gens.Count == 0)
            bw.ReportProgress(0,
                "Файлы GEN не обнаружены в заданном расположении, генерация ведомости карт раскроя и материальной ведомости невозможна!\r\n");
        else
            nestlist = NestList.NestlistGen(bw, gens);

        if (partlist.Count == 0 || nestlist == null)
            bw.ReportProgress(0, "Генерация материальной ведомости невозможна, не достаточно данных!\r\n");
        else
            NestList.MaterialListGen(bw, partlist, (Array) nestlist);

        bw.ReportProgress(0, "Работа завершена");
    }

    public static string Regexp(string s, string exp)
    {
        var regex = new Regex(exp);
        var matchCollection = regex.Matches(s);
        string result;
        if (matchCollection.Count > 1)
            result = matchCollection[1].Value;
        else if (matchCollection.Count == 1)
            result = matchCollection[0].Value;
        else
            result = s;
        return result;
    }
}