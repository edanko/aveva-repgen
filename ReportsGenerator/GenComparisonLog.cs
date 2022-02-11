﻿using ReportsGenerator.Properties;

namespace ReportsGenerator;

public static class ComparisonLog
{
    public static void Gen(Dictionary<string, Wcog> wcog, Dictionary<string, Docx> docx)
    {
        var items = new List<string[]>
        {
            new[]
            {
                "Позиция",
                "Тип ошибки",
                "WCOG",
                "Спец."
            }
        };

        var wcogKeys = wcog.Keys;
        var docxKeys = docx.Keys;

        var keysOnlyInDocx = docxKeys.Except(wcogKeys);
        foreach (var k in keysOnlyInDocx)
        {
            items.Add(new[]
            {
                k,
                "Отсутствует в WCOG",
            });
        }

        var keysOnlyInWcog = wcogKeys.Except(docxKeys);
        foreach (var k in keysOnlyInWcog)
        {
            items.Add(new[]
            {
                k,
                "Отсутствует в спецификации",
            });
        }

        var common = wcogKeys.Intersect(docxKeys).ToList();
        common.Sort();
        foreach (var k in common)
        {
            if (wcog[k].Quality != docx[k].Quality)
            {
                items.Add(new[]
                {
                    k,
                    "Конфликт материалов",
                    wcog[k].Quality,
                    docx[k].Quality
                });
            }

            if (wcog[k].Dimension != docx[k].Dimension)
            {
                items.Add(new[]
                {
                    k,
                    "Конфликт типоразмеров",
                    wcog[k].Dimension,
                    docx[k].Dimension
                });
            }
        }

        ExcelHelper.CreateXlsx($"{Settings.Default.WorkFolder}\\{Settings.Default.Drawing} - Лог сравнения WCOG и спецификации.xlsx", items);
    }
}