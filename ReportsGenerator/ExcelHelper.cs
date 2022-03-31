using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace ReportsGenerator;

public static class ExcelHelper
{
    private static Cell InsertCell(Row row, int columnIndex, uint styleIndex)
    {
        var cellReference = $"{(char) (columnIndex + 64)}{row.RowIndex}";

        var refCell = row.Descendants<Cell>().LastOrDefault();

        var newCell = new Cell {CellReference = cellReference, StyleIndex = styleIndex};
        row.InsertAfter(newCell, refCell);

        return newCell;
    }

    public static void CreateXlsx(string filepath, List<string[]> items)
    {
        var spreadsheetDocument =
            SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

        var workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        var stylePart = workbookpart.AddNewPart<WorkbookStylesPart>();
        stylePart.Stylesheet = GenerateStyleSheet();
        stylePart.Stylesheet.Save();
        
        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet();
        
        if (spreadsheetDocument.WorkbookPart != null)
        {
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
                {Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1"};
            sheets.Append(sheet);
        }

        var sheetData = new SheetData();

        for (var i = 0; i < items.Count; i++)
        {
            var row = new Row {RowIndex = (uint) i + 1};
            sheetData.Append(row);

            var styleIndex = i >= 2 ? (uint)2 : 0;
            
            for (var j = 0; j < items[i].Length; j++)
            {
                var cell = InsertCell(row, j + 1, styleIndex);
                cell.CellValue = new CellValue(items[i][j]);

                var isNumber = double.TryParse(items[i][j], NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out _);

                if (isNumber)
                {
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                }
                else
                {
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
        }

        var columns = AutoSize(sheetData);
        worksheetPart.Worksheet.Append(columns);
        worksheetPart.Worksheet.Append(sheetData);

        spreadsheetDocument.Close();
    }

    private static Columns AutoSize(SheetData sheetData)
    {
        var maxColWidth = GetMaxCharacterWidth(sheetData);

        Columns columns = new Columns();
        //this is the width of my font - yours may be different
        double maxWidth = 7;
        foreach (var item in maxColWidth)
        {
            //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

            //pixels=Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
            double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);

            //character width=Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
            double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

            Column col = new Column()
            {
                BestFit = true, Min = (UInt32) (item.Key + 1), Max = (UInt32) (item.Key + 1), CustomWidth = true,
                Width = (DoubleValue) width
            };
            columns.Append(col);
        }

        return columns;
    }

    private static Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
    {
        //iterate over all cells getting a max char value for each column
        Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
        var rows = sheetData.Elements<Row>().ToList();
        UInt32[] numberStyles = {5, 6, 7, 8}; //styles that will add extra chars
        UInt32[] boldStyles = {1, 2, 3, 4, 6, 7, 8}; //styles that will bold
        for (int j = 2; j < rows.Count(); j++)
        {
            var r = rows[j];
        // }
        // foreach (var r in rows)
        // {
            var cells = r.Elements<Cell>().ToArray();

            //using cell index as my column
            for (int i = 0; i < cells.Length; i++)
            {
                var cell = cells[i];
                var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                var cellTextLength = cellValue.Length;

                if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                {
                    int thousandCount = (int) Math.Truncate((double) cellTextLength / 4);

                    //add 3 for '.00' 
                    cellTextLength += (3 + thousandCount);
                }

                if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                {
                    //add an extra char for bold - not 100% acurate but good enough for what i need.
                    cellTextLength += 1;
                }

                if (maxColWidth.ContainsKey(i))
                {
                    var current = maxColWidth[i];
                    if (cellTextLength > current)
                    {
                        maxColWidth[i] = cellTextLength;
                    }
                }
                else
                {
                    maxColWidth.Add(i, cellTextLength);
                }
            }
        }

        return maxColWidth;
    }
    
     //Метод генерирует стили для ячеек (за основу взят код, найденный где-то в интернете)
        static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Стиль под номером 0 - Шрифт по умолчанию.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Стиль под номером 1 - Жирный шрифт Times New Roman.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                    new Font(                                                               // Стиль под номером 2 - Обычный шрифт Times New Roman.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" }),
                    new Font(                                                               // Стиль под номером 3 - Шрифт Times New Roman размером 14.
                        new FontSize() { Val = 14 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(                                                           // Стиль под номером 0 - Заполнение ячейки по умолчанию.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Стиль под номером 1 - Заполнение ячейки серым цветом
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFAAAAAA" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill(                                                           // Стиль под номером 2 - Заполнение ячейки красным.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFAAAA" } }
                        )
                        { PatternType = PatternValues.Solid })
                )
                ,
                new Borders(
                    new Border(                                                         // Стиль под номером 0 - Грани.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Стиль под номером 1 - Грани
                        new LeftBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Medium },
                        new RightBorder(
                            new Color() { Indexed = (UInt32Value)64U }
                        )
                        { Style = BorderStyleValues.Medium },
                        new TopBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Medium },
                        new BottomBorder(
                            new Color() { Indexed = (UInt32Value)64U }
                        )
                        { Style = BorderStyleValues.Medium },
                        new DiagonalBorder()),
                    new Border(                                                         // Стиль под номером 2 - Грани.
                        new LeftBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color() { Indexed = (UInt32Value)64U }
                        )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color() { Auto = true }
                        )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color() { Indexed = (UInt32Value)64U }
                        )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                          // Стиль под номером 0 - The default cell style.  (по умолчанию)
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true },       // Стиль под номером 1 - Bold 
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 2, FillId = 0, BorderId = 2, ApplyFont = true },       // Стиль под номером 2 - REgular
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 2, ApplyFont = true, NumberFormatId = 4 },       // Стиль под номером 3 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Стиль под номером 4 - Yellow Fill
                    new CellFormat(                                                                   // Стиль под номером 5 - Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true },      // Стиль под номером 6 - Border
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center, WrapText = true }) { FontId = 2, FillId = 0, BorderId = 2, ApplyFont = true, NumberFormatId = 4 }       // Стиль под номером 7 - Задает числовой формат полю.
                )
            ); // Выход
        }


}