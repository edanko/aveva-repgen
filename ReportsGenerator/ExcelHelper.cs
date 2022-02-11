using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReportsGenerator;

public static class ExcelHelper
{
    private static Cell InsertCell(Row row, int columnIndex)
    {
        var cellReference = $"{(char)(columnIndex + 64)}{row.RowIndex}";

        var refCell = row.Descendants<Cell>().LastOrDefault();

        var newCell = new Cell { CellReference = cellReference };
        row.InsertAfter(newCell, refCell);

        return newCell;
    }

    public static void CreateXlsx(string filepath, List<string[]> items)
    {
        var spreadsheetDocument =
            SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

        var workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        if (spreadsheetDocument.WorkbookPart != null)
        {
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet {Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1"};
            sheets.Append(sheet);
        }

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        
        if (sheetData != null)
        {
            for (var i = 0; i < items.Count; i++)
            {
                var row = new Row { RowIndex = (uint)i + 1 };
                sheetData.Append(row);

                for (var j = 0; j < items[i].Length; j++)
                {
                    var cell = InsertCell(row, j + 1);
                    cell.CellValue = new CellValue(items[i][j]);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
        }
        spreadsheetDocument.Close();
    }
}