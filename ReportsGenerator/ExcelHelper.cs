using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReportsGenerator;

public static class ExcelHelper
{
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

            var sheet = new Sheet {Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Лист 1"};
            sheets.Append(sheet);
        }

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        
        if (sheetData != null)
        {
            for (var i = 0; i < items.Count; i++)
            {
                for (var j = 0; j < items[i].Length; j++)
                {
                    UpdateCell(worksheetPart, items[i][j], i, j + 1);
                }
            }
        }
        spreadsheetDocument.Close();
    }

    public static void UpdateCell(WorksheetPart worksheetPart, string text, int rowIndex, int columnIndex)
    {
        var columnName = $"{(char)(columnIndex + 64)}";
        UpdateCell(worksheetPart, text,rowIndex, columnName);
    }

    public static void UpdateCell(WorksheetPart worksheetPart, string text, int rowIndex, string columnName)
    {
        if (worksheetPart != null)
        {
            var cell = GetCell(worksheetPart.Worksheet, columnName, (uint)rowIndex);
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            // Save the worksheet.
            worksheetPart.Worksheet.Save();
        }
    }

    static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
    {
        //Get the relationship id of the sheetname
        string relId = workbookPart.Workbook
            .Descendants<Sheet>()
            .First(s => s.Name.Value.Equals(sheetName))
            .Id;

        return (WorksheetPart)workbookPart.GetPartById(relId);
    }

    public static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
    {
        var sheets = document.WorkbookPart?.Workbook.GetFirstChild<Sheets>()
            ?.
            Elements<Sheet>().Where(s => s.Name == sheetName);
        if (!sheets.Any())
        {
            return null;
        }
        var relationshipId = sheets.First().Id.Value;
        var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
        return worksheetPart;
    }

    private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
    {
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        
        var cellReference = columnName + rowIndex;
        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        if (row.Elements<Cell>().Any(c => c.CellReference.Value == cellReference))
        {
            return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
        }

        /*Cell refCell = null;
        foreach (var cell in row.Elements<Cell>())
        {
            if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
            {
                refCell = cell;
                break;
            }
        }
        var newCell = new Cell()
        {
            CellReference = cellReference,
            StyleIndex = (UInt32Value)1U

        };
        row.InsertBefore(newCell, refCell);*/

        // var refCell = row.Descendants<Cell>().LastOrDefault();
        //
        // var newCell = new Cell { CellReference = cellReference };
        // row.InsertAfter(newCell, refCell);
        //
        // worksheet.Save();
        // return newCell;

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }
}