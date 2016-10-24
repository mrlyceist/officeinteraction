using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelInteraction
{
    public class ExcelDocument
    {
        public ExcelDocument(string fileName)
        {
            CreateSpreadSheetWorkBook(fileName);
        }

        private static void CreateSpreadSheetWorkBook(string fileName)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Sheet sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheet),
                SheetId = 1,
                Name = "report"
            };

            sheets.Append(sheet);

            workbook.Workbook.Save();
            document.Close();
        }

        //public void CreatePackage(string filePath)
        //{
        //    using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        //    {
        //        CreateParts(package);
        //    }
        //}


        //public void Save(SpreadsheetDocument document)
        //{
        //    throw new NotImplementedException();
        //}


        public static void InsertText(string fileName, string text)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fileName, true))
            {
                SharedStringTablePart sharedStringPart;
                if (spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    sharedStringPart = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    sharedStringPart = spreadsheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                int index = InsertSharedStringItem(text, sharedStringPart);

                WorksheetPart worksheetPart = InsertWorkSheet(spreadsheet.WorkbookPart);

                Cell cell = InsertCellInWorkSheet("A", 1, worksheetPart);

                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                worksheetPart.Worksheet.Save();
            }
        }

        private static Cell InsertCellInWorkSheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            else
            {
                row = new Row() {RowIndex = rowIndex};
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true)>0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() {CellReference = cellReference};
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static WorksheetPart InsertWorkSheet(WorkbookPart workbookPart)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            string sheetName = $"Sheet {sheetId}";
            Sheet sheet = new Sheet()
            {
                Id = relationshipId,
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart sharedStringPart)
        {
            if (sharedStringPart.SharedStringTable == null)
                sharedStringPart.SharedStringTable = new SharedStringTable();

            int i = 0;

            foreach (SharedStringItem item in sharedStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text) return i;
                i++;
            }

            sharedStringPart.SharedStringTable.AppendChild(
                new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            sharedStringPart.SharedStringTable.Save();

            return i;
        }
    }
}
