using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelInteraction
{
    public class ExcelDocument
    {
        public ExcelDocument(string fileName)
        {
            CreateSpreadSheetWorkBook(fileName);
        }


        /// <summary>
        /// Создает книгу Excel
        /// </summary>
        /// <param name="fileName"></param>
        public static void CreateSpreadSheetWorkBook(string fileName)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            
            SetPackageProperties(document);

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
        
        public static void InsertText(string fileName, string text, string sheetName, string columnName, uint rowIndex)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fileName, true))
            {
                SharedStringTablePart sharedStringPart;
                if (spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                    sharedStringPart = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else sharedStringPart = spreadsheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                int index = InsertSharedStringItem(text, sharedStringPart);

                //WorksheetPart worksheetPart = InsertWorkSheet(spreadsheet.WorkbookPart);
                //WorksheetPart worksheetPart = spreadsheet.WorkbookPart.GetPartsOfType<WorksheetPart>().First();
                WorksheetPart worksheetPart = null;
                int sheetIndex = 0;
                foreach (WorksheetPart part in spreadsheet.WorkbookPart.WorksheetParts)
                {
                    Worksheet worksheet = part.Worksheet;
                    string name = spreadsheet.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;
                    if (name == sheetName)
                    {
                        worksheetPart = spreadsheet.WorkbookPart.GetPartsOfType<WorksheetPart>().ElementAt(sheetIndex);
                        break;
                    }
                    sheetIndex++;
                }

                if (worksheetPart == null)
                {
                    worksheetPart = InsertWorkSheet(spreadsheet.WorkbookPart, sheetName);
                }

                Cell cell = InsertCellInWorkSheet(columnName, rowIndex, worksheetPart);

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
            if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            else
            {
                row = new Row() {RowIndex = rowIndex};
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
                return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
            else
            {
                Cell refCell = row.Elements<Cell>().FirstOrDefault(cell => string.Compare(cell.CellReference.Value, cellReference, true) > 0);

                Cell newCell = new Cell() {CellReference = cellReference};
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static WorksheetPart InsertWorkSheet(WorkbookPart workbookPart, string sheetName)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            
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

        /// <summary>
        /// Вставляет текст?
        /// </summary>
        /// <param name="text"></param>
        /// <param name="sharedStringPart"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Задает основные свойства докуменат Excel
        /// </summary>
        /// <param name="document"></param>
        private static void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2016-10-24T20:16:51Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "BitLLC";
        }
    }
}
