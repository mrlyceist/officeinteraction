using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelInteraction
{
    public class ExcelDocument
    {
        private SpreadsheetDocument _document;
        private WorkbookPart _workbookPart;
        private WorksheetPart _worksheetPart;
        private Sheets _sheets;

        public ExcelDocument(string fileName)
        {
            CreateSpreadSheetWorkBook(fileName);
        }

        private void CreateSpreadSheetWorkBook(string fileName)
        {
            _document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            //SetPackageProperties(_document);

            //WorkbookStylesPart _workbookStyles = _workbook.AddNewPart<WorkbookStylesPart>("rId3");
            //GenerateStyles(_workbookStyles);

            _workbookPart = _document.AddWorkbookPart();

            _workbookPart.Workbook = new Workbook();
            //_worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            //_worksheetPart.Worksheet = new Worksheet(new SheetData());

            //_sheets = _document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            _sheets = _workbookPart.Workbook.AppendChild(new Sheets());
        }

        public void Save()
        {
            _workbookPart.Workbook.Save();
            _document.Close();
        }

        public void AddSpreadSheet(string sheetName)
        {
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());
            _worksheetPart.Worksheet.Save();

            Sheets sheets = _workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = _workbookPart.GetIdOfPart(_worksheetPart);

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

            _workbookPart.Workbook.Save();
        }
    }
}