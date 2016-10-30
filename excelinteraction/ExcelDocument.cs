using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

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

            //ExtendedFilePropertiesPart extendedPropertiesPart = _document.AddExtendedFilePropertiesPart();
            //GenerateExtendedProperties(extendedPropertiesPart);

            //WorkbookStylesPart _workbookStyles = _workbook.AddNewPart<WorkbookStylesPart>("rId3");
            //GenerateStyles(_workbookStyles);

            _workbookPart = _document.AddWorkbookPart();

            _workbookPart.Workbook = new Workbook();
            //_worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            //_worksheetPart.Worksheet = new Worksheet(new SheetData());

            //_sheets = _document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            _sheets = _workbookPart.Workbook.AppendChild(new Sheets());
        }

        private void GenerateExtendedProperties(ExtendedFilePropertiesPart extendedPropertiesPart)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt",
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application = new Ap.Application {Text = "Microsoft Excel"};
            Ap.DocumentSecurity documentSecurity = new Ap.DocumentSecurity {Text = "0"};
            Ap.ScaleCrop scaleCrop = new Ap.ScaleCrop() {Text = "false"};

            Ap.HeadingPairs headingPairs = new Ap.HeadingPairs();

            Vt.VTVector vtVector = new Vt.VTVector()
            {
                BaseType = Vt.VectorBaseValues.Variant,
                Size = (UInt32Value) 2U
            };
            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vtlpstr = new Vt.VTLPSTR() {Text = "Spreadsheets"};
            variant1.Append(vtlpstr);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vtInt32 = new Vt.VTInt32() {Text = "1"};
            variant2.Append(vtInt32);

            vtVector.Append(variant1);
            vtVector.Append(variant2);

            headingPairs.Append(vtVector);

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

        public void InsertText(string text, string sheetName, string columnName, uint rowIndex)
        {
            SharedStringTablePart sharedStringPart;
            if (_workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                sharedStringPart = _workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();

            int index = InsertSharedString(text, sharedStringPart);

            int sheetIndex = 0;
            foreach (WorksheetPart part in _workbookPart.WorksheetParts)
            {
                Worksheet worksheet = part.Worksheet;
                string name = _workbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;
                if (name == sheetName)
                {
                    _worksheetPart = _workbookPart.GetPartsOfType<WorksheetPart>().ElementAt(sheetIndex);
                    break;
                }
                sheetIndex++;
            }

            if (_worksheetPart == null)
                AddSpreadSheet(sheetName);

            Cell cell = InsertCellInWorkSheet(columnName, rowIndex);

            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            _worksheetPart.Worksheet.Save();
        }

        private Cell InsertCellInWorkSheet(string columnName, uint rowIndex)
        {
            Worksheet worksheet = _worksheetPart.Worksheet;
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

            if (row.Elements<Cell>().Any(c => c.CellReference == cellReference))
                return row.Elements<Cell>().First(c => c.CellReference == cellReference);

            Cell refCell =
                row.Elements<Cell>()
                    .FirstOrDefault(cell => string.Compare(cell.CellReference.Value, cellReference, true) > 0);
            Cell newCell = new Cell() {CellReference = cellReference};
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }

        private int InsertSharedString(string text, SharedStringTablePart sharedStringPart)
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