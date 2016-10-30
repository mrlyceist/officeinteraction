using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ExcelInteraction
{
    public class ExcelDocument2
    {
        public ExcelDocument2(string fileName)
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

            WorkbookStylesPart workbookStylesPart1 = workbook.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

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
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                SharedStringTablePart sharedStringPart;
                if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                    sharedStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();

                int index = InsertSharedStringItem(text, sharedStringPart);

                WorksheetPart worksheetPart = null;
                int sheetIndex = 0;
                foreach (WorksheetPart part in document.WorkbookPart.WorksheetParts)
                {
                    Worksheet worksheet = part.Worksheet;
                    string name = document.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;
                    if (name == sheetName)
                    {
                        worksheetPart = document.WorkbookPart.GetPartsOfType<WorksheetPart>().ElementAt(sheetIndex);
                        break;
                    }
                    sheetIndex++;
                }

                if (worksheetPart == null)
                {
                    worksheetPart = InsertWorkSheet(document.WorkbookPart, sheetName);
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

        public static void SetBorder(string fileName, string sheetName, string columnName, uint rowIndex)
        {
            //GenerateBorder(BorderStyleValues.Thick);
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fileName, true))
            {
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

                Cell cell = GetCell(worksheetPart, columnName, rowIndex);

                CellFormat cellFormat = cell.StyleIndex != null
                    ? GetCellFormat(spreadsheet.WorkbookPart, cell.StyleIndex).CloneNode(true) as CellFormat
                    : new CellFormat();
                cellFormat.BorderId = InsertBorder(spreadsheet.WorkbookPart, GenerateBorder(BorderStyleValues.Thick));

                worksheetPart.Worksheet.Save();
            }

        }

        private static uint InsertBorder(WorkbookPart workbookPart, Border border)
        {
            Borders borders = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            return borders.Count++;
        }

        private static CellFormat GetCellFormat(WorkbookPart workbookPart, uint styleIndex)
        {
            return
                workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>()
                    .First()
                    .Elements<CellFormat>()
                    .ElementAt((int) styleIndex);
        }

        private static Cell GetCell(WorksheetPart worksheetPart, string columnName, uint rowIndex)
        {
            var cellAddres = $"{columnName}{rowIndex}";
            return worksheetPart.Worksheet.Descendants<Cell>().SingleOrDefault(c => cellAddres.Equals(c.CellReference));
        }

        private static Border GenerateBorder(BorderStyleValues borderStyle)
        {
            Border border = new Border();

            LeftBorder leftBorder = new LeftBorder() { Style = borderStyle };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };
            leftBorder.Append(color1);

            RightBorder rightBorder = new RightBorder() { Style = borderStyle };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };
            rightBorder.Append(color2);

            TopBorder topBorder = new TopBorder() { Style = borderStyle };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };
            topBorder.Append(color3);

            BottomBorder bottomBorder = new BottomBorder() { Style = borderStyle };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };
            bottomBorder.Append(color4);

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);

            return border;
        }

        private static void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color2);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color3);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color4);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color5);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />");

            stylesheetExtension2.Append(openXmlUnknownElement3);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
    }
}
