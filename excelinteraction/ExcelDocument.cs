using System;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using BottomBorder = DocumentFormat.OpenXml.Spreadsheet.BottomBorder;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using FontCharSet = DocumentFormat.OpenXml.Spreadsheet.FontCharSet;
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using LeftBorder = DocumentFormat.OpenXml.Spreadsheet.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Spreadsheet.RightBorder;
using TopBorder = DocumentFormat.OpenXml.Spreadsheet.TopBorder;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

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
            SetPackageProperties();

            ExtendedFilePropertiesPart extendedPropertiesPart = _document.AddExtendedFilePropertiesPart();
            GenerateExtendedProperties(extendedPropertiesPart);

            _workbookPart = _document.AddWorkbookPart();

            WorkbookStylesPart workbookStyles = _workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateStyles(workbookStyles);

            _workbookPart.Workbook = new Workbook();
            _sheets = _workbookPart.Workbook.AppendChild(new Sheets());
        }

        private void GenerateStyles(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "x14ac"}
            };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            #region Fonts
            Fonts fonts = new Fonts()
            {
                Count = 1U,
                KnownFonts = true
            };
            Font font = new Font();
            FontSize fontSize = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = 1U };
            FontName fontName = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet = new FontCharSet() { Val = 204 };
            FontScheme fontScheme = new FontScheme() { Val = FontSchemeValues.Minor };

            font.Append(fontSize);
            font.Append(color1);
            font.Append(fontName);
            font.Append(fontFamilyNumbering);
            font.Append(fontCharSet);
            font.Append(fontScheme);

            fonts.Append(font); 
            #endregion

            #region Fills
            Fills fills = new Fills() { Count = 2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            fills.Append(fill1);
            fills.Append(fill2);
            #endregion

            #region Borders

            //Borders borders = new Borders() {Count = 2U};
            Borders borders = new Borders() {Count = 1U};

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder);
            border1.Append(diagonalBorder1);

            //Border border2 = new Border();
            //LeftBorder leftBorder2 = new LeftBorder() {Style = BorderStyleValues.Medium};
            //Color color2 = new Color() {Indexed = 64U};
            //leftBorder2.Append(color2);
            //RightBorder rightBorder2 = new RightBorder() {Style = BorderStyleValues.Medium};
            //Color color3 = new Color() {Indexed = 64U};
            //rightBorder2.Append(color3);
            //TopBorder topBorder2 = new TopBorder() {Style = BorderStyleValues.Medium};
            //Color color4 = new Color() {Indexed = 64U};
            //topBorder2.Append(color4);
            //BottomBorder bottomBorder2 = new BottomBorder() {Style = BorderStyleValues.Medium};
            //Color color5 = new Color() {Indexed = 64U};
            //bottomBorder2.Append(color5);
            //DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            //border2.Append(leftBorder2);
            //border2.Append(rightBorder2);
            //border2.Append(topBorder2);
            //border2.Append(bottomBorder2);
            //border2.Append(diagonalBorder2);

            borders.Append(border1);
            //borders.Append(border2);
            #endregion

            #region Cell Styles And Formats

            CellStyleFormats cellStyleFormats = new CellStyleFormats() {Count = 1U};
            CellFormat cellFormat1 = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            };
            cellStyleFormats.Append(cellFormat1);

            //CellFormats cellFormats = new CellFormats() {Count = 2U};
            CellFormats cellFormats = new CellFormats() {Count = 1U};
            CellFormat cellFormat2 = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };
            //CellFormat cellFormat3 = new CellFormat()
            //{
            //    NumberFormatId = 0U,
            //    FontId = 0U,
            //    FillId = 0U,
            //    BorderId = 1U,
            //    FormatId = 0U,
            //    ApplyBorder = true
            //};

            cellFormats.Append(cellFormat2);
            //cellFormats.Append(cellFormat3);

            CellStyles cellStyles = new CellStyles() {Count = 1U};
            CellStyle cellStyle = new CellStyle()
            {
                Name = "Обычный",
                FormatId = 0U,
                BuiltinId = 0U
            };
            cellStyles.Append(cellStyle);
            #endregion

            DifferentialFormats differentialFormats = new DifferentialFormats() {Count = 0U};
            TableStyles tableStyles = new TableStyles()
            {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };

            #region Stylesheet Extensions
            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension()
            {
                Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
            };
            stylesheetExtension1.AddNamespaceDeclaration("x14",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };
            stylesheetExtension1.Append(slicerStyles);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension()
            {
                Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}"
            };
            stylesheetExtension2.AddNamespaceDeclaration("x15",
                "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            OpenXmlUnknownElement unknownElement =
                OpenXmlUnknownElement.CreateOpenXmlUnknownElement(
                    "<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />");
            stylesheetExtension2.Append(unknownElement);

            stylesheetExtensionList.Append(stylesheetExtension1);
            stylesheetExtensionList.Append(stylesheetExtension2);
            #endregion

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(stylesheetExtensionList);

            workbookStylesPart.Stylesheet = stylesheet;
        }

        private void SetPackageProperties()
        {
            _document.PackageProperties.Creator = "Phoenix";
            _document.PackageProperties.Created = XmlConvert.ToDateTime(DateTime.Now.ToString("O"),
                XmlDateTimeSerializationMode.RoundtripKind);
            _document.PackageProperties.Modified = XmlConvert.ToDateTime(DateTime.Now.ToString("O"),
                XmlDateTimeSerializationMode.RoundtripKind);
            _document.PackageProperties.LastModifiedBy = "Phoenix";
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
                Size = 2U
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

            Ap.TitlesOfParts titlesOfParts = new Ap.TitlesOfParts();

            Vt.VTVector vtVector2 = new Vt.VTVector()
            {
                BaseType = Vt.VectorBaseValues.Lpstr,
                Size = 1U
            };
            Vt.VTLPSTR vtlpstr2 = new Vt.VTLPSTR() {Text = "testSheet"};
            vtVector2.Append(vtlpstr2);

            titlesOfParts.Append(vtVector2);
            Ap.Company company = new Ap.Company() {Text = "BIT LLC"};
            Ap.LinksUpToDate linksUpToDate = new Ap.LinksUpToDate() {Text = "false"};
            Ap.SharedDocument sharedDocument = new Ap.SharedDocument() {Text = "false"};
            Ap.HyperlinksChanged hyperlinksChanged = new Ap.HyperlinksChanged() {Text = "false"};
            Ap.ApplicationVersion applicationVersion = new Ap.ApplicationVersion() {Text = "15.0300"};

            properties.Append(application);
            properties.Append(documentSecurity);
            properties.Append(scaleCrop);
            properties.Append(headingPairs);
            properties.Append(titlesOfParts);
            properties.Append(company);
            properties.Append(linksUpToDate);
            properties.Append(sharedDocument);
            properties.Append(hyperlinksChanged);
            properties.Append(applicationVersion);

            extendedPropertiesPart.Properties = properties;
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

            GetSpreadSheet(sheetName);

            Cell cell = InsertCellInWorkSheet(columnName, rowIndex);

            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            _worksheetPart.Worksheet.Save();
        }

        private void GetSpreadSheet(string sheetName)
        {
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

        public void SetBorder(string sheetName, string columnName, uint rowIndex, BorderStyleValues thickness)
        {
            GetSpreadSheet(sheetName);

            Cell cell = GetCell(columnName, rowIndex);

            CellFormat cellFormat = cell.StyleIndex != null
                ? GetCellFormat(cell.StyleIndex).CloneNode(true) as CellFormat
                : new CellFormat();

            cellFormat.BorderId = InsertBorder(GenerateBorder(thickness));

            cell.StyleIndex = InsertCellFormat(cellFormat);
        }

        public void MakeBold(string sheetName, string columnName, uint rowIndex)
        {
            GetSpreadSheet(sheetName);
            Cell cell = GetCell(columnName, rowIndex);

            CellFormat cellFormat = cell.StyleIndex != null
                ? GetCellFormat(cell.StyleIndex).CloneNode(true) as CellFormat
                : new CellFormat();

            Font font = new Font();
            Bold bold = new Bold();
            FontSize fontSize = new FontSize() {Val = 11D};
            Color color = new Color() {Theme = 1U};
            FontName name = new FontName() {Val = "Calibri"};
            FontFamilyNumbering numbering = new FontFamilyNumbering() {Val = 2};
            FontScheme scheme = new FontScheme() {Val = FontSchemeValues.Minor};

            font.Append(bold);
            font.Append(fontSize);
            font.Append(color);
            font.Append(name);
            font.Append(numbering);
            font.Append(scheme);

            //Fonts fonts = _workbookPart.WorkbookStylesPart.Stylesheet.Fonts;
            //fonts.Append(font);
            //fonts.Count.Value++;

            cellFormat.FontId = InsertFont(font);
            cellFormat.ApplyFont = true;
            //_workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(cellFormat);
            //_workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count.Value++;
            //_workbookPart.WorkbookStylesPart.Stylesheet.Save();

            cell.StyleIndex = InsertCellFormat(cellFormat);
            //cell.StyleIndex = InsertFont(font);
        }

        private uint InsertFont(Font font)
        {
            Fonts fonts = _workbookPart.WorkbookStylesPart.Stylesheet.Fonts;
            fonts.Append(font);
            return fonts.Count++;
        }

        private uint InsertCellFormat(CellFormat cellFormat)
        {
            CellFormats cellFormats = _workbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
            cellFormats.Append(cellFormat);
            return cellFormats.Count++;
        }

        private Cell GetCell(string columnName, uint rowIndex)
        {
            var cellAddress = $"{columnName}{rowIndex}";
            return _worksheetPart.Worksheet.Descendants<Cell>()
                .SingleOrDefault(c => cellAddress.Equals(c.CellReference));
        }

        private CellFormat GetCellFormat(uint styleIndex)
        {
            return
                _workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>()
                    .First()
                    .Elements<CellFormat>()
                    .ElementAt((int) styleIndex);
        }

        private UInt32Value InsertBorder(Border border)
        {
            Borders borders = _workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            return borders.Count++;
        }

        private Border GenerateBorder(BorderStyleValues thickness)
        {
            Border border = new Border();

            LeftBorder leftBorder = new LeftBorder() { Style = thickness };
            Color color1 = new Color() { Indexed = 64U };
            leftBorder.Append(color1);

            RightBorder rightBorder = new RightBorder() { Style = thickness };
            Color color2 = new Color() { Indexed = 64U };
            rightBorder.Append(color2);

            TopBorder topBorder = new TopBorder() { Style = thickness };
            Color color3 = new Color() { Indexed = 64U };
            topBorder.Append(color3);

            BottomBorder bottomBorder = new BottomBorder() { Style = thickness };
            Color color4 = new Color() { Indexed = 64U };
            bottomBorder.Append(color4);

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);

            return border;
        }
    }
}