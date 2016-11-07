using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

[assembly: InternalsVisibleTo("ExcelInteractionTests")]
namespace ExcelInteraction
{
    /// <summary>
    /// Класс взаимодействия с документами Excel.
    /// </summary>
    public class ExcelDocument
    {
        #region Private Fields
        private SpreadsheetDocument _document;
        private WorkbookPart _workbookPart;
        private WorksheetPart _worksheetPart;
        private Worksheet _workSheet;
        #endregion

        #region Public Properties
        /// <summary>
        /// Объект документа Excel
        /// </summary>
        public SpreadsheetDocument Document
        {
            get { return _document; }
            set { _document = value; }
        }

        /// <summary>
        /// Имя первого по счету листа в книге.
        /// </summary>
        public string FirstSheetName => _workbookPart.Workbook.Descendants<Sheet>().First().Name;
        #endregion

        #region Constructors
        /// <summary>
        /// Создает новый пустой документ Excel. В этот документ необходимо добавть листы!
        /// </summary>
        /// <param name="fileName">Полный путь к документу</param>
        public ExcelDocument(string fileName)
        {
            CreateSpreadSheetWorkBook(fileName);
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Создает заготовку для документа Excel
        /// </summary>
        /// <param name="fileName">Полный путь к файлу</param>
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
            _workbookPart.Workbook.AppendChild(new Sheets());
        }

        /// <summary>
        /// Создает стилевую составляющую документа
        /// </summary>
        /// <param name="workbookStylesPart">Стилевая часть книги</param>
        private void GenerateStyles(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
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
            Borders borders = new Borders() { Count = 1U };

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

            borders.Append(border1);
            #endregion

            #region Cell Styles And Formats
            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = 1U };
            CellFormat cellFormat1 = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            };
            cellStyleFormats.Append(cellFormat1);

            CellFormats cellFormats = new CellFormats() { Count = 1U };
            CellFormat cellFormat2 = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };

            cellFormats.Append(cellFormat2);

            CellStyles cellStyles = new CellStyles() { Count = 1U };
            CellStyle cellStyle = new CellStyle()
            {
                Name = "Обычный",
                FormatId = 0U,
                BuiltinId = 0U
            };
            cellStyles.Append(cellStyle);
            #endregion

            DifferentialFormats differentialFormats = new DifferentialFormats() { Count = 0U };
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

        /// <summary>
        /// Создает основные свойства документа - автора, дату и время создания и модификации
        /// </summary>
        private void SetPackageProperties()
        {
            _document.PackageProperties.Creator = "Phoenix";
            _document.PackageProperties.Created = XmlConvert.ToDateTime(DateTime.Now.ToString("O"),
                XmlDateTimeSerializationMode.RoundtripKind);
            _document.PackageProperties.Modified = XmlConvert.ToDateTime(DateTime.Now.ToString("O"),
                XmlDateTimeSerializationMode.RoundtripKind);
            _document.PackageProperties.LastModifiedBy = "Phoenix";
        }

        /// <summary>
        /// Создает расширенные свойства документа
        /// </summary>
        /// <param name="extendedPropertiesPart"></param>
        private void GenerateExtendedProperties(ExtendedFilePropertiesPart extendedPropertiesPart)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt",
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application = new Ap.Application { Text = "Microsoft Excel" };
            Ap.DocumentSecurity documentSecurity = new Ap.DocumentSecurity { Text = "0" };
            Ap.ScaleCrop scaleCrop = new Ap.ScaleCrop() { Text = "false" };

            Ap.HeadingPairs headingPairs = new Ap.HeadingPairs();

            Vt.VTVector vtVector = new Vt.VTVector()
            {
                BaseType = Vt.VectorBaseValues.Variant,
                Size = 2U
            };
            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vtlpstr = new Vt.VTLPSTR() { Text = "Spreadsheets" };
            variant1.Append(vtlpstr);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vtInt32 = new Vt.VTInt32() { Text = "1" };
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
            Vt.VTLPSTR vtlpstr2 = new Vt.VTLPSTR() { Text = "testSheet" };
            vtVector2.Append(vtlpstr2);

            titlesOfParts.Append(vtVector2);
            Ap.Company company = new Ap.Company() { Text = "BIT LLC" };
            Ap.LinksUpToDate linksUpToDate = new Ap.LinksUpToDate() { Text = "false" };
            Ap.SharedDocument sharedDocument = new Ap.SharedDocument() { Text = "false" };
            Ap.HyperlinksChanged hyperlinksChanged = new Ap.HyperlinksChanged() { Text = "false" };
            Ap.ApplicationVersion applicationVersion = new Ap.ApplicationVersion() { Text = "15.0300" };

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

        /// <summary>
        /// Ищет в документе лист с определенным именем. Если не находит - создает.
        /// </summary>
        /// <param name="sheetName">Имя листа</param>
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
                    _workSheet = _worksheetPart.Worksheet;
                    break;
                }
                sheetIndex++;
            }

            if (_worksheetPart == null)
                AddSpreadSheet(sheetName);
        }

        /// <summary>
        /// Добавляет ячейку на лист
        /// </summary>
        /// <param name="columnName">Адрес ячейки, имя столбца</param>
        /// <param name="rowIndex">Адрес ячейки, номер строки</param>
        /// <returns>Ячейка</returns>
        private Cell InsertCellInWorkSheet(string columnName, uint rowIndex)
        {
            Worksheet worksheet = _workSheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
                row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Any(c => c.CellReference == cellReference))
                return row.Elements<Cell>().First(c => c.CellReference == cellReference);

            Cell refCell =
                row.Elements<Cell>()
                    .FirstOrDefault(cell => string.Compare(cell.CellReference.Value, cellReference, true) > 0);
            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }

        /// <summary>
        /// Добавляет строковое содержимое в структуру документа
        /// </summary>
        /// <param name="text">добавляемая строка</param>
        /// <param name="sharedStringPart">Часть документа, содержащая строки</param>
        /// <returns>Номер строки в части строк</returns>
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

        /// <summary>
        /// Ищет формат ячейки. Если не находит - создает новый.
        /// </summary>
        /// <param name="cell">Ячейка, у которой ищем формат</param>
        /// <returns>Формат ячейки</returns>
        private CellFormat GetOrCreateCellFormat(Cell cell)
        {
            return cell.StyleIndex != null
                            ? GetCellFormat(cell.StyleIndex).CloneNode(true) as CellFormat
                            : new CellFormat();
        }

        /// <summary>
        /// Добавляет шрифт к набору шрифтов в документе
        /// </summary>
        /// <param name="font">Добавляемый шрифт</param>
        /// <returns>Номер последнего по счету шрифта в документе</returns>
        private uint InsertFont(Font font)
        {
            Fonts fonts = _workbookPart.WorkbookStylesPart.Stylesheet.Fonts;
            fonts.Append(font);
            return fonts.Count++;
        }

        /// <summary>
        /// Добавляет формат ячейки к набору форматов в документе
        /// </summary>
        /// <param name="cellFormat">Добавляемый формат ячейки</param>
        /// <returns>Номер последнего по счету формата ячейки в документе</returns>
        private uint InsertCellFormat(CellFormat cellFormat)
        {
            CellFormats cellFormats = _workbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
            cellFormats.Append(cellFormat);
            return cellFormats.Count++;
        }

        /// <summary>
        /// Находит ячейку на листе по заданному адресу.
        /// </summary>
        /// <param name="columnName">Адрес искомой ячейки, имя столбца</param>
        /// <param name="rowIndex">Адрес искомой ячейки, номер строки</param>
        /// <returns>Искомая ячейка</returns>
        private Cell GetCell(string columnName, uint rowIndex)
        {
            var cellAddress = $"{columnName}{rowIndex}";
            CreateCellIfNotExists(columnName, rowIndex);
            return _workSheet.Descendants<Cell>()
                .SingleOrDefault(c => cellAddress.Equals(c.CellReference));
        }

        /// <summary>
        /// Проверяет наличие в листе ячейки с заданным адресом,
        /// Если не находит - создает.
        /// </summary>
        /// <param name="cellColumn">Имя столбца проверяемой ячейки</param>
        /// <param name="cellRow">Номер ряда проверяемой ячейки</param>
        private void CreateCellIfNotExists(string cellColumn, uint cellRow)
        {
            string cellReference = $"{cellColumn}{cellRow}";
            var rows = _workSheet.Descendants<Row>().Where(r => r.RowIndex.Value == cellRow);
            if (!rows.Any())
            {
                Row row = new Row() { RowIndex = new UInt32Value(cellRow) };
                Cell cell = new Cell() { CellReference = new StringValue(cellReference) };
                row.Append(cell);
                _workSheet.Descendants<SheetData>().First().Append(row);
                _workSheet.Save();
            }
            else
            {
                Row row = rows.First();
                var cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
                if (cells.Any()) return;
                Cell cell = new Cell() { CellReference = new StringValue(cellReference) };
                row.Append(cell);
                _workSheet.Save();
            }
        }

        /// <summary>
        /// Ищет формат ячейки по заданному стилю.
        /// </summary>
        /// <param name="styleIndex">Стиль?</param>
        /// <returns>Формат ячейки</returns>
        private CellFormat GetCellFormat(uint styleIndex)
        {
            return
                _workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>()
                    .First()
                    .Elements<CellFormat>()
                    .ElementAt((int)styleIndex);
        }

        #region Borders Section
        /// <summary>
        /// Добавляет границу к набуору границ в документе
        /// </summary>
        /// <param name="border">Добавляемая граница</param>
        /// <returns>Номер последней по счету границы в документе</returns>
        private UInt32Value InsertBorder(Border border)
        {
            Borders borders = _workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            return borders.Count++;
        }

        /// <summary>
        /// Создает границу черного цвета с заданной толщиной линии
        /// </summary>
        /// <param name="thickness">Толщина границы</param>
        /// <param name="addLeftBorder">Задает наличие левой границы</param>
        /// <param name="addRightBorder">Задает наличие правой границы</param>
        /// <param name="addTopBorder">Задает наличие верхней границы</param>
        /// <param name="addBottomBorder">Задает наличие нижней границы</param>
        /// <returns>Граница</returns>
        private Border GenerateBorder(BorderStyleValues thickness, bool addLeftBorder = false,
            bool addRightBorder = false, bool addTopBorder = false, bool addBottomBorder = false)
        {
            Border border = new Border();
            if (addLeftBorder)
            {
                LeftBorder leftBorder = GeneerateLeftBorder(thickness);
                border.Append(leftBorder);
            }
            if (addRightBorder)
            {
                RightBorder rightBorder = GenerateRightBorder(thickness);
                border.Append(rightBorder);
            }
            if (addTopBorder)
            {
                TopBorder topBorder = GenerateTopBorder(thickness);
                border.Append(topBorder);
            }
            if (addBottomBorder)
            {
                BottomBorder bottomBorder = GenerateBottomBorder(thickness);
                border.Append(bottomBorder);
            }

            return border;
        }

        private static BottomBorder GenerateBottomBorder(BorderStyleValues thickness)
        {
            BottomBorder bottomBorder = new BottomBorder() { Style = thickness };
            Color color = new Color() { Indexed = 64U };
            bottomBorder.Append(color);
            return bottomBorder;
        }

        private static TopBorder GenerateTopBorder(BorderStyleValues thickness)
        {
            TopBorder topBorder = new TopBorder() { Style = thickness };
            Color color = new Color() { Indexed = 64U };
            topBorder.Append(color);
            return topBorder;
        }

        private static RightBorder GenerateRightBorder(BorderStyleValues thickness)
        {
            RightBorder rightBorder = new RightBorder() { Style = thickness };
            Color color = new Color() { Indexed = 64U };
            rightBorder.Append(color);
            return rightBorder;
        }

        private static LeftBorder GeneerateLeftBorder(BorderStyleValues thickness)
        {
            LeftBorder leftBorder = new LeftBorder() { Style = thickness };
            Color color = new Color() { Indexed = 64U };
            leftBorder.Append(color);
            return leftBorder;
        }
        #endregion

        /// <summary>
        /// Преобразует числовой номер столбца в буквенное имя
        /// </summary>
        /// <param name="index">Номер столбца</param>
        /// <returns>Имя столбца</returns>
        internal string GetColumnName(int index)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";
            if (index >= alphabet.Length)
                value += alphabet[index / alphabet.Length - 1];
            value += alphabet[index % alphabet.Length - 1];
            return value;
        }

        /// <summary>
        /// Преобразует буквенное имя столбца в числовой номер, начиная с 1
        /// </summary>
        /// <param name="columnName">Буквенное имя столбца</param>
        /// <returns>Числовой номер столбца, начиная с 1</returns>
        internal int GetIndexFromName(string columnName)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (columnName.Length > 1)
                return alphabet.IndexOf(columnName.Substring(1)) + 27;

            return alphabet.IndexOf(columnName) + 1;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Сохраняет изменения в документе.
        /// </summary>
        public void Save()
        {
            _workbookPart.Workbook.Save();
            _document.Close();
        }

        /// <summary>
        /// Добавляет лист в документ Excel
        /// </summary>
        /// <param name="sheetName">Имя добавляемого листа</param>
        public void AddSpreadSheet(string sheetName)
        {
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());
            _workSheet = _worksheetPart.Worksheet;
            _workSheet.Save();

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

            //if (_worksheetPart.Worksheet.Elements<Sheet>().All(s => s.SheetId != sheetId))
            //    _worksheetPart.AddPart<Worksheet>(sheet);


            _workbookPart.Workbook.Save();
        }

        /// <summary>
        /// Записывает текст в виде строки в ячейку на листе
        /// </summary>
        /// <param name="text">Записываемый текст</param>
        /// <param name="sheetName">Имя листа</param>
        /// <param name="columnName">Адрес ячейки, имя столбца</param>
        /// <param name="rowIndex">Адрес ячейки, номер ряда</param>
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
            _workSheet.Save();
        }

        /// <summary>
        /// Задает границу определенной толщины у ячейки на листе
        /// </summary>
        /// <param name="sheetName">Лист, содержащий редактируемую ячейку</param>
        /// <param name="columnName">Адрес ячейки, имя столбца</param>
        /// <param name="rowIndex">Адрес ячейки, номер строки</param>
        /// <param name="thickness">Толщина границы</param>
        public void SetBorder(string sheetName, string columnName, uint rowIndex, BorderStyleValues thickness)
        {
            GetSpreadSheet(sheetName);

            Cell cell = GetCell(columnName, rowIndex);

            CellFormat cellFormat = GetOrCreateCellFormat(cell);

            cellFormat.BorderId = InsertBorder(GenerateBorder(thickness, true, true, true, true));

            cell.StyleIndex = InsertCellFormat(cellFormat);
        }

        /// <summary>
        /// Задает границу определенной толщины у диапазона ячеек на листе
        /// </summary>
        /// <param name="sheetName">Лист, содержащий редактируемый диапазно</param>
        /// <param name="columnName1">Имя столбца верхней левой ячейки диапазона</param>
        /// <param name="rowIndex1">Номер строки верхней левой ячейки диапазона</param>
        /// <param name="columnName2">Имя столбца нижней правой ячейки диапазона</param>
        /// <param name="rowIndex2">Номер строки нижней правой ячейки диапазона</param>
        /// <param name="thickness">Толщина границы</param>
        public void SetBorder(string sheetName, string columnName1, uint rowIndex1, string columnName2, uint rowIndex2,
            BorderStyleValues thickness)
        {
            GetSpreadSheet(sheetName);

            if (rowIndex1 == rowIndex2)
            {
                var rowCells = new List<Cell>();
                for (int i = GetIndexFromName(columnName1) + 1; i < GetIndexFromName(columnName2); i++)
                    rowCells.Add(GetCell(GetColumnName(i), rowIndex1));
                foreach (Cell cell in rowCells)
                {
                    CellFormat cellFormat = GetOrCreateCellFormat(cell);
                    cellFormat.BorderId =
                        InsertBorder(GenerateBorder(thickness, addBottomBorder: true, addTopBorder: true));
                    cell.StyleIndex = InsertCellFormat(cellFormat);
                }
                Cell leftCell = GetCell(columnName1, rowIndex1);
                CellFormat leftCellFormat = GetOrCreateCellFormat(leftCell);
                leftCellFormat.BorderId =
                    InsertBorder(GenerateBorder(thickness, addBottomBorder: true, addLeftBorder: true,
                        addTopBorder: true));
                leftCell.StyleIndex = InsertCellFormat(leftCellFormat);

                // todo: мусульманин (про Афган)
                Cell rightCell = GetCell(columnName2, rowIndex1);
                CellFormat rightCellFormat = GetOrCreateCellFormat(rightCell);
                rightCellFormat.BorderId =
                    InsertBorder(GenerateBorder(thickness, addBottomBorder: true, addTopBorder: true,
                        addRightBorder: true));
                rightCell.StyleIndex = InsertCellFormat(rightCellFormat);
            }

            else
            {
                // top border:
                var topCells = new List<Cell>();
                for (int i = GetIndexFromName(columnName1); i < GetIndexFromName(columnName2) + 1; i++)
                    topCells.Add(GetCell(GetColumnName(i), rowIndex1));
                foreach (Cell cell in topCells)
                {
                    CellFormat cellFormat = GetOrCreateCellFormat(cell);
                    cellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addTopBorder: true));
                    cell.StyleIndex = InsertCellFormat(cellFormat);
                }

                // bottom border:
                var bottomCells = new List<Cell>();
                for (int i = GetIndexFromName(columnName1); i < GetIndexFromName(columnName2) + 1; i++)
                    bottomCells.Add(GetCell(GetColumnName(i), rowIndex2));
                foreach (Cell cell in bottomCells)
                {
                    CellFormat cellFormat = GetOrCreateCellFormat(cell);
                    cellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addBottomBorder: true));
                    cell.StyleIndex = InsertCellFormat(cellFormat);
                }

                // left border:
                var leftCells = new List<Cell>();
                for (uint i = rowIndex1; i < rowIndex2 + 1; i++)
                    leftCells.Add(GetCell(columnName1, i));
                foreach (Cell cell in leftCells)
                {
                    CellFormat cellFormat = GetOrCreateCellFormat(cell);
                    cellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addLeftBorder: true));
                    cell.StyleIndex = InsertCellFormat(cellFormat);
                }

                // right border:
                var rightCells = new List<Cell>();
                for (uint i = rowIndex1; i < rowIndex2 + 1; i++)
                    rightCells.Add(GetCell(columnName2, i));
                foreach (Cell cell in rightCells)
                {
                    CellFormat cellFormat = GetOrCreateCellFormat(cell);
                    cellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addRightBorder: true));
                    cell.StyleIndex = InsertCellFormat(cellFormat);
                }

                // corner borders
                Cell leftTopCell = GetCell(columnName1, rowIndex1);
                CellFormat leftTopCellFormat = GetOrCreateCellFormat(leftTopCell);
                leftTopCellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addLeftBorder: true, addTopBorder: true));
                leftTopCell.StyleIndex = InsertCellFormat(leftTopCellFormat);

                Cell rightTopCell = GetCell(columnName2, rowIndex1);
                CellFormat rightTopCellFormat = GetOrCreateCellFormat(rightTopCell);
                rightTopCellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addRightBorder: true, addTopBorder: true));
                rightTopCell.StyleIndex = InsertCellFormat(rightTopCellFormat);

                Cell leftBottomCell = GetCell(columnName1, rowIndex2);
                CellFormat leftBottomCellFormat = GetOrCreateCellFormat(leftBottomCell);
                leftBottomCellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addLeftBorder: true, addBottomBorder: true));
                leftBottomCell.StyleIndex = InsertCellFormat(leftBottomCellFormat);

                Cell rightBottomCell = GetCell(columnName2, rowIndex2);
                CellFormat rightBottomCellFormat = GetOrCreateCellFormat(rightBottomCell);
                rightBottomCellFormat.BorderId = InsertBorder(GenerateBorder(thickness, addRightBorder: true, addBottomBorder: true));
                rightBottomCell.StyleIndex = InsertCellFormat(rightBottomCellFormat); 
            }
        }

        /// <summary>
        /// Меняет начертание текста ячейки на полужирное
        /// </summary>
        /// <param name="sheetName">Имя листа, содержащего редактируемую ячейку</param>
        /// <param name="columnName">Адрес ячейки, имя столбца</param>
        /// <param name="rowIndex">Адрес ячейки, номер строки</param>
        public void MakeBold(string sheetName, string columnName, uint rowIndex)
        {
            GetSpreadSheet(sheetName);
            Cell cell = GetCell(columnName, rowIndex);
            CellFormat cellFormat = GetOrCreateCellFormat(cell);

            Font font = new Font();
            Bold bold = new Bold();
            FontSize fontSize = new FontSize() { Val = 11D };
            Color color = new Color() { Theme = 1U };
            FontName name = new FontName() { Val = "Calibri" };
            FontFamilyNumbering numbering = new FontFamilyNumbering() { Val = 2 };
            FontScheme scheme = new FontScheme() { Val = FontSchemeValues.Minor };

            font.Append(bold);
            font.Append(fontSize);
            font.Append(color);
            font.Append(name);
            font.Append(numbering);
            font.Append(scheme);

            cellFormat.FontId = InsertFont(font);
            cellFormat.ApplyFont = true;

            cell.StyleIndex = InsertCellFormat(cellFormat);
        }

        /// <summary>
        /// Задает размер и шрифт текста в ячейке на листе
        /// </summary>
        /// <param name="sheetName">Лист, содержащий редактируемую ячейку</param>
        /// <param name="columnName">Адрес редактируемой ячейки - буквенное имя столбца</param>
        /// <param name="rowIndex">Адрес редактируемой ячейки - номер ряда (начиная с 1)</param>
        /// <param name="fontSize">Новый размер шрифта</param>
        /// <param name="fontName">Новый шрифт (по умолчанию - Calibri)</param>
        public void SetFont(string sheetName, string columnName, uint rowIndex, double fontSize, string fontName="Calibri")
        {
            GetSpreadSheet(sheetName);
            Cell cell = GetCell(columnName, rowIndex);
            CellFormat cellFormat = GetOrCreateCellFormat(cell);

            Font font = new Font();
            FontSize size = new FontSize() {Val = fontSize};
            Color color = new Color() {Theme = 1U};
            FontName name = new FontName() {Val = fontName};
            FontFamilyNumbering numbering = new FontFamilyNumbering() {Val = 2};
            FontScheme scheme = new FontScheme() {Val = FontSchemeValues.Minor};

            font.Append(size);
            font.Append(color);
            font.Append(name);
            font.Append(numbering);
            font.Append(scheme);

            cellFormat.FontId = InsertFont(font);
            cellFormat.ApplyFont = true;

            cell.StyleIndex = InsertCellFormat(cellFormat);
        }

        /// <summary>
        /// Вставляет в лист Excel столбец определенной ширины.
        /// </summary>
        /// <param name="sheetName">Лист, в который вставляем столбец</param>
        /// <param name="columnIndex">Номер столбца по порядку, начиная с 1</param>
        /// <param name="columnWidth">Ширина вставляемого столбца</param>
        public void AddColumn(string sheetName, uint columnIndex, double columnWidth)
        {
            GetSpreadSheet(sheetName);
            Columns columns = _workSheet.Elements<Columns>().FirstOrDefault();
            if (columns == null)
            {
                SheetData sheetData = _workSheet.Elements<SheetData>().FirstOrDefault();
                if (sheetData != null)
                    columns = _workSheet.InsertBefore(new Columns(), sheetData);
                else
                {
                    columns = new Columns();
                    _workSheet.Append(columns);
                }
            }
            Column column = new Column()
            {
                Min = columnIndex,
                Max = columnIndex,
                Width = columnWidth,
                CustomWidth = true
            };
            columns.Append(column);
        }

        /// <summary>
        /// Объединяет ячейки на листе
        /// </summary>
        /// <param name="sheetName">Целевой лист Excel</param>
        /// <param name="firstCellColumn">Имя столбца первой ячейки</param>
        /// <param name="firstCellRow">Номер строки первой ячейки</param>
        /// <param name="secondCellColumn">Имя столбца второй ячейки</param>
        /// <param name="secondCellRow">Номер строки второй ячейки</param>
        public void MergeCells(string sheetName, string firstCellColumn, uint firstCellRow, string secondCellColumn, uint secondCellRow)
        {
            GetSpreadSheet(sheetName);
            if (string.IsNullOrEmpty(firstCellColumn + firstCellRow) ||
                string.IsNullOrEmpty(secondCellColumn + secondCellRow)) return;
            CreateCellIfNotExists(firstCellColumn, firstCellRow);
            CreateCellIfNotExists(secondCellColumn, secondCellRow);

            MergeCells mergeCells;
            if (_workSheet.Elements<MergeCells>().Any())
                mergeCells = _workSheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();

                if (_workSheet.Elements<CustomSheetView>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<CustomSheetView>().First());
                else if (_workSheet.Elements<DataConsolidate>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<DataConsolidate>().First());
                else if (_workSheet.Elements<SortState>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<SortState>().First());
                else if (_workSheet.Elements<AutoFilter>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<AutoFilter>().First());
                else if (_workSheet.Elements<Scenarios>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<Scenarios>().First());
                else if (_workSheet.Elements<ProtectedRanges>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<ProtectedRanges>().First());
                else if (_workSheet.Elements<SheetProtection>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<SheetProtection>().First());
                else if (_workSheet.Elements<SheetCalculationProperties>().Any())
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<SheetCalculationProperties>().First());
                else
                    _workSheet.InsertAfter(mergeCells, _workSheet.Elements<SheetData>().First());
            }

            MergeCell mergeCell = new MergeCell()
            {
                Reference = new StringValue($"{firstCellColumn}{firstCellRow}:{secondCellColumn}{secondCellRow}")
            };
            mergeCells.Append(mergeCell);

            _workSheet.Save();
        }

        /// <summary>
        /// Поворачивает документ в ландшафтную ориентацию
        /// </summary>
        public void RotateLandscape()
        {
            var worksheetIds = _workbookPart.Workbook.Descendants<Sheet>().Select(s => s.Id.Value);
            foreach (string worksheetId in worksheetIds)
            {
                PageSetup pageSetup = _workSheet.Descendants<PageSetup>().FirstOrDefault();
                if (pageSetup != null)
                {
                    pageSetup = new PageSetup
                    {
                        Orientation = OrientationValues.Landscape,
                        PaperSize = 9U
                    };
                    _workSheet.AppendChild(pageSetup);
                }
                else
                {
                    pageSetup = new PageSetup() { Orientation = OrientationValues.Landscape };
                    _workSheet.AppendChild(pageSetup);
                }
                _workSheet.Save();
            }
            _workbookPart.Workbook.Save();
        }

        /// <summary>
        /// Удаляет ряд с листа
        /// </summary>
        /// <param name="sheetName">Лист, с которого необходимо удалить ряд</param>
        /// <param name="rowIndex">Номер удаляемого ряда, начиная с 1</param>
        public void RemoveRow(string sheetName, uint rowIndex)
        {
            GetSpreadSheet(sheetName);
            Row row = _workSheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex.Value == rowIndex);
            row.Remove();

            _workSheet.Save();
        }

        /// <summary>
        /// Задает выравнивание ячейки по размеру текста
        /// </summary>
        /// <param name="sheetName">Лист, содержащий редактируемую ячейку</param>
        /// <param name="columnName">Адрес редактируемой ячейки - буквенное имя столбца</param>
        /// <param name="rowIndex">Адрес редактируемой ячейки - номер ряда (начиная с 1)</param>
        /// <param name="isWrapped">Выравнивание ячейки по размеру текста (по умолчанию - да)</param>
        public void SetWrapping(string sheetName, string columnName, uint rowIndex, bool isWrapped = true)
        {
            GetSpreadSheet(sheetName);

            Cell cell = GetCell(columnName, rowIndex);
            CellFormat cellFormat = new CellFormat(new Alignment() {WrapText = isWrapped});
            cell.StyleIndex = InsertCellFormat(cellFormat);
        }
        #endregion
    }
}