using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelInteraction;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocumentTests
    {
        private readonly string _testFile = Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
        private Excel.ApplicationClass _application;
        private Excel.Workbook _workbook;
        private string _sheetName;
        private ExcelDocument _xlDoc;

        #region Initialize And Cleanup
        [TestInitialize]
        public void DeleteTestFile()
        {
            if (File.Exists(_testFile))
                File.Delete(_testFile);
            GenerateExcel();
        }

        [TestCleanup]
        public void CleanupComObjects()
        {
            if (_application != null || _workbook != null)
                ExcelClose(_workbook, _application);
            _xlDoc = null;
        } 
        #endregion

        [TestMethod]
        public void CanCreateBlankDocument()
        {
            _xlDoc.Save();

            Assert.IsTrue(File.Exists(_testFile));
        }

        [TestMethod]
        public void TestDocumentContainsOneSpreadSheet()
        {
            _xlDoc.Save();

            _application = new Excel.ApplicationClass();
            _workbook = _application.Workbooks.Open(_testFile);
            var sheetsCount = _workbook.Worksheets.Count;

            Assert.AreEqual(1, sheetsCount);
        }

        [TestMethod]
        public void CanWriteTextIntoACell()
        {
            string testText = "testText";
            _xlDoc.InsertText(testText, "testSheet", "A", 1);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            string testValue = (string) cell.Value;

            Assert.AreEqual(testText, testValue);
        }

        //[TestMethod]
        public void CanReadViaOleDb()
        {
            string testText = "testText";
            _xlDoc.InsertText(testText, "testSheet", "A", 1);
            _xlDoc.Save();

            //var dataTable = NCore.General.GetTableFromExcel("D:\\RefBook.xlsx");
            var dataTable = NCore.General.GetTableFromExcel(_testFile);
            string testValue = dataTable.Rows[1][1].ToString();

            Assert.AreEqual(testText, testValue);
        }

        [TestMethod]
        public void CanSetCellBorder()
        {
            _xlDoc.InsertText("testText", _sheetName, "A", 1);
            _xlDoc.SetBorder(_sheetName, "A", 1, BorderStyleValues.Thick);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            Assert.IsTrue(bottomBorder);
        }

        [TestMethod]
        public void CanMakeCellBold()
        {
            _xlDoc.InsertText("testText", _sheetName, "A", 1);
            _xlDoc.MakeBold(_sheetName, "A", 1);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool isBold = (bool) cell.Font.Bold;

            Assert.IsTrue(isBold);
        }

        [TestMethod]
        public void BoldAndBorderAreAppliedBoth()
        {
            _xlDoc.InsertText("testText", _sheetName, "A", 1);
            _xlDoc.MakeBold(_sheetName, "A", 1);
            _xlDoc.SetBorder(_sheetName, "A", 1, BorderStyleValues.Medium);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool isBold = (bool) cell.Font.Bold;
            bool bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            Assert.IsTrue(isBold&&bottomBorder);
        }

        [TestMethod]
        public void CanSetNewCellWidth()
        {
            _xlDoc.AddColumn(_sheetName, 1, 15D);
            _xlDoc.InsertText("testText", _sheetName, "A", 1);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            double cellWidth = (double)cell.ColumnWidth;

            Assert.AreEqual(14.29, cellWidth);
        }

        //[TestMethod]
        public void TwoSheetsWorkFineTogether()
        {
            _xlDoc.AddSpreadSheet(DateTime.Now.Date.ToString());
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            var sheetsCount = _workbook.Sheets.Count;

            Assert.AreEqual(2, sheetsCount);
        }

        [TestMethod]
        public void CanMergeCells()
        {
            _xlDoc.MergeCells(_sheetName, "A", 1, "B", 1);
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool isMerged = (bool) cell.MergeCells;

            Assert.IsTrue(isMerged);
        }

        [TestMethod]
        public void CanRotateDocument()
        {
            _xlDoc.InsertText("testText", _sheetName, "A", 1);
            _xlDoc.RotateLandscape();
            _xlDoc.Save();

            Excel.Range cell = GetTestCell();
            Excel.Worksheet worksheet = (Excel.Worksheet) _workbook.Worksheets.Item[1];
            bool isLandscape = worksheet.PageSetup.Orientation == Excel.XlPageOrientation.xlLandscape;

            Assert.IsTrue(isLandscape);
        }

        [TestMethod]
        public void SheetNameReturnsRightString()
        {
            string name = _xlDoc.FirstSheetName;
            _xlDoc.Save();

            Assert.AreEqual(_sheetName, name);
        }

        [TestMethod]
        public void CanSetBorderOnRange()
        {
            _xlDoc.SetBorder(_sheetName, "A", 1, "B", 2, BorderStyleValues.Thick);
            _xlDoc.Save();

            _application = new Excel.ApplicationClass();
            _workbook = _application.Workbooks.Open(_testFile);
            Excel.Sheets sheets = _workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Item[1];
            var range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 2]];
            bool hasOuterBorders = range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                   Excel.XlLineStyle.xlLineStyleNone.GetHashCode() &&
                                   range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle.GetHashCode() !=
                                   Excel.XlLineStyle.xlLineStyleNone.GetHashCode() &&
                                   range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle.GetHashCode() !=
                                   Excel.XlLineStyle.xlLineStyleNone.GetHashCode() &&
                                   range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle.GetHashCode() !=
                                   Excel.XlLineStyle.xlLineStyleNone.GetHashCode();
            bool hasNotInnerBorders = range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle.GetHashCode() ==
                                      Excel.XlLineStyle.xlLineStyleNone.GetHashCode() &&
                                      range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle.GetHashCode() ==
                                      Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            Assert.IsTrue(hasOuterBorders && hasNotInnerBorders);
        }

        [TestMethod]
        public void CanDeleteRow()
        {
            _xlDoc.InsertText("TEST", _sheetName, "A", 1);
            _xlDoc.RemoveRow(_sheetName, 1);
            _xlDoc.Save();

            var testCell = GetTestCell();
            string testValue = (string) testCell.Value;

            Assert.IsTrue(string.IsNullOrEmpty(testValue));
        }

        [TestMethod]
        public void GetIndexFromNameReturns2ForB()
        {
            _xlDoc.Save();
            int index = _xlDoc.GetIndexFromName("B");
            int longIndex = _xlDoc.GetIndexFromName("AA");

            Assert.AreEqual(2, index);
            Assert.AreEqual(27, longIndex);
        }

        #region Private Methods
        private static void ExcelClose(Excel._Workbook workbook, Excel._Application application)
        {
            var missingObject = System.Reflection.Missing.Value;

            workbook.Close(false, missingObject, missingObject);
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);

            GC.Collect();
        }

        private void GenerateExcel()
        {
            _sheetName = "testSheet";
            _xlDoc = new ExcelDocument(_testFile);
            _xlDoc.AddSpreadSheet(_sheetName);
        }

        private Excel.Range GetTestCell()
        {
            _application = new Excel.ApplicationClass();
            _workbook = _application.Workbooks.Open(_testFile);
            Excel.Sheets sheets = _workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Item[1];
            var cell = (Excel.Range)sheet.Cells[1, 1];
            return cell;
        }
        #endregion
    }
}