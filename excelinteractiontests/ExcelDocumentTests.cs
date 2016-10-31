using System;
using System.IO;
using System.Runtime.InteropServices;
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

        [TestInitialize]
        public void DeleteTestFile()
        {
            if (File.Exists(_testFile))
                File.Delete(_testFile);
        }

        [TestCleanup]
        public void CleanupComObjects()
        {
            if (_application != null || _workbook != null)
                ExcelClose(_workbook, _application);
        }

        [TestMethod]
        public void CanCreateBlankDocument()
        {
            var xldoc = new ExcelDocument(_testFile);
            xldoc.AddSpreadSheet("test");
            xldoc.Save();

            Assert.IsTrue(File.Exists(_testFile));
        }

        [TestMethod]
        public void TestDocumentContainsOneSpreadSheet()
        {
            var xldoc = new ExcelDocument(_testFile);
            xldoc.AddSpreadSheet("test");
            xldoc.Save();

            _application = new Excel.ApplicationClass();
            _workbook = _application.Workbooks.Open(_testFile);
            //Excel.Sheets sheets = workbook.Worksheets;
            var sheetsCount = _workbook.Worksheets.Count;

            Assert.AreEqual(1, sheetsCount);
        }

        [TestMethod]
        public void CanWriteTextIntoACell()
        {
            string testText = "testText";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet("testSheet");
            xlDoc.InsertText(testText, "testSheet", "A", 1);
            xlDoc.Save();

            Excel.Range cell = GetTestCell();
            string testValue = (string) cell.Value;

            Assert.AreEqual(testText, testValue);
        }

        //[TestMethod]
        public void CanReadViaOleDb()
        {
            string testText = "testText";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet("testSheet");
            xlDoc.InsertText(testText, "testSheet", "A", 1);
            xlDoc.Save();

            //var dataTable = NCore.General.GetTableFromExcel("D:\\RefBook.xlsx");
            var dataTable = NCore.General.GetTableFromExcel(_testFile);
            string testValue = dataTable.Rows[1][1].ToString();

            Assert.AreEqual(testText, testValue);
        }

        [TestMethod]
        public void CanSetCellBorder()
        {
            string sheetName = "testSheet";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet(sheetName);
            xlDoc.InsertText("testText", sheetName, "A", 1);
            xlDoc.SetBorder(sheetName, "A", 1, BorderStyleValues.Thick);
            xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            Assert.IsTrue(bottomBorder);
        }

        [TestMethod]
        public void CanMakeCellBold()
        {
            string sheetName = "testSheet";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet(sheetName);
            xlDoc.InsertText("testText", sheetName, "A", 1);
            xlDoc.MakeBold(sheetName, "A", 1);
            xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool isBold = (bool) cell.Font.Bold;

            Assert.IsTrue(isBold);
        }

        [TestMethod]
        public void BoldAndBorderAreAppliedBoth()
        {
            string sheetName = "testSheet";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet(sheetName);
            xlDoc.InsertText("testText", sheetName, "A", 1);
            xlDoc.MakeBold(sheetName, "A", 1);
            xlDoc.SetBorder(sheetName, "A", 1, BorderStyleValues.Medium);
            xlDoc.Save();

            Excel.Range cell = GetTestCell();
            bool isBold = (bool) cell.Font.Bold;
            bool bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            Assert.IsTrue(isBold&&bottomBorder);
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