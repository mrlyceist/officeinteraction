using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using ExcelInteraction;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocument2Tests
    {
        private string _fileName = "test.xlsx";

        [TestCleanup]
        public void DeleteTestFiles()
        {
            if (File.Exists(_fileName))
                File.Delete(_fileName);
        }

        [TestMethod]
        public void BlankExcelDocumentIsCreated()
        {
            var xlDoc = new ExcelDocument2(_fileName);

            //var dataTable = NCore.General.GetTableFromExcel(_fileName);
            
            Assert.IsTrue(File.Exists(_fileName));
        }

        [TestMethod]
        public void StringValueIsAddedToTheFirstCellInFirstRow()
        {
            var xlDoc = new ExcelDocument2(_fileName);
            ExcelDocument2.InsertText(_fileName, "test", "report", "A", 1);

            Excel.Application application = new Excel.ApplicationClass();
            Excel.Workbook workbook = application.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), _fileName));
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
            var cell = (Excel.Range)sheet.Cells[1, 1];

            string test = (string)cell.Value;

            ExcelClose(workbook, application);

            Assert.AreEqual("test", test);
        }

        [TestMethod]
        public void CanCreateWithStaticMethods()
        {
            ExcelDocument2.CreateSpreadSheetWorkBook(_fileName);

            Assert.IsTrue(File.Exists(_fileName));
        }

        [TestMethod]
        public void CanWriteWithStaticMethods()
        {
            ExcelDocument2.CreateSpreadSheetWorkBook(_fileName);
            ExcelDocument2.InsertText(_fileName, "test", "report", "A", 1);

            Excel.Application application = new Excel.ApplicationClass();
            Excel.Workbook workbook = application.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), _fileName));
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
            var cell = (Excel.Range)sheet.Cells[1, 1];

            string test = (string)cell.Value;

            ExcelClose(workbook, application);

            Assert.AreEqual("test", test);
        }

        [TestMethod]
        public void CanSetCellBorder()
        {
            ExcelDocument2.CreateSpreadSheetWorkBook(_fileName);
            ExcelDocument2.InsertText(_fileName, "test", "report", "A", 1);
            ExcelDocument2.SetBorder(_fileName, "report", "A", 1);

            Excel.Application application = new Excel.ApplicationClass();
            Excel.Workbook workbook =
                application.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), _fileName));
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet) sheets.Item[1];
            var cell = (Excel.Range) sheet.Cells[1, 1];

            bool bottomBorder = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.GetHashCode() !=
                                Excel.XlLineStyle.xlLineStyleNone.GetHashCode();

            ExcelClose(workbook, application);

            Assert.IsTrue(bottomBorder);
        }

        #region Private Methods
        private void ExcelClose(Excel.Workbook workbook, Excel.Application application)
        {
            var missingObj = System.Reflection.Missing.Value;

            workbook.Close(false, missingObj, missingObj);
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);

            GC.Collect();
        } 
        #endregion
    }
}
