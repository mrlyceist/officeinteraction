using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using ExcelInteraction;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocumentTests
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
            var xlDoc = new ExcelDocument(_fileName);

            //var dataTable = NCore.General.GetTableFromExcel(_fileName);
            
            Assert.IsTrue(File.Exists(_fileName));
        }

        [TestMethod]
        public void StringValueIsAddedToTheFirstCellInFirstRow()
        {
            var xlDoc = new ExcelDocument(_fileName);
            ExcelDocument.InsertText(_fileName, "test", "report", "A", 1);

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
            ExcelDocument.CreateSpreadSheetWorkBook(_fileName);

            Assert.IsTrue(File.Exists(_fileName));
        }

        [TestMethod]
        public void CanWriteWithStaticMethods()
        {
            ExcelDocument.CreateSpreadSheetWorkBook(_fileName);
            ExcelDocument.InsertText(_fileName, "test", "report", "A", 1);

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
