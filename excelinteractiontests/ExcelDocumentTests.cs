using System;
using System.IO;
using ExcelInteraction;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocumentTests
    {
        //private string _testFile = "test.xlsx";
        private readonly string _testFile = Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
        
        [TestInitialize]
        public void DeleteTestFile()
        {
            if (File.Exists(_testFile))
                File.Delete(_testFile);
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

            Excel.Application application = new Excel.ApplicationClass();
            Excel.Workbook workbook = application.Workbooks.Open(_testFile);
            //Excel.Sheets sheets = workbook.Worksheets;
            var sheetsCount = workbook.Worksheets.Count;

            ExcelClose(workbook, application);

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

            Excel.Application application = new Excel.ApplicationClass();
            Excel.Workbook workbook = application.Workbooks.Open(_testFile);
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet) sheets.Item[1];
            var cell = (Excel.Range) sheet.Cells[1, 1];
            string testValue = (string) cell.Value;
            ExcelClose(workbook, application);

            Assert.AreEqual(testText, testValue);
        }

        [TestMethod]
        public void CanReadViaOleDb()
        {
            string testText = "testText";
            var xlDoc = new ExcelDocument(_testFile);
            xlDoc.AddSpreadSheet("testSheet");
            xlDoc.InsertText(testText, "testSheet", "A", 1);
            xlDoc.Save();

            var dataTable = NCore.General.GetTableFromExcel(_testFile);
            string testValue = dataTable.Rows[0][0].ToString();

            Assert.AreEqual(testText, testValue);
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
        #endregion
    }
}