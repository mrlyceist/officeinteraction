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

        //[TESTCLEANUP]
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
            Excel.Sheets sheets = workbook.Worksheets;
            var sheetsCount = workbook.Worksheets.Count;

            ExcelClose(workbook, application);

            Assert.AreEqual(1, sheetsCount);
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