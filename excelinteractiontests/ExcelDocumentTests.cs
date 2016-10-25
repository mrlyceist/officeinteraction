using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocumentTests
    {
        private string _fileName = "test.xlsx";

        //[TestCleanup]
        public void DeleteTestFiles()
        {
            if (File.Exists(_fileName))
                File.Delete(_fileName);
        }

        [TestMethod]
        public void BlankExcelDocumentIsCreated()
        {
            var xlDoc = new ExcelInteraction.ExcelDocument(_fileName);

            //var dataTable = NCore.General.GetTableFromExcel(_fileName);
            
            Assert.IsTrue(File.Exists(_fileName));
        }

        [TestMethod]
        public void StringValueIsAddedToTheFirstCellInFirstRow()
        {
            var xlDoc = new ExcelInteraction.ExcelDocument(_fileName);
            ExcelInteraction.ExcelDocument.InsertText(_fileName, "test");

            //var dataTable = NCore.General.GetTableFromExcel(_fileName);

            //var test = dataTable.Rows.Count > 0 ? dataTable.Rows[0][0].ToString() : "empty";
            Excel.Application application = new Excel.ApplicationClass();
            //Excel.Workbook workBook = application.Workbooks.Open(_fileName, 0, false, 5,"","",false, )
            Excel.Workbook workbook = application.Workbooks.Open(_fileName);
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.Item[1];
            var cell = (Excel.Range)sheet.Cells[1, 1];

            string test = (string)cell.Value;

            ExcelClose(workbook, application);

            Assert.AreEqual("test", test);
        }

        private void ExcelClose(Excel.Workbook workbook, Excel.Application application)
        {
            var missingObj = System.Reflection.Missing.Value;

            workbook.Close(false, missingObj, missingObj);
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);

            GC.Collect();
        }
    }
}
