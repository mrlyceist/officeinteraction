using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

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

            var dataTable = NCore.General.GetTableFromExcel(_fileName);
            
            Assert.IsTrue(File.Exists(_fileName));
        }

        //[TestMethod]
        public void StringValueIsAddedToTheFirstCellInFirstRow()
        {
            var xlDoc = new ExcelInteraction.ExcelDocument(_fileName);
            ExcelInteraction.ExcelDocument.InsertText(_fileName, "test");

            var dataTable = NCore.General.GetTableFromExcel(_fileName);

            var test = dataTable.Rows.Count > 0 ? dataTable.Rows[0][0].ToString() : "empty";

            Assert.AreEqual("test", test);
        }
    }
}
