using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace ExcelInteractionTests
{
    [TestClass]
    public class ExcelDocumentTests
    {
        [TestMethod]
        public void BlankExcelDocumentIsCreated()
        {
            var filename = "test.xlsx";
            var xlDoc = new ExcelInteraction.ExcelDocument(filename);
            //xlDoc.Save(filename);

            Assert.IsTrue(File.Exists(filename));
        }

        [TestMethod]
        public void StringValueIsAddedToTheFirstCellInFirstRow()
        {
            var fileName = "test.xlsx";
            var xlDoc = new ExcelInteraction.ExcelDocument(fileName);
            //xlDoc.InsertText()
            ExcelInteraction.ExcelDocument.InsertText(fileName, "test");

            var dataTable = NCore.General.GetTableFromExcel(fileName);

            string test = dataTable.Rows[0][0].ToString();

            Assert.AreEqual("test", test);
        }
    }
}
