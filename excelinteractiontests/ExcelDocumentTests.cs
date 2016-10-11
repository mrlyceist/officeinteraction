using System;
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
            var xlDoc = new ExcelInteraction.ExcelDocument();
            xlDoc.Save("test.xlsx");

            Assert.IsTrue(File.Exists("test.xlsx"));
        }
    }
}
