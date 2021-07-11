using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXMLExtension.Spreadsheet;
using System;
using System.IO;

namespace UnitTestProject
{
    [TestClass]
    public class SpreadsheetUnitTest
    {
        private string FilePath = @".\sampleFile.xlsx";
        private string SheetName1 = "ÉVÅ[Ég1";

        [TestMethod]
        public void CreateFile()
        {
            if (File.Exists(FilePath) == true)
            {
                File.Delete(FilePath);
            }
            using (var book = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook))
            {
                book.AddNewSheet(SheetName1);
            }
            Assert.IsTrue(File.Exists(FilePath));

            using (var book = SpreadsheetDocument.Open(FilePath, true))
            {
                Assert.IsTrue(book.HasSheet(SheetName1));
            }
        }

    }
}
