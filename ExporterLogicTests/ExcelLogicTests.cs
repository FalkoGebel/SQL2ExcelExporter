using DocumentFormat.OpenXml.Packaging;
using ExporterLogicLibrary;
using FluentAssertions;

namespace ExporterLogicTests
{
    [TestClass]
    public class ExcelLogicTests
    {
        private static readonly string _testPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Sql2ExcelExporterTest\\ExcelLogicLibraryTests\\";

        [ClassInitialize]
#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen
        public static void ClassInitialize(TestContext context)
#pragma warning restore IDE0060 // Nicht verwendete Parameter entfernen
        {
            DirectoryInfo di = new(_testPath);
            if (!di.Exists)
            {
                di.Create();
            }
        }

        [TestMethod]
        public void CreateExcelFileWithoutGivenSheetNameAndExists()
        {
            string fileName = _testPath + "\\NoSheetNameGiven.xlsx";

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreateExcelFileWithGivenSheetNameAndExists()
        {
            string fileName = _testPath + "\\SheetNameGiven.xlsx";

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, "Create");
            s.SaveAndClose();

            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreatedExcelFileWithoutSheetNameAndInsertSheetAndExists()
        {
            string fileName = _testPath + "\\NoSheetNameAndInsert.xlsx";

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertWorksheet(s, "Insert");
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }
    }
}