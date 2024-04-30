using DocumentFormat.OpenXml.Packaging;
using ExporterLogicLibrary;
using ExporterLogicLibrary.Models;
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

        [TestMethod]
        public void CreateExcelFileWithSheetNameAndInsertHeaderLineForInvalidSheetNameAndException()
        {
            string fileName = _testPath + "\\SheetNameAndHeaderLineInvalidSheetName.xlsx";
            string baseSheet = "Sheet Name";
            string invalidSheetName = "Invalid Sheet Name";

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            try
            {
                Action act = () => ExcelLogic.InsertHeaderLine(s, invalidSheetName, []);
                act.Should().Throw<ArgumentException>().WithMessage($"Worksheet name \"{invalidSheetName}\" does not exist");
            }
            finally
            {
                s.SaveAndClose();
                File.Exists(fileName).Should().BeTrue();
            }
        }

        [TestMethod]
        public void CreateExcelFileWithSheetNameAndInsertHeaderLineAndExists()
        {
            string fileName = _testPath + "\\SheetNameAndHeaderLineSuccess.xlsx";
            string baseSheet = "Sheet Name";
            List<string> headerFields = ["col 1", "col 2", "col 3", "col 4", "col 5"];

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreateExcelFileInsertDataLineOnlyTextAndExists()
        {
            string fileName = _testPath + "\\DataLineOnlyTextSuccess.xlsx";
            string baseSheet = "Only Text Data";
            List<string> dataFields = ["cell 1", "cell 2", "other text", "once more", "number 5"];

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields.Select(f => new CellModel() { Type = "string", Value = f }).ToList());
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreateExcelFileInsertHeaderAndDataLineOnlyTextAndExists()
        {
            string fileName = _testPath + "\\HeaderAndDataLineOnlyTextSuccess.xlsx";
            string baseSheet = "Only Text Data";
            List<string> headerFields = ["col 1", "col 2", "col 3", "col 4", "col 5"];
            List<string> dataFields = ["cell 1", "cell 2", "other text", "once more", "number 5"];

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields.Select(f => new CellModel() { Type = "string", Value = f }).ToList());
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreateExcelFileInsertHeaderAndDataLineOnlyNumbersAndExists()
        {
            string fileName = _testPath + "\\HeaderAndDataLineOnlyNumberSuccess.xlsx";
            string baseSheet = "Only Text Data";
            List<string> headerFields = ["col 1", "col 2", "col 3", "col 4", "col 5"];
            List<CellModel> dataFields = [
                new CellModel() { Type = "int", Value = "2" },
                new CellModel() { Type = "int", Value = "5" },
                new CellModel() { Type = "int", Value = "6" },
                new CellModel() { Type = "int", Value = "99" },
                new CellModel() { Type = "int", Value = "1" }];

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        public void CreateExcelFileInsertHeaderAndTwoDataLinesOnlyNumbersAndExists()
        {
            string fileName = _testPath + "\\HeaderAndTwoDataLinesOnlyNumberSuccess.xlsx";
            string baseSheet = "Only Text Data";
            List<string> headerFields = ["col 1", "col 2", "col 3", "col 4", "col 5"];
            List<CellModel> dataFields = [
                new CellModel() { Type = "int", Value = "2" },
                new CellModel() { Type = "int", Value = "5" },
                new CellModel() { Type = "int", Value = "6" },
                new CellModel() { Type = "int", Value = "99" },
                new CellModel() { Type = "int", Value = "1" }];
            List<CellModel> dataFields2 = [
                new CellModel() { Type = "int", Value = "4" },
                new CellModel() { Type = "int", Value = "54" },
                new CellModel() { Type = "int", Value = "634" },
                new CellModel() { Type = "int", Value = "100" },
                new CellModel() { Type = "int", Value = "7" }];

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields2);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        [DataRow(1)]
        [DataRow(10)]
        [DataRow(100)]
        //[DataRow(1000)] // 10.5 seconds
        //[DataRow(10000)] // 12.5 minutes
        public void CreateExcelFileInsertHeaderAndTwoDataLinesOnlyNumbersAndExistsWithGivenNumberOfColumns(int numberOfColumns)
        {
            string fileName = _testPath + $"\\NumberOfColumns_{numberOfColumns}.xlsx";
            string baseSheet = "Only Text Data";
            List<string> headerFields = Enumerable.Range(1, numberOfColumns).Select(i => $"col {i}").ToList();
            List<CellModel> dataFields =
                Enumerable.Range(1, numberOfColumns).Select(i => new CellModel() { Type = "int", Value = $"{i}" }).ToList();
            List<CellModel> dataFields2 =
                Enumerable.Range(1, numberOfColumns).Select(i => new CellModel() { Type = "int", Value = $"{2 * i}" }).ToList();

            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields);
            ExcelLogic.InsertDataLine(s, baseSheet, dataFields2);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }

        [TestMethod]
        [DataRow(30, 1)]
        [DataRow(30, 10)]
        //[DataRow(30, 100)] // 10.8 seconds
        //[DataRow(30, 1000)] // 14 minutes
        //[DataRow(30, 10000)] // unfinished
        public void CreateExcelFileInsertHeaderAndTwoDataLinesOnlyNumbersAndExistsWithGivenNumberOfColumnsAndRows(int numberOfColumns, int numberOfRows)
        {
            string fileName = _testPath + $"\\NumberOfColumns_{numberOfColumns}_AndRows_{numberOfRows}.xlsx";
            string baseSheet = "Only Text Data";
            List<string> headerFields = Enumerable.Range(1, numberOfColumns).Select(i => $"col {i}").ToList();
            List<List<CellModel>> dataFields =
                Enumerable.Range(1, numberOfRows)
                          .Select(r => Enumerable.Range(1, numberOfColumns)
                                                 .Select(c => new CellModel() { Type = "int", Value = $"{r * c}" }).ToList()).ToList();
            SpreadsheetDocument s = ExcelLogic.CreateSpreadsheetDocument(fileName, baseSheet);
            s.SaveAndClose();
            s = ExcelLogic.OpenSpreadsheetDocument(fileName);
            ExcelLogic.InsertHeaderLine(s, baseSheet, headerFields);
            foreach (List<CellModel> row in dataFields)
                ExcelLogic.InsertDataLine(s, baseSheet, row);
            s.SaveAndClose();
            File.Exists(fileName).Should().BeTrue();
        }
    }
}