using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExporterLogicLibrary.Models;

namespace ExporterLogicLibrary
{
    public static class ExcelLogic
    {
        public static SpreadsheetDocument CreateSpreadsheetDocument(string fileName, string baseSheet = "")
        {
            SpreadsheetDocument output = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = output.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = GenerateStylesheet();
            workbookStylesPart.Stylesheet.Save();

            if (baseSheet != "")
                InsertWorksheet(output, baseSheet);
            else
                InsertWorksheet(output, Properties.Resources.STANDARD_SHEET_NAME);

            return output;
        }

        public static void SaveAndClose(this SpreadsheetDocument s)
        {
            s.WorkbookPart?.Workbook.Save();
            s.Dispose();
        }

        public static void InsertWorksheet(this SpreadsheetDocument s, string sheetName)
        {
            if (sheetName == "")
                throw new ArgumentException(Properties.Resources.EXCEPTION_MISSING_SHEET_NAME);

            WorkbookPart workbookPart = s.WorkbookPart ?? s.AddWorkbookPart();
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = (sheets.Elements<Sheet>().Select(s => s.SheetId?.Value).Max() + 1) ?? (uint)sheets.Elements<Sheet>().Count() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
        }

        public static SpreadsheetDocument OpenSpreadsheetDocument(string fileName)
        {
            SpreadsheetDocument output = SpreadsheetDocument.Open(fileName, true);
            return output;
        }

        private static void InsertLine(SpreadsheetDocument s, string sheetName, List<CellModel> fields, uint? styleIndex = null)
        {
            if (sheetName == "")
                throw new ArgumentException(Properties.Resources.EXCEPTION_MISSING_SHEET_NAME);

            WorkbookPart workbookPart = s.WorkbookPart ?? s.AddWorkbookPart();
            Workbook workbook = workbookPart.Workbook;
            Sheet sheet = workbook.Descendants<Sheet>().FirstOrDefault(sht => sht.Name == sheetName)
                ?? throw new ArgumentException(Properties.Resources.EXCEPTION_INVALID_SHEET_NAME.Replace("{SHEET_NAME}", sheetName));
            WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            Row lastRow = worksheetPart.Worksheet.Descendants<Row>().LastOrDefault();
            Row row = new()
            {
                RowIndex = lastRow == null ? 1 : lastRow.RowIndex + 1
            };

            for (int i = 0; i < fields.Count; i++)
            {
                row.Append(
                    GetNewCell(
                        fields[i].CellValueDataType,
                        GetExcelColumnName(i) + row.RowIndex,
                        fields[i].Value,
                        styleIndex != null ? (uint)styleIndex : fields[i].CellValueStyleIndex));
            }

            sheetData.Append(row);
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            string output = string.Empty;
            columnIndex++;

            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                output = Convert.ToChar('A' + modulo) + output;
                columnIndex = (columnIndex - modulo) / 26;
            }

            return output;
        }

        private static Cell GetNewCell(CellValues dataType, string cellReference, string cellValue, uint styleIndex)
        {
            Cell cell = new()
            {
                DataType = dataType,
                CellReference = cellReference,
                StyleIndex = styleIndex,
                CellValue = new CellValue(cellValue)
            };

            return cell;
        }

        private static Stylesheet GenerateStylesheet()
        {
            Fonts fonts = new(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }
                ));

            Fills fills = new(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            NumberingFormats numberingFormats = new(
                    new NumberingFormat() { NumberFormatId = 100U, FormatCode = StringValue.FromString("@") }
                );

            CellFormats cellFormats = new(
                    new CellFormat(), // default
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true, NumberFormatId = 100U }, // header
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true } // body
                );

            return new Stylesheet(fonts, fills, borders, cellFormats, numberingFormats);
        }

        public static void InsertDataLine(SpreadsheetDocument s, string baseSheet, List<CellModel> dataFields)
        {
            InsertLine(s, baseSheet, dataFields);
        }

        public static void InsertHeaderLine(SpreadsheetDocument s, string baseSheet, List<string> headerFields)
        {
            InsertLine(s, baseSheet, headerFields.Select(f => new CellModel() { Type = "", Value = f }).ToList(), 1);
        }
    }
}
