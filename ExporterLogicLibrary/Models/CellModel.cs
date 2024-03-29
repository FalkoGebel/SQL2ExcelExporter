using DocumentFormat.OpenXml.Spreadsheet;

namespace ExporterLogicLibrary.Models
{
    public class CellModel
    {
        public required string Type { get; set; }
        public required string Value { get; set; }
        public CellValues CellValueDataType
        {
            get
            {
                return Type switch
                {
                    "Int" => CellValues.Number,
                    _ => CellValues.InlineString,
                };
            }
        }
    }
}
