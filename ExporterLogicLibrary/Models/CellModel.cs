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
                    // Exact numerics -> Boolean
                    "bit" => CellValues.Boolean,

                    // Exact numerics -> Number
                    "bigint" => CellValues.Number,
                    "decimal" => CellValues.Number,
                    "int" => CellValues.Number,
                    "money" => CellValues.Number,
                    "numeric" => CellValues.Number,
                    "smallint" => CellValues.Number,
                    "smallmoney" => CellValues.Number,
                    "tinyint" => CellValues.Number,

                    // Approximate numerics -> Number
                    "float" => CellValues.Number,
                    "real" => CellValues.Number,

                    // Date and time -> Date
                    "date" => CellValues.Date,
                    "datetime2" => CellValues.Date,
                    "datetime" => CellValues.Date,
                    "datetimeoffset" => CellValues.Date,
                    "smalldatetime" => CellValues.Date,
                    "time" => CellValues.Date,

                    // Default -> String
                    _ => CellValues.String,
                };
            }
        }
        public uint CellValueStyleIndex
        {
            get
            {
                return CellValueDataType switch
                {
                    _ => 2
                };
            }
        }
    }
}
