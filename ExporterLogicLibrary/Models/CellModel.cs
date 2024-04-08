using DocumentFormat.OpenXml.Spreadsheet;

namespace ExporterLogicLibrary.Models
{
    public class CellModel
    {
        public required string Type { get; set; }
        public required string Value { get; set; }
        public CellFormatDefinition FormatDefinition { get; set; } = new();
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
        public CellFormatDefinition CellFormatDefintion
        {
            get
            {
                if (FormatDefinition != null)
                {
                    if (CellValueDataType == CellValues.Boolean)
                    {
                        FormatDefinition.NumberingFormat = new() { FormatCode = "LOGICAL" };
                    }
                    else if (CellValueDataType == CellValues.Number)
                    {
                        FormatDefinition.NumberingFormat = new() { FormatCode = "#,##0.00" };
                    }
                    else if (CellValueDataType == CellValues.Date)
                    {
                        FormatDefinition.NumberingFormat = new() { FormatCode = "mm-dd-yy" };
                    }
                    else
                    {
                        FormatDefinition.NumberingFormat = new() { FormatCode = "@" };
                    }
                }

                return FormatDefinition ?? new CellFormatDefinition();
            }
        }
    }
}
