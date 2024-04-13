namespace ExporterLogicLibrary
{
    public static class Mapping
    {
        public static string FormatCode(this string sqlDataType)
        {
            return sqlDataType switch
            {
                "bit" => "BOOLEAN",
                "int" => "0",
                "money" => "#,##0.00",
                "float" => "#,##0.00",
                "tinyint" => "BOOLEAN",

                // Not checked yet
                "bigint" => "",
                "decimal" => "",
                "numeric" => "",
                "smallint" => "",
                "smallmoney" => "",
                "real" => "",
                "date" => "",
                "datetime2" => "",
                "datetime" => "",
                "datetimeoffset" => "",
                "smalldatetime" => "",
                "time" => "",

                // Not supported
                "binary" => "",
                "image" => "",
                "varbinary" => "",

                // Default -> String
                _ => "@",
            };
        }
    }
}
