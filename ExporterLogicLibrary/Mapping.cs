namespace ExporterLogicLibrary
{
    public static class Mapping
    {
        public static string FormatCode(this string sqlDataType)
        {
            return sqlDataType switch
            {
                "bit" => "BOOLEAN",
                "decimal" => "#,##0.00",
                "int" => "0",
                "money" => "#,##0.00",
                "float" => "#,##0.00",
                "tinyint" => "BOOLEAN",

                // Not checked yet
                "bigint" => "",
                "numeric" => "",
                "smallint" => "",
                "smallmoney" => "",
                "real" => "",
                "date" => "",
                "datetime" => "",
                "datetime2" => "",
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
