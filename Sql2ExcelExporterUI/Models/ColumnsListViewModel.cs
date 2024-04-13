namespace Sql2ExcelExporterUI.Models
{
    public class ColumnsListViewModel
    {
        public required bool Supported { get; set; }
        public required bool Selected { get; set; }
        public required string Name { get; set; }
        public required string Type { get; set; }
    }
}
