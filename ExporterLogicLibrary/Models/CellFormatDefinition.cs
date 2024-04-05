using DocumentFormat.OpenXml.Spreadsheet;

namespace ExporterLogicLibrary.Models
{
    public class CellFormatDefinition
    {
        public Font Font { get; set; } = new Font(new FontSize() { Val = 10 });
        public Fill Fill { get; set; } = new Fill(new PatternFill() { PatternType = PatternValues.None });
        public bool ApplyFill { get; set; } = false;
        public Border Border { get; set; } = new Border();
        public bool ApplyBorder { get; set; } = false;
        public NumberingFormat? NumberingFormat { get; set; } = null;
    }
}
