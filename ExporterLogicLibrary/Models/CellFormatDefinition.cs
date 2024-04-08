using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;

namespace ExporterLogicLibrary.Models
{
    public class CellFormatDefinition
    {
        public string? FontName { get; set; } = null;
        public double? FontSize { get; set; } = null;
        public System.Drawing.Color? FontColor { get; set; } = null;
        public bool Bold { get; set; } = false;
        public bool Italic { get; set; } = false;
        public bool Underline { get; set; } = false;
        public Font Font
        {
            get
            {
                Font font = new();

                if (!string.IsNullOrWhiteSpace(FontName))
                    font.Append(new FontName() { Val = FontName });

                FontSize fontSize = new();
                if (FontSize != null)
                    fontSize.Val = FontSize;
                else
                    fontSize.Val = 10;
                font.Append(fontSize);

                if (FontColor != null)
                {
                    font.Append(new DocumentFormat.OpenXml.Spreadsheet.Color()
                    {
                        Rgb = new HexBinaryValue
                        {
                            Value = ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    ((System.Drawing.Color)FontColor).A,
                                    ((System.Drawing.Color)FontColor).R,
                                    ((System.Drawing.Color)FontColor).G,
                                    ((System.Drawing.Color)FontColor).B)).Replace("#", "")
                        }
                    });
                }

                if (Bold)
                    font.Append(new Bold());

                if (Italic)
                    font.Append(new Italic());

                if (Underline)
                    font.Append(new Underline());

                return font;
            }
        }
        public System.Drawing.Color? FillColor { get; set; } = null;
        public Fill Fill
        {
            get
            {
                if (FillColor != null)
                    return new Fill(new PatternFill(
                        new ForegroundColor()
                        {
                            Rgb = new HexBinaryValue()
                            {
                                Value = ColorTranslator.ToHtml(
                                    System.Drawing.Color.FromArgb(
                                        ((System.Drawing.Color)FillColor).A,
                                        ((System.Drawing.Color)FillColor).R,
                                        ((System.Drawing.Color)FillColor).G,
                                        ((System.Drawing.Color)FillColor).B)).Replace("#", "")
                            }
                        })
                    { PatternType = PatternValues.Solid });

                return new Fill(new PatternFill() { PatternType = PatternValues.None });
            }
        }
        public bool ApplyFill
        {
            get
            {
                return FillColor != null;
            }
        }
        public System.Drawing.Color? BorderColor { get; set; } = null;
        public bool BorderThick { get; set; } = false;
        public Border Border
        {
            get
            {
                if (BorderColor != null)
                {
                    return new Border(
                        new LeftBorder(new DocumentFormat.OpenXml.Spreadsheet.Color()
                        {
                            Rgb = new HexBinaryValue()
                            {
                                Value = ColorTranslator.ToHtml(
                                            System.Drawing.Color.FromArgb(
                                                ((System.Drawing.Color)BorderColor).A,
                                                ((System.Drawing.Color)BorderColor).R,
                                                ((System.Drawing.Color)BorderColor).G,
                                                ((System.Drawing.Color)BorderColor).B)).Replace("#", "")
                            }
                        })
                        { Style = BorderThick ? BorderStyleValues.Thick : BorderStyleValues.Thin },
                        new RightBorder(new DocumentFormat.OpenXml.Spreadsheet.Color()
                        {
                            Rgb = new HexBinaryValue()
                            {
                                Value = ColorTranslator.ToHtml(
                                            System.Drawing.Color.FromArgb(
                                                ((System.Drawing.Color)BorderColor).A,
                                                ((System.Drawing.Color)BorderColor).R,
                                                ((System.Drawing.Color)BorderColor).G,
                                                ((System.Drawing.Color)BorderColor).B)).Replace("#", "")
                            }
                        })
                        { Style = BorderThick ? BorderStyleValues.Thick : BorderStyleValues.Thin },
                        new TopBorder(new DocumentFormat.OpenXml.Spreadsheet.Color()
                        {
                            Rgb = new HexBinaryValue()
                            {
                                Value = ColorTranslator.ToHtml(
                                        System.Drawing.Color.FromArgb(
                                            ((System.Drawing.Color)BorderColor).A,
                                            ((System.Drawing.Color)BorderColor).R,
                                            ((System.Drawing.Color)BorderColor).G,
                                            ((System.Drawing.Color)BorderColor).B)).Replace("#", "")
                            }
                        })
                        { Style = BorderThick ? BorderStyleValues.Thick : BorderStyleValues.Thin },
                        new BottomBorder(new DocumentFormat.OpenXml.Spreadsheet.Color()
                        {
                            Rgb = new HexBinaryValue()
                            {
                                Value = ColorTranslator.ToHtml(
                                        System.Drawing.Color.FromArgb(
                                            ((System.Drawing.Color)BorderColor).A,
                                            ((System.Drawing.Color)BorderColor).R,
                                            ((System.Drawing.Color)BorderColor).G,
                                            ((System.Drawing.Color)BorderColor).B)).Replace("#", "")
                            }
                        })
                        { Style = BorderThick ? BorderStyleValues.Thick : BorderStyleValues.Thin },
                        new DiagonalBorder());
                }

                return new Border();
            }
        }
        public bool ApplyBorder
        {
            get
            {
                return BorderColor != null;
            }
        }
        public NumberingFormat? NumberingFormat { get; set; } = null;
    }
}
