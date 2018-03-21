
namespace Owl.Word {
    public class FormatStyle {

        public string StyleId { get; set; }

        public string Name { get; set; }

        public int SpacingBeforeLines { get; set; }

        public int SpacingAfterLines { get; set; }

        public string BasedOn { get; set; }

        public string NextParagraphStyle { get; set; }

        public bool IsBold { get; set; }

        public bool IsItalic { get; set; }

        public int FontSize { get; set; }
        
        public string FontName { get; set; }

        public Drawing.Color Color { get; set; }

        public HighlightColors HighlightColor { get; set; }

        public FormatStyle() {

            this.HighlightColor = FormatStyle.HighlightColors.None;
            this.Color = Drawing.Color.FromArgb(00, 00, 00);            
            this.FontName = "Arial";
            this.FontSize = 16;
            this.BasedOn = "Normal";
            this.NextParagraphStyle = "Normal";
        }

        public enum HighlightColors {
            Black = 0,
            Blue = 1,
            Cyan = 2,
            Green = 3,
            Magenta = 4,
            Red = 5,
            Yellow = 6,
            White = 7,
            DarkBlue = 8,
            DarkCyan = 9,
            DarkGreen = 10,
            DarkMagenta = 11,
            DarkRed = 12,
            DarkYellow = 13,
            DarkGray = 14,
            LightGray = 15,
            None = 16,
        }
    }
}
