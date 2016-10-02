
using System.Drawing;

namespace Owl.Word {

    public class TableStyle {

        public HorizontalAlignmentType Alignment { get; set; }

        public int RowFontSize { get; set; }

        public int TitleFontSize { get; set; }

        public int HeaderFontSize { get; set; }

        public string Title { get; set; }

        public bool ShowTitle { get; set; }

        public bool ShowHeader { get; set; }

        public bool EnableAlternativeBackgroundColor { set; get; }

        public Color TitleBackgroundColor { get; set; }

        public Color HeaderBackgroundColor { get; set; }

        public Color AlternativeBackgroundColor { get; set; }

        public TableStyle() {

            ShowHeader = true;
            ShowTitle = false;
            EnableAlternativeBackgroundColor = true;

            TitleFontSize = HeaderFontSize = RowFontSize = 9;

            Alignment = HorizontalAlignmentType.Left;

            TitleBackgroundColor = Color.FromArgb(128, 128, 128);
            HeaderBackgroundColor = Color.FromArgb(192, 192, 192);
            AlternativeBackgroundColor = Color.FromArgb(225, 230, 235);
        }
    }
}
