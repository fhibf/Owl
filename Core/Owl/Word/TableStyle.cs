
using System.Drawing;

namespace Owl.Word {

    public class TableStyle {

        public HorizontalAlignmentType Alignment { get; set; }

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

            Alignment = HorizontalAlignmentType.Left;

            TitleBackgroundColor = Color.FromArgb(105, 105, 105);
            HeaderBackgroundColor = Color.FromArgb(128, 128, 128);
            AlternativeBackgroundColor = Color.FromArgb(192, 192, 192);
        }
    }
}
