using Owl.Common;
using System.Drawing;

namespace Owl.Charts {
    public abstract class ChartContext {

        public ImageSize ImageSize { get; set; }

        public Color LineColor { get; set; }

        public string Title { get; set; }

        public string TitleFontName { get; set; }

        public int TitleFontSize { get; set; }

        public Color TitleFontColor { get; set; }
        
        public ChartsLegendStyle LegendStyle { get; set; }


        public ChartContext() {

            this.ImageSize = new ImageSize();
            this.ImageSize.Width = 1390;
            this.ImageSize.Height = 696;
            this.Title = string.Empty;
            this.TitleFontName = "Verdana";
            this.TitleFontSize = 12;
            this.TitleFontColor = Color.Black;
            this.LineColor = Color.Gray;
            this.LegendStyle = ChartsLegendStyle.Table;
        }        
    }
}
