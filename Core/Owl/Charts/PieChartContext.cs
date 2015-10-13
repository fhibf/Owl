
namespace Owl.Charts {
    public class PieChartContext : ChartContext {

        public bool Show3DStyle { get; set; }

        public bool IsValueShownAsLabel { get; set; }

        public bool IsLabelOutside { get; set; }

        public PieChartLabelStyle LabelStyle { get; set; }

        public PieChartContext() : base() {

            this.Show3DStyle = true;
            this.IsValueShownAsLabel = false;
            this.LabelStyle = PieChartLabelStyle.LabelValuePercent;
            this.IsLabelOutside = true;
        }        
    }
}
