
namespace Owl.Charts {
    public class LineChartContext : ChartContext {

        public ChartsValueTypes XValueType { get; set; }

        public LineChartContext() : base() {

            this.XValueType = ChartsValueTypes.DateTime;
        }        
    }
}
