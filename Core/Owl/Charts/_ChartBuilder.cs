using System.Collections.Generic;
using System.Data;

namespace Owl.Charts {
    public class ChartBuilder {

        public void BuildLineChart(string fileName, DataTable data, LineChartContext context) {

            LineChartCreator creator = new LineChartCreator();
            creator.BuildChart(fileName, data, context);
        }

        public void BuildPieChart(string fileName, IEnumerable<PieChartItem> data, PieChartContext context) {

            PieChartCreator creator = new PieChartCreator();
            creator.BuildChart(fileName, data, context);
        }
    }
}
