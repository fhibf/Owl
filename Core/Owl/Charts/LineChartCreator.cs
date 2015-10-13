using System;
using System.Data;
using System.Windows.Forms.DataVisualization.Charting;

namespace Owl.Charts {

    internal class LineChartCreator {

        public void BuildChart(string fileName, DataTable data, LineChartContext context) {
            
            if (data == null || data.Columns.Count == 0)
                throw new System.FormatException("There is not data to plot.");
            
            System.Drawing.Color[] colors = ChartHelper.GetColors();
            if (data.Columns.Count - 1 > colors.Length)                
                throw new IndexOutOfRangeException("We don't have enough colors to render this chart.");

            using (var targetChart = new Chart()) {

                ChartArea mainChartArea = new ChartArea("MainArea");
                targetChart.ChartAreas.Add(mainChartArea);

                ChartHelper.InitializeChart(targetChart, context.ImageSize);
                             
                if (!string.IsNullOrWhiteSpace(context.Title)) {

                    targetChart.Titles.Add(new Title(context.Title, Docking.Top, 
                                                     new System.Drawing.Font(context.TitleFontName, 
                                                                             context.TitleFontSize,
                                                                             System.Drawing.FontStyle.Bold),
                                                                             context.TitleFontColor));
                }
                
                targetChart.DataSource = data;

                string xval = data.Columns[0].ToString();

                for (int i = 1; i < data.Columns.Count; i++) {

                    string columnName = data.Columns[i].ColumnName;
                    
                    targetChart.Series.Add(columnName);

                    // Using column name as the legend
                    targetChart.Series[i - 1].Name = columnName;
                    targetChart.Series[i - 1].LegendText = columnName;

                    targetChart.Series[i - 1].XValueMember = xval;
                    targetChart.Series[i - 1].XValueType = ChartHelper.ParseCharValueType(context.XValueType);
                    targetChart.Series[i - 1].YValueMembers = data.Columns[i].ToString();
                    targetChart.Series[i - 1].ChartType = SeriesChartType.Line;
             
                    targetChart.Series[i - 1].Color = colors[i - 1];
                }

                System.Drawing.Color lineColor = context.LineColor;

                targetChart.ChartAreas[0].AxisX.LineColor =
                    targetChart.ChartAreas[0].AxisY.LineColor =
                        targetChart.ChartAreas[0].AxisX.InterlacedColor =
                            targetChart.ChartAreas[0].AxisY.InterlacedColor =
                                targetChart.ChartAreas[0].AxisX.MajorGrid.LineColor =
                targetChart.ChartAreas[0].AxisY.MajorGrid.LineColor =
                    targetChart.ChartAreas[0].AxisX.TitleForeColor =
                        targetChart.ChartAreas[0].AxisY.TitleForeColor =
                            targetChart.ChartAreas[0].AxisX.MajorTickMark.LineColor =
                                targetChart.ChartAreas[0].AxisY.MajorTickMark.LineColor = lineColor;

                ChartHelper.AddLegend(targetChart, ChartHelper.ParseLegendStyle(context.LegendStyle));

                targetChart.DataBind();

                targetChart.SaveImage(fileName, ChartImageFormat.Png);
            }
        }        
    }
}
