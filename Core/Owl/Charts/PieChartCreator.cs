using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;

namespace Owl.Charts {

    internal class PieChartCreator {

        public void BuildChart(string fileName, IEnumerable<PieChartItem> data, PieChartContext context) {
            
            if (data == null || data.Count() == 0)
                throw new System.FormatException("There is not data to plot.");

            using (var targetChart = new Chart()) {

                ChartArea mainChartArea = new ChartArea("MainArea");
                targetChart.ChartAreas.Add(mainChartArea);
                targetChart.ChartAreas[0].Area3DStyle.Enable3D = context.Show3DStyle;
                
                ChartHelper.InitializeChart(targetChart, context.ImageSize);

                if (!string.IsNullOrWhiteSpace(context.Title)) {

                    targetChart.Titles.Add(new Title(context.Title, Docking.Top,
                                                     new System.Drawing.Font(context.TitleFontName,
                                                                             context.TitleFontSize,
                                                                             System.Drawing.FontStyle.Bold),
                                                                             context.TitleFontColor));
                }

                targetChart.Series.Add(context.Title);
                targetChart.Series[0].XValueMember = "Label";
                targetChart.Series[0].YValueMembers = "Value";
                targetChart.DataSource = data;

                targetChart.Series[0].IsVisibleInLegend = true;
                targetChart.Series[0].ChartType = SeriesChartType.Pie;

                targetChart.Series[0].LegendText = "#VALX";
                
                // https://technet.microsoft.com/en-us/library/dd239373.aspx
                switch (context.LabelStyle) {
                    case PieChartLabelStyle.None:
                        targetChart.Series[0].Label = string.Empty;
                        targetChart.Series[0].XValueMember = string.Empty;
                        break;
                    case PieChartLabelStyle.Value:
                        targetChart.Series[0].Label = "#VALY";  
                        break;
                    case PieChartLabelStyle.Percent:
                        targetChart.Series[0].Label = "#PERCENT{P2}";                 
                        break;
                    case PieChartLabelStyle.ValuePercent:
                        targetChart.Series[0].Label = "#VALY / #PERCENT{P2}";                 
                        break;
                    case PieChartLabelStyle.Label:
                        targetChart.Series[0].Label = "#VALX";                 
                        break;
                    case PieChartLabelStyle.LabelValue:
                        targetChart.Series[0].Label = "#VALX (#VALY)";                 
                        break;
                    case PieChartLabelStyle.LabelPercent:
                        targetChart.Series[0].Label = "#VALX (#PERCENT{P2})";
                        break;
                    case PieChartLabelStyle.LabelValuePercent:
                        targetChart.Series[0].Label = "#VALX (#VALY / #PERCENT{P2})";                 
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("LabelStyle");
                }
                
                targetChart.Series[0].IsValueShownAsLabel = context.IsValueShownAsLabel;
           
                if (context.IsLabelOutside)
                    targetChart.Series[0]["PieLabelStyle"] = "Outside";
              
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
