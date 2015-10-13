using Owl.Common;
using System.Windows.Forms.DataVisualization.Charting;

namespace Owl.Charts {
    internal class ChartHelper {

        public static LegendStyle ParseLegendStyle(ChartsLegendStyle legendStyle) {

            return (LegendStyle)(int)legendStyle;
        }

        public static ChartValueType ParseCharValueType(ChartsValueTypes valueType) {

            return (ChartValueType)(int)valueType;
        }

        public static void AddLegend(Chart chart, LegendStyle style) {
            
            var legend = new Legend();
            legend.Enabled = true;
            legend.LegendStyle = style;
            legend.Name = "MyLegend";
            legend.Position.Auto = true;

            chart.Legends.Add(legend);
        }

        public static  void InitializeChart(Chart chart, ImageSize imageSize) {

            chart.Titles.Clear();
            chart.Height = (int)imageSize.Height;
            chart.Width = (int)imageSize.Width;
            chart.Series.Clear();
            chart.DataSource = null;
            chart.Legends.Clear();
        }

        public static System.Drawing.Color[] GetColors() {
            
            //TODO: Pensar em uma solução para a paleta de cores... 
            return new System.Drawing.Color[] { System.Drawing.Color.Red, 
                                                System.Drawing.Color.Blue,
                                                System.Drawing.Color.Green, 
                                                System.Drawing.Color.Black,
                                                System.Drawing.Color.Silver,
                                                System.Drawing.Color.Purple,
                                                System.Drawing.Color.Crimson,
                                                System.Drawing.Color.Azure,
                                                System.Drawing.Color.Pink,
                                                System.Drawing.Color.Beige,
                                                System.Drawing.Color.DarkOrchid,
                                                System.Drawing.Color.Plum,
                                                System.Drawing.Color.Yellow,
                                                System.Drawing.Color.Red, 
                                                System.Drawing.Color.Blue,
                                                System.Drawing.Color.Green, 
                                                System.Drawing.Color.Black,
                                                System.Drawing.Color.Silver,
                                                System.Drawing.Color.Purple,
                                                System.Drawing.Color.Crimson,
                                                System.Drawing.Color.Azure,
                                                System.Drawing.Color.Pink,
                                                System.Drawing.Color.Beige,
                                                System.Drawing.Color.DarkOrchid,
                                                System.Drawing.Color.Plum,
                                                System.Drawing.Color.Yellow,
                                                System.Drawing.Color.Red, 
                                                System.Drawing.Color.Blue,
                                                System.Drawing.Color.Green, 
                                                System.Drawing.Color.Black,
                                                System.Drawing.Color.Silver,
                                                System.Drawing.Color.Purple,
                                                System.Drawing.Color.Crimson,
                                                System.Drawing.Color.Azure,
                                                System.Drawing.Color.Pink,
                                                System.Drawing.Color.Beige,
                                                System.Drawing.Color.DarkOrchid,
                                                System.Drawing.Color.Plum,
                                                System.Drawing.Color.Yellow
            };
        }
    }
}
