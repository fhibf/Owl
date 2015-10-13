using Owl.Charts;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace ConChartPlayground {
    class Program {
        static void Main(string[] args) {

            //CreateLineChart();

            CreatePieChart();
        }

        private static void CreatePieChart() {

            string image = "pieChart.png";

            if (File.Exists(image))
                File.Delete(image);

            var data = new PieChartItem[] {
              new PieChartItem() { Label= "Group 1", Value = 100 },  
              new PieChartItem() { Label= "Group 2", Value = 102 },
              new PieChartItem() { Label= "Group 3", Value = 103.565M },
              new PieChartItem() { Label= "Group 4", Value = 203.47M },
              new PieChartItem() { Label= "Group 5", Value = 100 },  
              new PieChartItem() { Label= "Group 6", Value = 102 },
              new PieChartItem() { Label= "Group 7", Value = 103.565M },
              new PieChartItem() { Label= "Group 8", Value = 203.47M },
              new PieChartItem() { Label= "Group 9", Value = 100 },  
              new PieChartItem() { Label= "Group 10", Value = 120 },
              new PieChartItem() { Label= "Group 11", Value = 130.565M },
              new PieChartItem() { Label= "Group 12", Value = 2300.47M }
            };
            
            ChartBuilder chartBuilder = new ChartBuilder();
            chartBuilder.BuildPieChart(image,
                                       data,
                                       new PieChartContext() { Title = "My sample chart!",
                                                               IsLabelOutside = true,
                                                               LabelStyle = PieChartLabelStyle.LabelPercent
                                                              
                                       });

            ProcessStartInfo startInfo = new ProcessStartInfo(image);
            Process.Start(startInfo);
        }

        private static void CreateLineChart() {
            
            string image = "lineChart.png";

            if (File.Exists(image))
                File.Delete(image);

            DataTable data = new DataTable("My Sample Chart");
            data.Columns.AddRange(new DataColumn[]{
                new DataColumn("Time", typeof(DateTime)),
                new DataColumn("Group 1", typeof(int)),
                new DataColumn("Group 2", typeof(int)),
                new DataColumn("Group 3", typeof(int)),
                new DataColumn("Group 4", typeof(int)),
            });

            DateTime baseDate = new DateTime(2015, 01, 01);
            Random rnd = new Random();
            for (int i = 0; i < 10; i++) {

                int limiar = i * 10;

                data.Rows.Add(baseDate.AddDays(i),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30));
            }

            ChartBuilder chartBuilder = new ChartBuilder();
            chartBuilder.BuildLineChart(image,
                                        data,
                                        new LineChartContext() { Title = "My sample chart!" });

            ProcessStartInfo startInfo = new ProcessStartInfo(image);
            Process.Start(startInfo);
        }
    }
}
