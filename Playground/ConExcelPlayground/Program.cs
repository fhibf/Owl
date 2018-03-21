using Owl.Data;
using Owl.Excel;
using System;
using System.Diagnostics;
using System.IO;

namespace ConExcelPlayground {
    class Program {
        static void Main(string[] args) {

            string document = Path.Combine(Path.GetTempPath(), "test.xlsx");

            if (File.Exists(document))
                File.Delete(document);

            var data = GetData();

            using (ExcelBuilder builder = new ExcelBuilder()) {

                builder.CreateDocument(document);

                builder.ImportData(data);

                var table1 = GetTable("Simple import 1");
                builder.ImportData(table1);

                var table2 = GetTable("Simple import 2");
                builder.ImportData(table2, "Table Alias");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
        }

        static DataSet GetData() {

            DataSet returnValue = new DataSet();

            for (int i = 0; i < 100; i++) {

                var table = GetTable("My Sample Chart " + i);
                returnValue.Tables.Add(table);
            }
            
            return returnValue;
        }

        static DataTable GetTable(string name) {

            DataTable data = new DataTable(name);

            data.Columns.AddRange(new DataColumn[]{
                    new DataColumn("Time", typeof(DateTime)),
                    new DataColumn("Group 1", typeof(int)),
                    new DataColumn("Group 2", typeof(int)),
                    new DataColumn("Group 3", typeof(int)),
                    new DataColumn("Group 4", typeof(int)),
                });

            DateTime baseDate = new DateTime(2015, 01, 01);
            Random rnd = new Random();
            for (int j = 0; j < 1000; j++) {

                int limiar = j * 10;

                var row = data.NewRow();
                row[0] = baseDate.AddDays(j);
                row[1] = rnd.Next(limiar, limiar + 30);
                row[2] = rnd.Next(limiar, limiar + 30);
                row[3] = rnd.Next(limiar, limiar + 30);
                row[4] = rnd.Next(limiar, limiar + 30);

                data.AddRow(row);
            }

            return data;
        }
    }
}
