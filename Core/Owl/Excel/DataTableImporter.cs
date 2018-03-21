using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Owl.Data;
using System;
using System.Globalization;
using System.Linq;

namespace Owl.Excel {
    internal class DataTableImporter {

        private CultureInfo _culture;

        public DataTableImporter() {

            _culture = new CultureInfo("en-us");
        }

        public void ImportDataSet(SpreadsheetDocument package, DataSet dataset) {

            for (int i = 0; i < dataset.Tables.Count; i++) {

                DataTable targetTable = dataset.Tables[i];

                string sheetName = string.IsNullOrEmpty(targetTable.TableName) ? "sheet" + i : targetTable.TableName;

                ImportDataTable(package, targetTable, sheetName);
            }
        }

        public void ImportDataTable(SpreadsheetDocument package, DataTable table, string sheetName) {

            //populate the data into the spreadsheet
            WorkbookPart workbook = package.WorkbookPart;
            
            WorksheetPart worksheetPart = workbook.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            SheetData data = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            //add column names to the first row
            Row header = new Row();
            header.RowIndex = (UInt32)1;

            foreach (DataColumn column in table.Columns) {
                Cell headerCell = CreateTextCell(
                    table.Columns.IndexOf(column) + 1,
                    1,
                    column.ColumnName);

                header.AppendChild(headerCell);
            }
            data.AppendChild(header);

            //loop through each data row
            DataRow contentRow;
            for (int i = 0; i < table.Rows.Count; i++) {
                contentRow = table.Rows[i];
                data.AppendChild(CreateContentRow(contentRow, i + 2));
            }

            // Create Sheets object.
            Sheets sheets = workbook.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbook.GetIdOfPart(worksheetPart);

            // Create a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0) {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
        }


        /// <summary>
        /// Gets the Excel column name based on a supplied index number.
        /// </summary>
        /// <returns>1 = A, 2 = B... 27 = AA, etc.</returns>
        private string GetColumnName(int columnIndex) {

            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0) {
                modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

        private Cell CreateTextCell(int columnIndex, int rowIndex, object cellValue) { 
            
            Cell cell = new Cell();

            cell.DataType = CellValues.InlineString;
            cell.CellReference = GetColumnName(columnIndex) + rowIndex;

            InlineString inlineString = new InlineString();
            Text t = new Text();

            t.Text = cellValue.ToString();
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);

            return cell;
        }

        private Cell CreateNumberCell(int columnIndex, int rowIndex, object cellValue, Type dataType) {

            Cell cell = new Cell();

            cell.DataType = CellValues.Number;

            CellValue cellv = new CellValue();

            if (dataType == typeof(int))
                cellv.Text = ((int)cellValue).ToString(_culture);
            else if (dataType == typeof(float))
                cellv.Text = ((float)cellValue).ToString(_culture);
            else if (dataType == typeof(decimal))
                cellv.Text = ((decimal)cellValue).ToString(_culture);
            else if (dataType == typeof(double))
                cellv.Text = ((double)cellValue).ToString(_culture);
            else if (dataType == typeof(long))
                cellv.Text = ((long)cellValue).ToString(_culture);

            cell.Append(cellv);

            return cell;
        }

        private Row CreateContentRow(DataRow dataRow, int rowIndex) {

            Row row = new Row {
                RowIndex = (UInt32)rowIndex
            };

            for (int i = 0; i < dataRow.Table.Columns.Count; i++) {
                Cell dataCell = null;
                object value = dataRow[i];

                Type columnType = dataRow.Table.Columns[i].DataType;

                if (columnType == typeof(int) ||
                    columnType == typeof(decimal) ||
                    columnType == typeof(float) ||
                    columnType == typeof(long) ||
                    columnType == typeof(double)) {

                    dataCell = CreateNumberCell(i + 1, rowIndex, value, columnType);
                }
                else {
                    dataCell = CreateTextCell(i + 1, rowIndex, value);
                }

                row.AppendChild(dataCell);
            }
            return row;
        }
    }
}
