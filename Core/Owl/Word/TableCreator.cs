using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace Owl.Word {
    internal class TableCreator {

        public void AddTable(WordprocessingDocument package, DataTable table, HorizontalAlignmentType alignment) {

            Body body = package.MainDocumentPart.Document.Body;

            // Create an empty table.
            Table tbl = new Table();

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new BottomBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new LeftBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new RightBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new InsideHorizontalBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    },
                    new InsideVerticalBorder() {
                        Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 1
                    }
                )
            );

            // Append the TableProperties object to the empty table.
            tbl.AppendChild<TableProperties>(tblProp);

            for (int r = 0; r < table.Rows.Count; r++) {

                // Create a row.
                TableRow tr = new TableRow();

                for (int c = 0; c < table.Columns.Count; c++) {

                    // Create a cell.
                    TableCell tc1 = new TableCell();
                    
                    // Specify the width property of the table cell.
                    //tc1.Append(new TableCellProperties(
                    //    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

                    // Specify the table cell content.
                    tc1.Append(new Paragraph(new Run(new Text(table.Rows[r][c].ToString()))));

                    // Append the table cell to the table row.
                    tr.Append(tc1);
                    
                    //// Create a second table cell by copying the OuterXml value of the first table cell.
                    //TableCell tc2 = new TableCell(tc1.OuterXml);

                    //// Append the table cell to the table row.
                    //tr.Append(tc2);
                }

                // Append the table row to the table.
                tbl.Append(tr);
            }
                       
            body.Append(tbl);
        }
    }
}
