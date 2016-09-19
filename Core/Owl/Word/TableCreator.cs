using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace Owl.Word {
    internal class TableCreator {

        public void AddTable(WordprocessingDocument package, DataTable table, TableStyle tableStyle) {

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

            SetTableAlignment(tableStyle.Alignment, tblProp);
            
            // Append the TableProperties object to the empty table.
            tbl.AppendChild<TableProperties>(tblProp);
            
            if (tableStyle.ShowTitle) {
                AddTitleRow(table, tableStyle, tbl);
            }

            if (tableStyle.ShowHeader) {
                AddHeaderRow(table, tbl, tableStyle);
            }

            AddRows(table, tableStyle, tbl, tblProp);

            body.Append(tbl);
        }

        private static void AddRows(DataTable table, TableStyle tableStyle, Table wordTable, TableProperties tableProperties) {
            
            for (int r = 0; r < table.Rows.Count; r++) {

                // Create a row.
                TableRow tRow = new TableRow();

                for (int c = 0; c < table.Columns.Count; c++) {

                    // Create a cell.
                    TableCell tCell = new TableCell();

                    if ((r % 2 == 0) && (tableStyle.EnableAlternativeBackgroundColor)) {

                        var tCellProperties = new TableCellProperties();

                        // Set Cell Color
                        Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.AlternativeBackgroundColor) };
                        tCellProperties.Append(tCellColor);

                        tCell.Append(tCellProperties);
                    }

                    string rowContent = table.Rows[r][c] == null ? string.Empty : table.Rows[r][c].ToString();


                    ParagraphProperties paragProperties = new ParagraphProperties();
                    SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
                    paragProperties.Append(spacingBetweenLines1);

                    var parag = new Paragraph();
                    parag.Append(paragProperties);
                    parag.Append(new Run(new Text(rowContent)));

                    // Specify the table cell content.
                    tCell.Append(parag);

                    // Append the table cell to the table row.
                    tRow.Append(tCell);
                }

                // Append the table row to the table.
                wordTable.Append(tRow);
            }
        }

        private static void AddHeaderRow(DataTable table, Table wordTable, TableStyle tableStyle) {
            
            // Create a row.
            TableRow tRow = new TableRow();

            foreach (DataColumn iColumn in table.Columns) {

                // Create a cell.
                TableCell tCell = new TableCell();

                TableCellProperties tCellProperties = new TableCellProperties();

                // Set Cell Color
                Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.HeaderBackgroundColor) };
                tCellProperties.Append(tCellColor);
                
                // Append properties to the cell
                tCell.Append(tCellProperties);

                ParagraphProperties paragProperties = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
                paragProperties.Append(spacingBetweenLines1);

                var parag = new Paragraph();
                parag.Append(paragProperties);
                parag.Append(new Run(new Text(iColumn.ColumnName)));

                // Specify the table cell content.
                tCell.Append(parag);
                
                // Append the table cell to the table row.
                tRow.Append(tCell);
            }

            // Append the table row to the table.
            wordTable.Append(tRow);
        }

        private static void AddTitleRow(DataTable table, TableStyle tableStyle, Table wordTable) {
            
            // Create a row.
            TableRow tRow = new TableRow();
            
            // Create a cell.
            TableCell tCell = new TableCell();

            TableCellProperties tCellProperties = new TableCellProperties();

            // Set Cell Color
            Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.TitleBackgroundColor) };
            tCellProperties.Append(tCellColor);

            // Set Cell Span
            GridSpan tCellSpan = new GridSpan() { Val = table.Columns.Count };
            tCellProperties.Append(tCellSpan);

            // Append properties to the cell
            tCell.Append(tCellProperties);
            
            ParagraphProperties paragProperties = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
            paragProperties.Append(justification1);
            paragProperties.Append(spacingBetweenLines1);
            
            var titleParagraph = new Paragraph();
            titleParagraph.Append(paragProperties);
            titleParagraph.Append(new Run(new Text(tableStyle.Title)));

            // Specify the table cell content.
            tCell.Append(titleParagraph);

            // Append the table cell to the table row.
            tRow.Append(tCell);

            // Append the table row to the table.
            wordTable.Append(tRow);
        }

        private static void SetTableAlignment(HorizontalAlignmentType alignment, TableProperties tblProp) {

            TableJustification tblJustification = null;

            if (alignment == HorizontalAlignmentType.Center) {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Center };                
            }
            else if (alignment == HorizontalAlignmentType.Right) {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Right };
            }
            else {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Left };
            }

            tblProp.Append(tblJustification);
        }
    }
}
