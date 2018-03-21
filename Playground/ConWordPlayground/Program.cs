using Owl.Word;
using Owl.Drawing;
using System.Diagnostics;
using System.IO;
using Owl.Data;

namespace ConWordPlayground
{
    class Program
    {
        static void Main(string[] args)
        {

            string document = Path.Combine(Path.GetTempPath(), "test.docx");

            if (File.Exists(document))
                File.Delete(document);

            using (WordBuilder builder = new WordBuilder())
            {

                builder.CreateDocument(document);

                builder.AddCustomStyle(new FormatStyle()
                {
                    Color = Color.Maroon,
                    FontName = "Courier New",
                    FontSize = 14,
                    IsBold = true,
                    IsItalic = true,
                    Name = "Warning",
                    StyleId = "Warning",
                    HighlightColor = FormatStyle.HighlightColors.Yellow
                });

                builder.CreateTextParagraph(TextParagraphType.Title, "Owl");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Owl is a custom framework to make easier the document's development using OpenXML. Enjoy it!");
                builder.CreateTextParagraph(TextParagraphType.Normal, string.Empty);

                builder.CreateTextParagraph(TextParagraphType.Title, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, "Heading 1");
                builder.CreateTextParagraph(TextParagraphType.Heading2, "Heading 2");
                builder.CreateTextParagraph(TextParagraphType.Heading3, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, "None style text.");
                builder.CreateTextParagraph("Warning", "Warning to this point. Pay attention!");

                builder.CreateTextParagraph(TextParagraphType.Title, HorizontalAlignmentType.Center, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, HorizontalAlignmentType.Center, "Heading 1");
                builder.CreateTextParagraph(TextParagraphType.Heading2, HorizontalAlignmentType.Center, "Heading 2");
                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Center, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, HorizontalAlignmentType.Center, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, HorizontalAlignmentType.Center, "None style text.");
                builder.CreateTextParagraph("Warning", HorizontalAlignmentType.Center, "Warning to this point. Pay attention!");

                builder.CreateTextParagraph(TextParagraphType.Title, HorizontalAlignmentType.Right, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, HorizontalAlignmentType.Right, "Heading 1");
                builder.CreateTextParagraph(TextParagraphType.Heading2, HorizontalAlignmentType.Right, "Heading 2");
                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Right, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, HorizontalAlignmentType.Right, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, HorizontalAlignmentType.Right, "None style text.");
                builder.CreateTextParagraph("Warning", HorizontalAlignmentType.Right, "Warning to this point. Pay attention!");

                //builder.AddImage("Super-IT.png");
                //builder.AddImage("Super-IT.png", HorizontalAlignmentType.Center);
                //builder.AddImage("Super-IT.png", HorizontalAlignmentType.Right);
                //builder.AddImage("Super-IT.png", 0.75M);
                //builder.AddImage("Super-IT.png", 0.50M);
                //builder.AddImage("Super-IT.png", 0.4M, HorizontalAlignmentType.Center);
                //builder.AddImage("Super-IT.png", 0.25M, HorizontalAlignmentType.Right);
            }

            using (WordBuilder builder = new WordBuilder())
            {

                builder.OpenDocument(document);

                builder.CreateTextParagraph(TextParagraphType.Title, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, "Heading 1");
                builder.CreateTextParagraph(TextParagraphType.Heading2, "Heading 2");
                builder.CreateTextParagraph(TextParagraphType.Heading3, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, "None style text.");
                builder.CreateTextParagraph("Warning", "Warning to this point. Pay attention!");

                //builder.AddImage("Super-IT.png");
                //builder.AddImage("Super-IT.png", HorizontalAlignmentType.Center);
                //builder.AddImage("Super-IT.png", HorizontalAlignmentType.Right);

                AddTable(builder);
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
        }

        private static void AddTable(WordBuilder builder)
        {

            var table = new DataTable();

            table.Columns.Add("Item", typeof(string));
            table.Columns.Add("Min", typeof(string));
            table.Columns.Add("Avg", typeof(string));
            table.Columns.Add("Max", typeof(string));

            for (int i = 0; i < 10; i++)
            {

                var r = table.NewRow();
                r[0] = i.ToString(); r[1] = 10M.ToString(); r[2] = 11M.ToString(); r[3] = 12M.ToString();

                table.Rows.Add(r);
            }

            builder.CreateTable(table, new TableStyle()
            {
                Alignment = HorizontalAlignmentType.Center,
                ShowTitle = true,
                Title = "My Table"
            });

        }
    }
}
