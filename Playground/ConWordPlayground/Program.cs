﻿using Owl.Word;
using System.Diagnostics;
using System.IO;

namespace ConWordPlayground {
    class Program {
        static void Main(string[] args) {

            string document = "test.docx";

            if (File.Exists(document))
                File.Delete(document);

            using (WordBuilder builder = new WordBuilder()) {

                builder.CreateDocument(document);
                
                builder.AddCustomStyle(new FormatStyle() {
                    Color = System.Drawing.Color.Maroon,
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
                
                builder.AddImage("Super-IT.png");
                builder.AddImage("Super-IT.png", HorizontalAlignmentType.Center);
                builder.AddImage("Super-IT.png", HorizontalAlignmentType.Right);
                builder.AddImage("Super-IT.png", 0.75M);
                builder.AddImage("Super-IT.png", 0.50M);
                builder.AddImage("Super-IT.png", 0.4M, HorizontalAlignmentType.Center);
                builder.AddImage("Super-IT.png", 0.25M, HorizontalAlignmentType.Right);
            }
            
            using (WordBuilder builder = new WordBuilder()) {

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

                builder.AddImage("Super-IT.png");
                builder.AddImage("Super-IT.png", HorizontalAlignmentType.Center);
                builder.AddImage("Super-IT.png", HorizontalAlignmentType.Right);

                builder.CreateTable(20, 8);
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
        }
    }
}