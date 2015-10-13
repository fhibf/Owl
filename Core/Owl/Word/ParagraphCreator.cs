using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Owl.Word {
    internal class ParagraphCreator {

        public void CreateTextParagraph(WordprocessingDocument package, TextParagraphType paragraphType, string text) {

            CreateTextParagraph(package, paragraphType, HorizontalAlignmentType.Left, text);
        }


        public void CreateTextParagraph(WordprocessingDocument package, TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text) {

            CreateTextParagraph(package, paragraphType, alignment, text, null);
        }

        public void CreateTextParagraph(WordprocessingDocument package, TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle) {

            switch (paragraphType) {
                case TextParagraphType.Title:
                    CreateTextParagraph(package, StyleCreator.TitleStyle, alignment, text, formatStyle);
                    break;
                case TextParagraphType.Normal:
                    CreateTextParagraph(package, StyleCreator.NormalStyle, alignment, text, formatStyle);
                    break;
                case TextParagraphType.Heading1:
                    CreateTextParagraph(package, StyleCreator.HeadingOneStyle, alignment, text, formatStyle);
                    break;
                case TextParagraphType.Heading2:
                    CreateTextParagraph(package, StyleCreator.HeadingTwoStyle, alignment, text, formatStyle);
                    break;
                case TextParagraphType.Heading3:
                    CreateTextParagraph(package, StyleCreator.HeadingThreeStyle, alignment, text, formatStyle);
                    break;
                case TextParagraphType.None:
                default:
                    CreateTextParagraph(package, string.Empty, alignment, text, formatStyle);
                    break;
            }
        }

        internal void CreateTextParagraph(WordprocessingDocument package, string style, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle) {

            Body body = package.MainDocumentPart.Document.Body;
            
            if (!string.IsNullOrEmpty(style) && !WordHelper.IsStyleIdInDocument(package, style)) {

                throw new System.ArgumentOutOfRangeException(style);
            }

            Run textRun = new Run();

            if (formatStyle != null) {

                // Get a reference to the RunProperties object.
                RunProperties runProperties = textRun.AppendChild(new RunProperties());

                if (formatStyle.IsBold) {

                    Bold bold = new Bold();
                    //bold.Val = OnOffValue.FromBoolean(true);
                    runProperties.AppendChild(bold);
                }
            }

            textRun.AppendChild(new Text(text));

            // Add a paragraph with a run and some text.
            Paragraph newParagraph =
                new Paragraph(textRun);
            
            // If the paragraph has no ParagraphProperties object, create one.
            WordHelper.CreateParagraphPropertiesIfNonExists(newParagraph);
            
            // Set paragraph alignment
            WordHelper.SetParagraphAlignment(newParagraph, alignment);

            if (!string.IsNullOrEmpty(style)) {

                // Get a reference to the ParagraphProperties object.
                ParagraphProperties pPr = newParagraph.ParagraphProperties;

                // If a ParagraphStyleId object doesn't exist, create one.
                if (pPr.ParagraphStyleId == null)
                    pPr.ParagraphStyleId = new ParagraphStyleId();

                // Set the style of the paragraph.
                pPr.ParagraphStyleId.Val = style;
            }
                    
            // Append new paragraph
            body.AppendChild(newParagraph);
        }

    }
}
