using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Owl.Word {
    internal class PageBreaker {

        public void BreakPage(WordprocessingDocument package) {

            Body body = package.MainDocumentPart.Document.Body;

            var newParagraph = new Paragraph(
                                      new Run(
                                        new Break() { Type = BreakValues.Page }));
            
            // Append new paragraph
            body.AppendChild(newParagraph);
        }
    }
}
