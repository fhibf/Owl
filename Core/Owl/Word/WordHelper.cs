using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace Owl.Word {
    internal class WordHelper {

        internal static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleid) {

            // Get access to the Styles element for this document.
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style == null)
                return false;

            return true;
        }

        internal static void SetParagraphAlignment(Paragraph paragraph, HorizontalAlignmentType alignment) {

            // Set picture alignment
            if (alignment == HorizontalAlignmentType.Center)
                paragraph.ParagraphProperties.Append(new Justification() { Val = JustificationValues.Center });
            else if (alignment == HorizontalAlignmentType.Right)
                paragraph.ParagraphProperties.Append(new Justification() { Val = JustificationValues.Right });
            else
                paragraph.ParagraphProperties.Append(new Justification() { Val = JustificationValues.Left });
        }

        internal static void CreateParagraphPropertiesIfNonExists(Paragraph paragraph) {

            // If the paragraph has no ParagraphProperties object, create one.
            if (paragraph.Elements<ParagraphProperties>().Count() == 0) {
                paragraph.PrependChild<ParagraphProperties>(new ParagraphProperties());
            }
        }
    }
}
