using DocumentFormat.OpenXml.Packaging;

namespace Owl.Word {
    internal class DocumentOpener {

        // Opens a WordprocessingDocument.
        public WordprocessingDocument OpenPackage(string filePath) {

            WordprocessingDocument package = WordprocessingDocument.Open(filePath, true, new OpenSettings());

            return package;
        }
    }
}
