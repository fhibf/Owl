using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Owl.Word {
    internal class DocumentMerger {

        public void MergeDocument(WordprocessingDocument package, string documentPath) {

            string altChunkId = $"AltChunkId{Guid.NewGuid()}";

            MainDocumentPart mainPart = package.MainDocumentPart;

            AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, 
                                                                                        altChunkId);

            using (FileStream fileStream = File.Open(documentPath, FileMode.Open))
                chunk.FeedData(fileStream);

            AltChunk altChunk = new AltChunk();
            altChunk.Id = altChunkId;

            mainPart.Document.Body.AppendChild(altChunk);            
        }
    }
}
