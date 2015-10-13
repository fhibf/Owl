using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Owl.Word;

namespace ConWordChunk {
    class Program {
        static void Main(string[] args) {

            #region [ Create parent document ]

            string documentParent = "parent.docx";

            if (File.Exists(documentParent))
                File.Delete(documentParent);

            using (WordBuilder builder = new WordBuilder()) {

                builder.CreateDocument(documentParent);

                builder.CreateTextParagraph(TextParagraphType.Title, "Owl");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Owl is a custom framework to make easier the document's development using OpenXML. Enjoy it!");
                builder.CreateTextParagraph(TextParagraphType.Normal, string.Empty);
            }

            #endregion

            #region [ Create child document ]

            string documentChild = "child.docx";

            if (File.Exists(documentChild))
                File.Delete(documentChild);

            using (WordBuilder builder = new WordBuilder()) {

                builder.CreateDocument(documentChild);

                builder.CreateTextParagraph(TextParagraphType.Heading1, "Owl");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Owl is a custom framework to make easier the document's development using OpenXML. Enjoy it!");
                builder.CreateTextParagraph(TextParagraphType.Normal, string.Empty);
            }

            #endregion

            using (WordBuilder builder = new WordBuilder()) {

                builder.OpenDocument(documentParent);

                builder.MergeDocument(documentChild);
            }
            
            ProcessStartInfo startInfo = new ProcessStartInfo(documentParent);
            Process.Start(startInfo);
        }
    }
}
