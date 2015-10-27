using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Owl.Word {
    internal class BookmarkReplacer {

        public void ReplaceBookmark(WordprocessingDocument package, string bookmark, string content) {
            
            IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

            foreach (BookmarkStart bookmarkStart in package.MainDocumentPart.RootElement.Descendants<BookmarkStart>()) {
                bookmarkMap[bookmarkStart.Name] = bookmarkStart;
            }

            foreach (BookmarkStart bookmarkStart in bookmarkMap.Values) {

                if (bookmarkStart.Name == bookmark) {

                    Run bookmarkText = bookmarkStart.NextSibling<Run>();
                    if (bookmarkText != null) {
                        bookmarkText.GetFirstChild<Text>().Text = content;
                    }
                    else {
                        var textElement = new Text(content);
                        var runElement = new Run(textElement);

                        bookmarkStart.InsertAfterSelf(runElement);
                    }
                }
            }
        }
    }
}
