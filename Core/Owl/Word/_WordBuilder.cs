using DocumentFormat.OpenXml.Packaging;
using Owl.Data;
using System;

namespace Owl.Word {
    public class WordBuilder : IDisposable {

        private WordprocessingDocument _package;

        public void CreateDocument(string filePath) {

            var creator = new DocumentCreator();
            this._package = creator.CreatePackage(filePath);

            CreateCommonStyles();
        }

        public void OpenDocument(string filePath) {

            var creator = new DocumentOpener();
            this._package = creator.OpenPackage(filePath);
        }

        public void CreateCommonStyles() {

            StyleCreator styleCreator = new StyleCreator(this._package);
            styleCreator.CreateBasicStyles();
        }

        public void AddCustomStyle(FormatStyle customStyle) {

            if (customStyle == null)
                throw new ArgumentNullException("customStyle");
            if (string.IsNullOrEmpty(customStyle.StyleId))
                throw new ArgumentNullException("StyleId");
            if (string.IsNullOrEmpty(customStyle.Name))
                throw new ArgumentNullException("Name");

            StyleCreator styleCreator = new StyleCreator(this._package);
            styleCreator.AddCustomStyle(customStyle);
        }

        //public void AddImage(string picturePath) {

        //    var imageCreator = new ImageCreator(this._package);
        //    imageCreator.InsertPicture(picturePath, 1.0M, HorizontalAlignmentType.Left);
        //}

        //public void AddImage(string picturePath, HorizontalAlignmentType alignment) {

        //    var imageCreator = new ImageCreator(this._package);
        //    imageCreator.InsertPicture(picturePath, 1.0M, alignment);
        //}
        
        //public void AddImage(string picturePath, decimal resizablePercent) {

        //    var imageCreator = new ImageCreator(this._package);
        //    imageCreator.InsertPicture(picturePath, resizablePercent, HorizontalAlignmentType.Left);
        //}

        //public void AddImage(string picturePath, decimal resizablePercent, HorizontalAlignmentType alignment) {

        //    var imageCreator = new ImageCreator(this._package);
        //    imageCreator.InsertPicture(picturePath, resizablePercent, alignment);
        //}

        public void CreateTextParagraph(TextParagraphType paragraphType, string text) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, HorizontalAlignmentType.Left, text);
        }

        public void CreateTextParagraph(TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, alignment, text);
        }

        public void CreateTextParagraph(TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, alignment, text, formatStyle);
        }

        public void CreateTextParagraph(string styleName, string text) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, HorizontalAlignmentType.Left, text, null);
        }

        public void CreateTextParagraph(string styleName, HorizontalAlignmentType alignment, string text) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, alignment, text, null);
        }

        public void CreateTextParagraph(string styleName, string text, FormatStyle formatStyle) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, HorizontalAlignmentType.Left, text, formatStyle);
        }

        public void CreateTextParagraph(string styleName, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle) {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, alignment, text, formatStyle);
        }

        public void CreateTable(DataTable table) {

            var tStyle = new TableStyle();

            CreateTable(table, tStyle);
        }

        public void CreateTable(DataTable table, TableStyle style) {

            TableCreator tableCreator = new TableCreator();
            tableCreator.AddTable(this._package, table, style);
        }

        public void MergeDocument(string documentPath) {

            DocumentMerger merger = new DocumentMerger();
            merger.MergeDocument(this._package, documentPath);
        }

        public void AppendBreakPage() {

            PageBreaker pageBreaker = new PageBreaker();
            pageBreaker.BreakPage(this._package);
        }

        public void ReplaceBookmark(string bookmark, string content) {

            BookmarkReplacer replacer = new BookmarkReplacer();
            replacer.ReplaceBookmark(this._package, bookmark, content);
        }

        public void Dispose() {

            if (this._package != null) {

                this._package.Close();
                this._package.Dispose();
            }
        }
    }
}
