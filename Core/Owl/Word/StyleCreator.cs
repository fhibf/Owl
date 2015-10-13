using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace Owl.Word {
    internal class StyleCreator {

        internal const string StyledefinitionPartId = "rId1";

        public const string NoListStyle = "NoList";
        public const string DefaultParagraphFontStyle = "DefaultParagraphFont";
        public const string TableNormalStyle = "TableNormal";
        public const string HeadingOneCharStyle = "Heading1Char";
        public const string HeadingOneStyle = "Heading1";
        public const string HeadingThreeCharStyle = "Heading3Char";
        public const string HeadingThreeStyle = "Heading3";
        public const string HeadingTwoCharStyle = "Heading2Char";
        public const string HeadingTwoStyle = "Heading2";
        public const string NormalStyle = "Normal";
        public const string TitleCharStyle = "TitleChar";
        public const string TitleStyle = "Title";

        private WordprocessingDocument _document;

        public StyleCreator(WordprocessingDocument document) {

            this._document = document;
        }

        internal void CreateBasicStyles() {
            
            var styles = GetStyles();
            RegisterStyles(styles);
        }

        internal void AddCustomStyle(FormatStyle customStyle) {

            var newStyle = ParseStyle(customStyle);

            RegisterStyles(new Style[] { newStyle });
        }            
            
         private Style ParseStyle(FormatStyle customStyle) {

            // Create a new paragraph style and specify some of the properties.
            Style style = new Style() {
                Type = StyleValues.Paragraph,
                StyleId = customStyle.StyleId,
                CustomStyle = true
            };

            StyleName styleName1 = new StyleName() { Val = customStyle.Name };

            BasedOn basedOn1 = new BasedOn() { Val = customStyle.BasedOn };

            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = customStyle.NextParagraphStyle };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            string color = string.Format("#{0:X2}{1:X2}{2:X2}",
                                         customStyle.Color.R,
                                         customStyle.Color.G,
                                         customStyle.Color.B);

            Color color1 = new Color() { Val = color };
                         
            RunFonts font1 = new RunFonts() { Ascii = customStyle.FontName };

            // Specify a custom point size.
            FontSize fontSize1 = new FontSize() { Val = customStyle.FontSize.ToString() };

            if (customStyle.IsBold)
                styleRunProperties1.Append(new Bold());

            if (customStyle.HighlightColor != FormatStyle.HighlightColors.None)
                styleRunProperties1.Append(new Highlight() { Val = (HighlightColorValues)(int)customStyle.HighlightColor });

            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);

            if (customStyle.IsItalic)
                styleRunProperties1.Append(new Italic());

            style.Append(styleRunProperties1);

            return style;
        }
               
        private void RegisterStyles(IEnumerable<Style> styles) {

            //var styleDefinitionPart = (StyleDefinitionsPart)this._document.MainDocumentPart.GetPartById(StyledefinitionPartId);
            var styleDefinitionPart = this._document.MainDocumentPart.StyleDefinitionsPart;

            if (styles == null || styles.Count() == 0)
                return;

            foreach (Style itemStyle in styles) {

                styleDefinitionPart.Styles.Append(itemStyle);
            }
        }

        private IEnumerable<Style> GetStyles() {

            if (!WordHelper.IsStyleIdInDocument(this._document, DefaultParagraphFontStyle))
                yield return GenerateDefaultParagraphFontStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, NormalStyle))
                yield return GenerateNormalStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, TableNormalStyle))
                yield return GenerateTableNormalStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, NoListStyle))
                yield return GenerateNoListStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingOneStyle))
                yield return GenerateHeadingOneStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingOneCharStyle))        
                yield return GenerateHeadingOneCharStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingTwoStyle))        
                yield return GenerateHeadingTwoStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingTwoCharStyle))
                yield return GenerateHeadingTwoCharStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingThreeStyle))
                yield return GenerateHeadingThreeStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, HeadingThreeCharStyle))
                yield return GenerateHeadingThreeCharStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, TitleStyle))
                yield return GenerateTitleStyle();

            if (!WordHelper.IsStyleIdInDocument(this._document, TitleCharStyle))        
                yield return GenerateTitleCharStyle();
        }

        private Style GenerateHeadingThreeCharStyle() {

            Style style1 = new Style() { Type = StyleValues.Character, StyleId = HeadingThreeCharStyle, CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Heading 3 Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading3" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            //style1.Append(rsid1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateTitleCharStyle() {

            Style style1 = new Style() { Type = StyleValues.Character, StyleId = TitleCharStyle, CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Title Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Title" };
            UIPriority uIPriority1 = new UIPriority() { Val = 10 };
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Spacing spacing1 = new Spacing() { Val = -10 };
            Kern kern1 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize1 = new FontSize() { Val = "56" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "56" };
            
            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(spacing1);
            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);
            
            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            //style1.Append(rsid1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateTitleStyle() {

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = TitleStyle };
            StyleName styleName1 = new StyleName() { Val = "Title" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "TitleChar" };
            UIPriority uIPriority1 = new UIPriority() { Val = 10 };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(contextualSpacing1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Spacing spacing1 = new Spacing() { Val = -10 };
            Kern kern1 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize1 = new FontSize() { Val = "56" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "56" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(spacing1);
            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(primaryStyle1);
            //style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateHeadingTwoCharStyle() {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = HeadingTwoCharStyle, CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Heading 2 Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };
            
            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);
     
            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            //style1.Append(rsid1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateHeadingOneCharStyle() {

            Style style1 = new Style() { Type = StyleValues.Character, StyleId = HeadingOneCharStyle, CustomStyle = true };
            StyleName styleName1 = new StyleName() { Val = "Heading 1 Char" };
            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };
            
            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);
            
            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            //style1.Append(rsid1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateNoListStyle() {

            Style style1 = new Style() { Type = StyleValues.Numbering, StyleId = NoListStyle, Default = true };
            StyleName styleName1 = new StyleName() { Val = "No List" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style1.Append(styleName1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);

            return style1;
        }

        private Style GenerateTableNormalStyle() {

            Style style1 = new Style() { Type = StyleValues.Table, StyleId = TableNormalStyle, Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority1 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style1.Append(styleName1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);
            style1.Append(styleTableProperties1);

            return style1;
        }

        private Style GenerateDefaultParagraphFontStyle() {

            Style style1 = new Style() { Type = StyleValues.Character, StyleId = DefaultParagraphFontStyle, Default = true };
            StyleName styleName1 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style1.Append(styleName1);
            style1.Append(uIPriority1);
            style1.Append(semiHidden1);
            style1.Append(unhideWhenUsed1);

            return style1;
        }

        private Style GenerateHeadingThreeStyle() {

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = HeadingThreeStyle };
            StyleName styleName1 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading3Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(primaryStyle1);
            //style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateHeadingTwoStyle() {

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = HeadingTwoStyle };
            StyleName styleName1 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(unhideWhenUsed1);
            style1.Append(primaryStyle1);
            //style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateHeadingOneStyle() {

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = HeadingOneStyle };
            StyleName styleName1 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            //Rsid rsid1 = new Rsid() { Val = "0013195F" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(linkedStyle1);
            style1.Append(uIPriority1);
            style1.Append(primaryStyle1);
            //style1.Append(rsid1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        private Style GenerateNormalStyle() {

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = NormalStyle, Default = true };
            StyleName styleName1 = new StyleName() { Val = NormalStyle };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleRunProperties1);
        
            return style1;
        }
    }
}
