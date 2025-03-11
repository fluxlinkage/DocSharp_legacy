using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DrawingML = DocumentFormat.OpenXml.Drawing;

namespace DocSharp.Docx {

public class DocxToMarkdownConverter : DocxConverterBase
{
    /// <summary>
    /// If this property is set to an existing directory, images will be exported to that folder
    /// and a reference will be added in Markdown syntax,
    /// otherwise images are not converted.
    /// </summary>
    public string? ImagesOutputFolder { get; set; } = string.Empty;

    /// <summary>
    /// This property is used in combination with ImagesOutputFolder to determine 
    /// how the image files are specified in Markdown.
    /// 
    /// If this property is set to null, an absolute path such as "file:///c:/.../image.jpg" 
    /// will be created using the ImagesOutputFolder value and the image file name.
    /// 
    /// Otherwise, the base path (exluding the image file name) is replaced by this value.
    /// Possible values:
    /// - empty string or "." : images are expected to be in the same folder as the Markdown file.
    /// - relative paths such as "images" or "../images": images are expected to be in a subfolder or parent folder.
    /// - "/server/user/files/" or "C:\images": replaces the file path entirely
    /// (the image file name is still appended and Windows paths are converted to the file URI scheme).
    /// 
    /// This property does not affect where the images are actually saved, and can be useful if
    /// the Markdown document is not saved to file, or in environments with limited file system access.
    /// </summary>
    public string? ImagesBaseUriOverride { get; set; } = null;

    private char[] _specialChars = { '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '<', '>',
                                     '#', '+', '-', '!', '|', '~' };

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        var numberingProperties = OpenXmlHelpers.GetEffectiveProperty<NumberingProperties>(paragraph);
        if (numberingProperties != null)
        {
            ProcessListItem(numberingProperties, sb);
        }
        else if (paragraph.ParagraphProperties?.ParagraphStyleId != null)
        {
            var styles = paragraph.GetStylesPart();
            var style = styles.GetStyleFromId(paragraph.ParagraphProperties.ParagraphStyleId.Val, StyleValues.Paragraph);
            if (style?.StyleName?.Val?.Value != null)
            {
                switch (style.StyleName.Val.Value.ToLower())
                {
                    case "heading 1":
                    case "title":
                        sb.Append("# ");
                        break;
                    case "heading 2":
                    case "subtitle":
                        sb.Append("## ");
                        break;
                    case "heading 3":
                        sb.Append("### ");
                        break;
                    case "heading 4":
                        sb.Append("#### ");
                        break;
                    case "heading 5":
                        sb.Append("##### ");
                        break;
                    case "heading 6":
                        sb.Append("###### ");
                        break;
                }
            }
        }
        base.ProcessParagraph(paragraph, sb);
        sb.AppendLine();
        sb.AppendLine();
    }

    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID == numPr.NumberingId.Val);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                               x.LevelIndex == levelIndex);
                if (level != null &&
                    level.NumberingFormat?.Val is EnumValue<NumberFormatValues> listType &&
                    listType != NumberFormatValues.None)
                {
                    for (int i = 0; i < levelIndex; i++)
                    {
                        sb.Append("    "); // indentation
                    }
                    if (listType == NumberFormatValues.Bullet)
                    {
                        sb.Append("- ");
                    }
                    else
                    {
                        var startNumber = level.StartNumberingValue?.Val ?? 1;
                        sb.Append($"{startNumber}. "); // Markdown renderers will automatically increase the number.
                    }
                }
            }
        }
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var text = run.GetFirstChild<Text>()?.InnerText;
        bool hasText = !string.IsNullOrWhiteSpace(text);

        bool isBold, isItalic, isUnderline, isStrikethrough, isHighlight, isSubscript, isSuperscript;
        isBold = isItalic = isUnderline = isStrikethrough = isHighlight = isSubscript = isSuperscript = false;

        string leadingSpaces = string.Empty;
        string trailingSpaces = string.Empty;

        if (hasText)
        {
            leadingSpaces = StringHelpers.GetLeadingSpaces(text!);
            sb.Append(leadingSpaces);

            // TODO: consider last child for trailing spaces
            trailingSpaces = StringHelpers.GetTrailingSpaces(text!);

            // Formatting options of type OnOffValue such as bold and italic are considered enabled
            // if the element is present, unless value is explicitly set to false.
            // (e.g. <w:b /> without value means bold is enabled, otherwise it would not be present at all)
            isBold = OpenXmlHelpers.GetEffectiveProperty<Bold>(run) is Bold b && (b.Val is null || b.Val);
            isItalic = OpenXmlHelpers.GetEffectiveProperty<Italic>(run) is Italic i && (i.Val is null || i.Val);

            isUnderline = OpenXmlHelpers.GetEffectiveProperty<Underline>(run) is Underline u && 
                          u.Val != null && u.Val != UnderlineValues.None;

            isStrikethrough = (OpenXmlHelpers.GetEffectiveProperty<DoubleStrike>(run) is DoubleStrike ds &&
                          (ds.Val is null || ds.Val)) ||
                          (OpenXmlHelpers.GetEffectiveProperty<Strike>(run) is Strike s &&
                          (s.Val is null || s.Val));

            isHighlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run) is Highlight h &&
                          h.Val != null && h.Val != HighlightColorValues.None;

            var vta = OpenXmlHelpers.GetEffectiveProperty<VerticalTextAlignment>(run);
            isSubscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Subscript;
            isSuperscript = vta != null && vta.Val != null && vta.Val == VerticalPositionValues.Superscript;           

            if (isItalic)
                sb.Append("*");

            if (isBold)
                sb.Append("**");

            if (isStrikethrough)
                sb.Append("~~");

            if (isUnderline)
                sb.Append("<u>");

            if (isHighlight)
                sb.Append("<mark>");

            if (isSubscript)
                sb.Append("<sub>");
            else if (isSuperscript)
                sb.Append("<sup>");
        }

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);              
        }

        if (hasText)
        {
            if (isSubscript)
                sb.Append("</sub>");
            else if (isSuperscript)
                sb.Append("</sup>");

            if (isHighlight)
                sb.Append("</mark>");

            if (isUnderline)
                sb.Append("</u>");

            if (isStrikethrough)
                sb.Append("~~");

            if (isBold)
                sb.Append("**");

            if (isItalic)
                sb.Append("*");

            sb.Append(trailingSpaces);
        }
    }

    internal override void ProcessBreak(Break br, StringBuilder sb)
    {
        if (br.Type != null && br.Type == BreakValues.Page)
        {
            sb.AppendLine();
            sb.AppendLine("-----"); // rendered as horizontal rule
        }
        else
        {
            sb.AppendLine("  "); // soft break
        }
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        foreach(char c in text.InnerText.Trim())
        {
            if (_specialChars.Contains(c))
            {
                sb.Append('\\');
                sb.Append(c);
            }
            else if (c == '\r')
            {
                // Ignore as it's usually followed by \n
            }
            else if (c == '\n')
            {
                sb.AppendLine("  "); // soft break
            }
            else
            {
                sb.Append(c);
            }
        }
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        int rowCount = 0;
        foreach(var element in table.Elements())
        {
            switch (element)
            {
                case TableRow row:
                    if (rowCount == 0)
                    {
                        AddTableHeader(3, sb);
                    }
                    ProcessRow(row, sb);
                    ++rowCount;
                    break;
            }
        }
        sb.AppendLine();
        sb.AppendLine();
    }

    private void AddTableHeader(int columnCount, StringBuilder sb)
    {
        sb.Append("|");
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Append(" |");
        }
        sb.AppendLine();
        for (int i = 0; i < columnCount; ++i)
        {
            sb.Append("| --- ");
        }
        sb.AppendLine("|");
    }

    internal void ProcessRow(TableRow tableRow, StringBuilder sb)
    {
        sb.Append("| ");
        foreach (var element in tableRow.Elements())
        {
            switch (element)
            {
                case TableCell cell:
                    ProcessCell(cell, sb);
                    break;
            }
        }
        sb.AppendLine();
    }

    internal void ProcessCell(TableCell cell, StringBuilder sb)
    {
        var cellBuilder = new StringBuilder();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            // Join paragraphs as Markdown doesn't support multiple lines per cell
            if (paragraph != null)
                base.ProcessParagraph(paragraph, cellBuilder);

            cellBuilder.Append(' ');
        }
        sb.Append(cellBuilder.ToString());
        sb.Append(" | ");
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        var displayTextBuilder = new StringBuilder();
        foreach (var run in hyperlink.Elements<Run>())
        {
            if (run != null && run.GetFirstChild<Text>() is Text runText)
                ProcessText(runText, displayTextBuilder);

            displayTextBuilder.Append(' ');
        }
        if (hyperlink.Id?.Value is string rId)
        {
            var maindDocumentPart = OpenXmlHelpers.GetMainDocumentPart(hyperlink);
            if (maindDocumentPart?.HyperlinkRelationships.FirstOrDefault(x => x.Id == rId) is HyperlinkRelationship relationship)
            {
                string url = relationship.Uri.ToString();             
                sb.Append($" [{displayTextBuilder.ToString().Trim()}]({url}) ");
            }
        }
        //else if (hyperlink.Anchor?.Value is string anchor) // TODO
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)) && Directory.Exists(ImagesOutputFolder))
        {
            if (drawing.Descendants<DrawingML.Blip>().FirstOrDefault() is DrawingML.Blip blip &&
                blip.Embed?.Value is string relId)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
                ProcessImagePart(mainDocumentPart, relId, sb);
            }
        }
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        if ((!string.IsNullOrWhiteSpace(ImagesOutputFolder)) && Directory.Exists(ImagesOutputFolder))
        {
            if (picture.Descendants<ImageData>().FirstOrDefault() is ImageData imageData && 
                imageData.RelationshipId?.Value is string relId)
            {
                var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(picture);
                ProcessImagePart(mainDocumentPart, relId, sb);
            }
        }
    }

    internal void ProcessImagePart(MainDocumentPart? mainDocumentPart, string relId, StringBuilder sb)
    {
        if (mainDocumentPart?.GetPartById(relId!) is ImagePart imagePart)
        {
            string fileName = System.IO.Path.GetFileName(imagePart.Uri.OriginalString);
            string actualFilePath = ImagesOutputFolder+"/"+fileName;
            Uri uri;
            if (ImagesBaseUriOverride is null)
            {
                uri = new Uri(actualFilePath, UriKind.Absolute);
            }
            else
            {
                ImagesBaseUriOverride = ImagesBaseUriOverride.Trim('"');
                ImagesBaseUriOverride = ImagesBaseUriOverride.Replace('\\', '/');
                if (ImagesBaseUriOverride != string.Empty)
                {
                    if (ImagesBaseUriOverride.EndsWith("/") || ImagesBaseUriOverride.EndsWith("\\"))
                    {
                        ImagesBaseUriOverride = ImagesBaseUriOverride.Substring(0, ImagesBaseUriOverride.Length - 1);
                    }
                    ImagesBaseUriOverride += "/";
                    }
                ImagesBaseUriOverride += fileName;
                uri = new Uri(ImagesBaseUriOverride, UriKind.RelativeOrAbsolute);
            }

            using (var stream = imagePart.GetStream())
            using (var fileStream = new FileStream(actualFilePath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
            sb.Append(' ');
            sb.Append($"![{relId}]({uri})");
            sb.Append(' ');
        }
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmark, StringBuilder sb)
    {
        // TODO
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            string htmlEntity = string.Empty;
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                             out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    switch (symbolChar?.Font?.Value.ToLower())
                    {
                        case "wingdings":
                            htmlEntity = StringHelpers.WingdingsToUnicode((char)decimalValue);
                            break;
                        case "wingdings2":
                            htmlEntity = StringHelpers.Wingdings2ToUnicode((char)decimalValue);
                            break;
                        case "wingdings3":
                            htmlEntity = StringHelpers.Wingdings3ToUnicode((char)decimalValue);
                            break;
                        case "webdings":
                            htmlEntity = StringHelpers.WebdingsToUnicode((char)decimalValue);
                            break;
                    }
                }
            }
            if (string.IsNullOrWhiteSpace(htmlEntity))
            {
                htmlEntity = $"&#{decimalValue};";
            }
            sb.Append(htmlEntity);
        }        
    }

    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        // TODO
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmark, StringBuilder sb) { }
    internal override void ProcessFieldChar(FieldChar simpleField, StringBuilder sb) { }
    internal override void ProcessFieldCode(FieldCode simpleField, StringBuilder sb) { }
    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb) { }
    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb) { }
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb) { }
    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb) { }
    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb) { }
    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb) { }
    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb) { }
    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb) { }

}
}
