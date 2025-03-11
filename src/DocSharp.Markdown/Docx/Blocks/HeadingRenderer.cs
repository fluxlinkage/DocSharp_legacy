using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using DocSharp.Markdown;

namespace Markdig.Renderers.Docx.Blocks {

public class HeadingRenderer : LeafBlockParagraphRendererBase<HeadingBlock>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, HeadingBlock obj)
    {
        var styleId = renderer.Styles.MarkdownStyles["UndefinedHeading"];
        if (renderer.Styles.Headings.ContainsKey(obj.Level))
        {
            styleId = renderer.Styles.Headings[obj.Level];
        }

            string? bookmarkName = null;
        if (obj.Inline?.FindDescendants<LiteralInline>().FirstOrDefault() is LiteralInline literal)
        {
            bookmarkName = MarkdownUtils.GetBookmarkName(literal.Content.ToString());
        }

        WriteAsParagraph(renderer, obj, styleId, bookmarkName);
    }
}
}
