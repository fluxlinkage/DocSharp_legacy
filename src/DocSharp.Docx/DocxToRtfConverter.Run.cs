using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using DocSharp.Docx.Rtf;

namespace DocSharp.Docx {

public partial class DocxToRtfConverter
{
    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        if (!run.HasChildren)
            return;

        if (!isInField)
            sb.Append('{');

        ProcessRunFormatting(run, sb);
        sb.Append(' ');

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        if (!isInField)
            sb.Append('}');
    }

    internal void ProcessRunFormatting(Run run, StringBuilder sb)
    {
        var rtl = OpenXmlHelpers.GetEffectiveProperty<RightToLeftText>(run);
        if (rtl != null && (rtl.Val == null || rtl.Val))
        {
            sb.Append(@"\rtlch");
        }
        else
        {
            sb.Append(@"\ltrch");
        }

        var lang = OpenXmlHelpers.GetEffectiveProperty<Languages>(run);
        if (!string.IsNullOrEmpty(lang?.Val?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang.Val.Value);
            sb.Append(@"\lang" + code);
            sb.Append(@"\langnp" + code);
        }
        if (!string.IsNullOrEmpty(lang?.Bidi?.Value))
        {
            int code = RtfHelpers.GetLanguageCode(lang.Bidi.Value);
            sb.Append(@"\langfe" + code);
            sb.Append(@"\langfenp" + code);
        }

        if (OpenXmlHelpers.GetEffectiveProperty<NoProof>(run) is NoProof noProof)
        {
            if (noProof.Val == null || noProof.Val.Value)
            {
                sb.Append(@"\noproof\lang1024");
            }
        }

        // To be improved (Ascii value may not be present, although rare)
        string? font = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run)?.Ascii?.Value;
        if (!string.IsNullOrEmpty(font))
        {
            fonts.TryAddAndGetIndex(font, out int fontIndex);
            sb.Append($"\\f{fontIndex}");
        }
        else
        {
            // Calibri is already in the font table as last resort
            sb.Append(@"\f0");
        }

        string? color = OpenXmlHelpers.GetEffectiveProperty<Color>(run)?.Val;
        if ((!string.IsNullOrEmpty(color)) && 
             !color.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            colors.TryAddAndGetIndex(color, out int colorIndex);
            sb.Append($"\\cf{colorIndex}");
        }
        else
        {
            // If no color is specified, \cf0 is automatically handled by word processors.
            // Note: for this reason the color table uses 1-based index, while the font table should contain the f0 font.
            sb.Append(@"\cf0");
        }

        string? fontSize = OpenXmlHelpers.GetEffectiveProperty<FontSize>(run)?.Val;
        // Font size is in half-points in both DOCX and RTF
        if (int.TryParse(fontSize, out int fs))
        {
            sb.Append($"\\fs{fs}");
        }
        else
        {
            sb.Append($"\\fs{DefaultSettings.FontSize * 2}"); // Font size is in half-points
        }

        string? kerning = OpenXmlHelpers.GetEffectiveProperty<Kern>(run)?.Val;
        if (int.TryParse(kerning, out int k))
        {
            // Kerning is in half-points in both Open XML and RTF.
            sb.Append($"\\kerning{k}");
        }

        string? scaling = OpenXmlHelpers.GetEffectiveProperty<CharacterScale>(run)?.Val;
        if (int.TryParse(scaling, out int scale))
        {
            // Character scaling is expressed as percentage (100, 200, ...) in both Open XML and RTF.
            sb.Append($"\\charscalex{scale}");
        }

        string? fitText = OpenXmlHelpers.GetEffectiveProperty<FitText>(run)?.Val;
        if (int.TryParse(fitText, out int ft))
        {
            // FitText is in twips in both Open XML and RTF.
            sb.Append($"\\fittext{ft}");
        }

        string? spacing = OpenXmlHelpers.GetEffectiveProperty<Spacing>(run)?.Val;
        if (int.TryParse(spacing, out int sp))
        {
            // Character spacing is expressed in twips in Open XML;
            // in RTF it should also be specified in quarter-points for backward compatibility.
            sb.Append($"\\expnd{sp * 5}");
            sb.Append($"\\expndtw{sp}");
        }

        var bold = OpenXmlHelpers.GetEffectiveProperty<Bold>(run);
        // Formatting options such as bold are considered enabled if the element is present,
        // unless OnOffValue is explicitly set to false.
        // (e.g. <w:b /> without value means bold is enabled, otherwise it would not be present at all)
        if (bold != null && (bold.Val is null || bold.Val)) 
        {
            sb.Append(@"\b");
        }

        var italic = OpenXmlHelpers.GetEffectiveProperty<Italic>(run);
        if (italic != null && (italic.Val is null || italic.Val))
        {
            sb.Append(@"\i");
        }

        var underline = OpenXmlHelpers.GetEffectiveProperty<Underline>(run);
        if (underline?.Val != null)
        {
            string? ul = RtfUnderlineMapper.GetUnderlineType(underline.Val);
            if (!string.IsNullOrEmpty(ul))
            {
                sb.Append(ul);
            }

            if ((!string.IsNullOrEmpty(underline.Color?.Value)) && 
                !underline.Color.Value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                colors.TryAddAndGetIndex(underline.Color.Value, out int colorIndex);
                sb.Append($"\\ulc{colorIndex}");
            }
        }

        var doubleStrike = OpenXmlHelpers.GetEffectiveProperty<DoubleStrike>(run);
        if (doubleStrike != null && (doubleStrike.Val is null || doubleStrike.Val))
        {
            sb.Append(@"\striked1");
        }
        else
        {
            // Don't add strike if double strike is already active.
            var strike = OpenXmlHelpers.GetEffectiveProperty<Strike>(run);
            if (strike != null && (strike.Val is null || strike.Val))
            {
                sb.Append(@"\strike");
            }
        }

        var highlight = OpenXmlHelpers.GetEffectiveProperty<Highlight>(run);
        if (highlight?.Val != null)
        {
            if (highlight.Val == HighlightColorValues.None)
            {
                sb.Append(@"\highlight0");
            }
            else
            {
                string? hex = RtfHighlightMapper.GetHexColor(highlight.Val);
                if (!string.IsNullOrEmpty(hex))
                {
                    colors.TryAddAndGetIndex(hex, out int highlightIndex);
                    sb.Append($"\\highlight{highlightIndex}");
                }
            }
        }

        var verticalTextAlignment = OpenXmlHelpers.GetEffectiveProperty<VerticalTextAlignment>(run);
        if (verticalTextAlignment?.Val != null)
        {
            if (verticalTextAlignment.Val == VerticalPositionValues.Subscript)
            {
                sb.Append(@"\sub");
            }
            else if (verticalTextAlignment.Val == VerticalPositionValues.Superscript)
            {
                sb.Append(@"\super");
            }
            else
            {
                sb.Append(@"\nosupersub");
            }
        }
        else
        {
            var position = OpenXmlHelpers.GetEffectiveProperty<Position>(run);
            if (position?.Val != null && int.TryParse(position.Val.Value, out int pos))
            {
                if (pos < 0)
                {
                    sb.Append($"\\dn{pos}");
                }
                else if (pos > 0) 
                {
                    sb.Append($"\\up{pos}");
                }
            }
        }

        var em = OpenXmlHelpers.GetEffectiveProperty<Emphasis>(run);
        if (em?.Val != null)
        {
            if (em.Val == EmphasisMarkValues.None)
            {
                sb.Append(@"\accnone");
            }
            else if (em.Val == EmphasisMarkValues.Circle)
            {
                sb.Append(@"\acccircle");
            }
            else if (em.Val == EmphasisMarkValues.Comma)
            {
                sb.Append(@"\acccomma");
            }
            else if (em.Val == EmphasisMarkValues.Dot)
            {
                sb.Append(@"\accdot");
            }
            else if (em.Val == EmphasisMarkValues.UnderDot)
            {
                sb.Append(@"\accunderdot");
            }
        }

        var smallCaps = OpenXmlHelpers.GetEffectiveProperty<SmallCaps>(run);
        if (smallCaps != null && (smallCaps.Val is null || smallCaps.Val))
        {
            sb.Append(@"\scaps");
        }
        else
        {
            // Small caps and All caps are mutually exclusive
            var allCaps = OpenXmlHelpers.GetEffectiveProperty<Caps>(run);
            if (allCaps != null && (allCaps.Val is null || allCaps.Val))
            {
                sb.Append(@"\caps");
            }
        }

        var emboss = OpenXmlHelpers.GetEffectiveProperty<Emboss>(run);
        if (emboss != null && (emboss.Val is null || emboss.Val))
        {
            sb.Append(@"\embo");
        }

        var engrave = OpenXmlHelpers.GetEffectiveProperty<Imprint>(run);
        if (engrave != null && (engrave.Val is null || engrave.Val))
        {
            sb.Append(@"\impr");
        }

        // RTF does not support advanced shadow and outline effects introduced with Office 2010,
        // so they are converted to the legacy font effect.
        var shadow = OpenXmlHelpers.GetEffectiveProperty<Shadow>(run);
        if ((shadow != null && (shadow.Val is null || shadow.Val)) ||
             OpenXmlHelpers.GetEffectiveProperty<Shadow14>(run) != null)
        {
            sb.Append(@"\shad");
        }

        var outline = OpenXmlHelpers.GetEffectiveProperty<Outline>(run);        
        if ((outline != null && (outline.Val is null || outline.Val)) ||
             OpenXmlHelpers.GetEffectiveProperty<Outline14>(run) != null)
        {
            sb.Append(@"\outl");
        }

        var hidden = OpenXmlHelpers.GetEffectiveProperty<Vanish>(run);
        if (hidden != null && (hidden.Val is null || hidden.Val))
        {
            sb.Append(@"\v");
        }

        var border = OpenXmlHelpers.GetEffectiveProperty<Border>(run);
        if (border != null)
        {
            sb.Append(@"\chbrdr");
            ProcessBorder(border, sb);
        }

        var shading = OpenXmlHelpers.GetEffectiveProperty<Shading>(run);
        if (shading != null)
        {
            ProcessShading(shading, sb, ShadingType.Character);
        }

        var snapToGrid = OpenXmlHelpers.GetEffectiveProperty<SnapToGrid>(run);
        if (snapToGrid?.Val != null && !snapToGrid.Val) // True by default
        {
            sb.Append(@"\cgrid0");
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value) && 
            !string.IsNullOrEmpty(symbolChar?.Font?.Value))
        {
            fonts.TryAddAndGetIndex(symbolChar.Font.Value, out int fontIndex);
            sb.Append('{');
            sb.Append($"\\f{fontIndex} ");
            sb.AppendRtfUnicodeChar(symbolChar.Char.Value);
            sb.Append('}');
        }
    }
}
}
