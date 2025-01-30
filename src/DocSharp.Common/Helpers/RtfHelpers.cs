using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DocSharp.Helpers;

public static class RtfHelpers
{
    public static void AppendRtfEscaped(this StringBuilder sb, string value)
    {
        foreach (char c in value)
        {
            if (c == '\\' || c == '{' || c == '}')
            {
                sb.Append(new string(['\\', c]));
            }
            else if (c == '\t')
            {
                sb.Append("\\tab ");
            }
            else if (c == '\f')
            {
                sb.Append("\\page ");
            }
            else if (c == '\r')
            {
                // Ignore as it's usually followed by \n
            }
            else if (c == '\n')
            {
                sb.Append("\\line ");
            }
            else if (c < 32 || c > 127)
            {
                sb.AppendFormat("\\u{0}?", (int)c);
            }
            else
            {
                sb.Append(c);
            }
        }
    }

    public static string? ConvertToRtfColor(string hexColor)
    {
        hexColor = hexColor.TrimStart('#').ToLower();
        int length = hexColor.Length;
        switch (length)
        {
            case 3:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(0, 1) + hexColor.Substring(0, 1), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(1, 1) + hexColor.Substring(1, 1), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(2, 2) + hexColor.Substring(2, 2), 16)};";
            case 6:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(0, 2), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(2, 2), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(4, 2), 16)};";
            case 8:
                return $"\\red{System.Convert.ToInt32(hexColor.Substring(2, 2), 16)}" +
                          $"\\green{System.Convert.ToInt32(hexColor.Substring(4, 2), 16)}" +
                          $"\\blue{System.Convert.ToInt32(hexColor.Substring(6, 2), 16)};";
            default:
                // Unknown format
                return null;
        }
    }

    public static int GetLanguageCode(string langId)
    {
        try
        {
            var culture = new CultureInfo(langId);
            return culture.LCID;
        }
        catch
        {
            return 1024; // None/unspecified
        }
    }
}
