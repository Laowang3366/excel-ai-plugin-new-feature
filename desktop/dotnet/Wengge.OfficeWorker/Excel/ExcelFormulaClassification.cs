using System.Text.RegularExpressions;

namespace Wengge.OfficeWorker.Excel;

internal enum ExcelFormulaKind
{
    Plain,
    Dynamic,
    LegacyArray,
}

internal static class ExcelFormulaClassification
{
    private static readonly Regex FunctionCallHeadRegex = new(
        @"^\s*(?<name>[A-Za-z_][A-Za-z0-9_.]*)\s*\(",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex RangeReferenceRegex = new(
        @"(?<![A-Za-z0-9_.])(?:\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?\d+\s*:\s*\$?\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HashSet<string> LegacySpillFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "TRANSPOSE", "FREQUENCY", "MMULT", "LINEST", "LOGEST", "TREND", "GROWTH",
    };

    private static readonly HashSet<string> ScalarRangeFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
        "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MIN", "MAX", "MEDIAN", "PRODUCT", "SUBTOTAL",
        "AGGREGATE", "LOOKUP", "VLOOKUP", "HLOOKUP", "MATCH", "INDEX", "AND", "OR",
    };

    private static readonly IReadOnlyDictionary<string, string> FunctionPrefixes =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["LET"] = "_xlfn.", ["LAMBDA"] = "_xlfn.", ["XLOOKUP"] = "_xlfn.", ["XMATCH"] = "_xlfn.",
            ["TAKE"] = "_xlfn.", ["DROP"] = "_xlfn.", ["CHOOSEROWS"] = "_xlfn.", ["CHOOSECOLS"] = "_xlfn.",
            ["MAP"] = "_xlfn.", ["REDUCE"] = "_xlfn.", ["SCAN"] = "_xlfn.", ["HSTACK"] = "_xlfn.",
            ["VSTACK"] = "_xlfn.", ["MAKEARRAY"] = "_xlfn.", ["BYROW"] = "_xlfn.", ["BYCOL"] = "_xlfn.",
            ["TEXTSPLIT"] = "_xlfn.", ["TEXTBEFORE"] = "_xlfn.", ["TEXTAFTER"] = "_xlfn.",
            ["FILTER"] = "_xlfn._xlws.", ["SORT"] = "_xlfn._xlws.", ["SORTBY"] = "_xlfn._xlws.",
            ["UNIQUE"] = "_xlfn._xlws.", ["SEQUENCE"] = "_xlfn._xlws.", ["RANDARRAY"] = "_xlfn._xlws.",
            ["TOCOL"] = "_xlfn._xlws.", ["TOROW"] = "_xlfn._xlws.", ["WRAPROWS"] = "_xlfn._xlws.",
            ["WRAPCOLS"] = "_xlfn._xlws.", ["GROUPBY"] = "_xlfn._xlws.", ["PIVOTBY"] = "_xlfn._xlws.",
            ["EXPAND"] = "_xlfn._xlws.",
        };

    public static bool IsFormula(string? text) => text is { Length: > 1 } && text[0] == '=';

    public static ExcelFormulaKind Classify(string formula, bool legacyCse = false)
    {
        if (!IsFormula(formula)) throw new ArgumentException("公式必须以 '=' 开头", nameof(formula));
        if (legacyCse) return ExcelFormulaKind.LegacyArray;
        return IsDynamicArray(formula)
            ? ExcelFormulaKind.Dynamic
            : ExcelFormulaKind.Plain;
    }

    public static bool IsDynamicArray(string? formula, bool forceLegacyArray = false) =>
        !forceLegacyArray && IsFormula(formula) &&
        (ContainsModernFunctionCall(formula!) || ContainsLegacySpillFunction(formula!) || ContainsArrayExpression(formula!));

    public static string? LeadingFunction(string? formula)
    {
        if (!IsFormula(formula)) return null;
        var match = FunctionCallHeadRegex.Match(formula![1..]);
        if (!match.Success) return null;
        var qualified = match.Groups["name"].Value;
        var bare = qualified[(qualified.LastIndexOf('.') + 1)..];
        return bare.ToUpperInvariant();
    }

    public static string NormalizeForOpenXml(string formula)
    {
        if (!IsFormula(formula)) throw new ArgumentException("公式必须以 '=' 开头", nameof(formula));
        var body = formula[1..];
        var output = new System.Text.StringBuilder(body.Length + 16);
        for (var index = 0; index < body.Length;)
        {
            if (body[index] == '"')
            {
                CopyStringLiteral(body, output, ref index);
                continue;
            }
            if (body[index] == '\'')
            {
                CopyQuotedSheetName(body, output, ref index);
                continue;
            }
            if (!IsIdentifierStart(body[index]))
            {
                output.Append(body[index++]);
                continue;
            }

            var start = index++;
            while (index < body.Length && IsIdentifierPart(body[index])) index++;
            var qualified = body[start..index];
            var next = index;
            while (next < body.Length && char.IsWhiteSpace(body[next])) next++;
            var function = qualified[(qualified.LastIndexOf('.') + 1)..];
            if (next < body.Length && body[next] == '(' &&
                !qualified.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) &&
                FunctionPrefixes.TryGetValue(function, out var prefix))
                output.Append(prefix).Append(function);
            else
                output.Append(qualified);
        }
        return output.ToString();
    }

    private static bool ContainsModernFunctionCall(string formula)
    {
        var body = formula[1..];
        for (var index = 0; index < body.Length;)
        {
            if (body[index] == '"')
            {
                SkipStringLiteral(body, ref index);
                continue;
            }
            if (body[index] == '\'')
            {
                SkipQuotedSheetName(body, ref index);
                continue;
            }
            if (!IsIdentifierStart(body[index]))
            {
                index++;
                continue;
            }

            var start = index++;
            while (index < body.Length && IsIdentifierPart(body[index])) index++;
            var qualified = body[start..index];
            var next = index;
            while (next < body.Length && char.IsWhiteSpace(body[next])) next++;
            var function = qualified[(qualified.LastIndexOf('.') + 1)..];
            if (next < body.Length && body[next] == '(' && FunctionPrefixes.ContainsKey(function)) return true;
        }
        return false;
    }

    private static bool ContainsLegacySpillFunction(string formula) =>
        LeadingFunction(formula) is { } function && LegacySpillFunctions.Contains(function);

    private static bool ContainsArrayExpression(string formula)
    {
        var searchable = RemoveQuotedContent(formula[1..]);
        if (!RangeReferenceRegex.IsMatch(searchable)) return false;
        var leadingFunction = LeadingFunction(formula);
        return leadingFunction is null || !ScalarRangeFunctions.Contains(leadingFunction);
    }

    private static string RemoveQuotedContent(string text)
    {
        var output = text.ToCharArray();
        for (var index = 0; index < output.Length;)
        {
            if (output[index] is not ('"' or '\''))
            {
                index++;
                continue;
            }
            var quote = output[index];
            output[index++] = ' ';
            while (index < output.Length)
            {
                var current = output[index];
                output[index++] = ' ';
                if (current != quote) continue;
                if (index < output.Length && output[index] == quote)
                {
                    output[index++] = ' ';
                    continue;
                }
                break;
            }
        }
        return new string(output);
    }

    private static void CopyStringLiteral(string text, System.Text.StringBuilder output, ref int index)
    {
        output.Append(text[index++]);
        while (index < text.Length)
        {
            var current = text[index++];
            output.Append(current);
            if (current != '"') continue;
            if (index < text.Length && text[index] == '"')
            {
                output.Append(text[index++]);
                continue;
            }
            return;
        }
    }

    private static void CopyQuotedSheetName(string text, System.Text.StringBuilder output, ref int index)
    {
        output.Append(text[index++]);
        while (index < text.Length)
        {
            var current = text[index++];
            output.Append(current);
            if (current != '\'') continue;
            if (index < text.Length && text[index] == '\'')
            {
                output.Append(text[index++]);
                continue;
            }
            return;
        }
    }

    private static void SkipStringLiteral(string text, ref int index)
    {
        index++;
        while (index < text.Length)
        {
            if (text[index++] != '"') continue;
            if (index < text.Length && text[index] == '"')
            {
                index++;
                continue;
            }
            return;
        }
    }

    private static void SkipQuotedSheetName(string text, ref int index)
    {
        index++;
        while (index < text.Length)
        {
            if (text[index++] != '\'') continue;
            if (index < text.Length && text[index] == '\'')
            {
                index++;
                continue;
            }
            return;
        }
    }

    private static bool IsIdentifierStart(char value) => char.IsLetter(value) || value == '_';
    private static bool IsIdentifierPart(char value) => char.IsLetterOrDigit(value) || value is '_' or '.';
}
