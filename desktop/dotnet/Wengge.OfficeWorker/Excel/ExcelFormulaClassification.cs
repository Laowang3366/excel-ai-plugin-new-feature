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
        return ContainsModernFunctionCall(formula)
            ? ExcelFormulaKind.Dynamic
            : ExcelFormulaKind.Plain;
    }

    public static bool IsDynamicArray(string? formula, bool forceLegacyArray = false) =>
        !forceLegacyArray && IsFormula(formula) && Classify(formula!) == ExcelFormulaKind.Dynamic;

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
