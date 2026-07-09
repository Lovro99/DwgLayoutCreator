using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace LayoutCreator;

// Port ele:nums / ele:valid-layout-p / ele:layout< iz
// AutoLisp/ExportLayoutsToExcel.lsp. Koriste ju i SORTTABSBATCH (poredak tabova)
// i VERIFYBATCH (ocekivani poredak).
//
// VAZNO: logika MORA ostati IDENTICNA Python modulu batch/layout_names.py —
// mijenjaj ih SAMO zajedno. Oba su port istih ele: funkcija.
//
// Kompajlira se u LayoutCreatorCore (a i u UI plugin ne smeta jer nema AutoCAD
// ovisnosti); trenutno ju referenciraju samo core komande.
internal static class LayoutNameOrder
{
    // Uzorak valjanog imena: ^[0-9]+-[0-9]+x[0-9]+$ (malo 'x'; striktno ASCII
    // znamenke, jer je ele:all-digits-p striktno 0-9, a vl-string-search je
    // case-sensitive pa 'x' ne smije biti veliko).
    private static readonly Regex ValidRe =
        new("^[0-9]+-[0-9]+x[0-9]+$", RegexOptions.Compiled);

    // Port ele:valid-layout-p.
    internal static bool IsValidLayout(string? name) =>
        name != null && ValidRe.IsMatch(name);

    // Port ele:digit-p: (< 47 c 58) — striktno ASCII 0-9.
    private static bool IsDigit(int code) => code >= 48 && code <= 57;

    // Port ele:nums-char. a = kod prethodnog znaka, b = tekuci, c = sljedeci
    // (a/c = -1 na rubovima). Zadrzi b ako je znamenka, vodeci minus, ili
    // decimalna tocka izmedu dvije znamenke; inace zamijeni razmakom (32).
    private static int NumsChar(int a, int b, int c)
    {
        if (b >= 48 && b <= 57) return b;
        if (b == 45 && IsDigit(c) && !IsDigit(a)) return b;   // vodeci '-'
        if (b == 46 && IsDigit(a) && IsDigit(c)) return b;    // '.' izmedu znamenki
        return 32;
    }

    // Port ele:nums — izvlaci brojeve, "10-2x3" -> [10, 2, 3].
    internal static List<double> Nums(string text)
    {
        var sb = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            int a = i > 0 ? text[i - 1] : -1;
            int b = text[i];
            int c = i < text.Length - 1 ? text[i + 1] : -1;
            sb.Append((char)NumsChar(a, b, c));
        }
        var result = new List<double>();
        foreach (var tok in sb.ToString().Split(
                     (char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            if (double.TryParse(tok, NumberStyles.Float,
                                CultureInfo.InvariantCulture, out double d))
                result.Add(d);
        }
        return result;
    }

    // Kljuc za stabilno sortiranje: (prvi broj, drugi broj). Treci se ignorira —
    // vjerno ele:layout< koji usporeduje samo (car) pa (cadr).
    internal static (double first, double second) SortKey(string name)
    {
        var n = Nums(name);
        double first = n.Count > 0 ? n[0] : 0.0;
        double second = n.Count > 1 ? n[1] : 0.0;
        return (first, second);
    }

    // Filtriraj na valjana imena i STABILNO sortiraj po (prvi, drugi).
    // LINQ OrderBy/ThenBy je stabilan (plan §4.3).
    internal static List<string> SortValid(IEnumerable<string> names)
    {
        return names.Where(IsValidLayout)
                    .OrderBy(n => SortKey(n).first)
                    .ThenBy(n => SortKey(n).second)
                    .ToList();
    }
}
