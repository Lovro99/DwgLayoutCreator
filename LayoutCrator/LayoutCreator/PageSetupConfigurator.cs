using System.Collections.Specialized;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace LayoutCreator;

public static class PageSetupConfigurator
{
    private const string PreferredPlotter = "DWG To PDF.pc3";
    private const string StyleSheet       = "monochrome.ctb";

    // Configures PlotSettings on the given layout (must be open ForWrite).
    // Returns a description string for logging.
    //
    // KEY FIX: GetCanonicalMediaNameList returns CANONICAL names (e.g. "ISO_A4_(210.00_x_297.00_MM)")
    // NOT display names (e.g. "ISO A4 (210.00 x 297.00 MM)").
    // We use GetLocaleMediaName() to convert canonical → display, then compare against
    // info.PaperName (which holds the display name from LayoutDefinitions).
    public static string Configure(Layout layout, PaperSizeInfo info, Editor ed)
    {
        var psv = PlotSettingsValidator.Current;

        // Step A: resolve and set plotter
        string plotter = ResolvePlotter(psv, ed);
        ed.WriteMessage($"\n    A) Plotter: '{plotter}'");
        psv.SetPlotConfigurationName(layout, plotter, null);
        psv.RefreshLists(layout);

        // Step B: find canonical media name by matching display (locale) names
        string canonical = FindCanonicalByLocaleName(psv, layout, info.PaperName, ed);
        ed.WriteMessage($"\n    B) Paper canonical: '{canonical}'");
        psv.SetCanonicalMediaName(layout, canonical);

        // Step C: units, plot type, scale, origin
        psv.SetPlotPaperUnits(layout, Autodesk.AutoCAD.DatabaseServices.PlotPaperUnit.Millimeters);
        psv.SetPlotType(layout, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
        psv.SetUseStandardScale(layout, true);
        psv.SetStdScaleType(layout, Autodesk.AutoCAD.DatabaseServices.StdScaleType.StdScale1To1);
        psv.SetPlotOrigin(layout, new Autodesk.AutoCAD.Geometry.Point2d(0, 0));
        ed.WriteMessage($"\n    C) Units/Type/Scale set.");

        // Step D: plot style table
        try
        {
            psv.SetCurrentStyleSheet(layout, StyleSheet);
            ed.WriteMessage($"\n    D) StyleSheet: '{StyleSheet}'");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    D) StyleSheet '{StyleSheet}' not found ({ex.Message}) — skipping.");
        }

        // Step E: orientation
        // Paper names in LayoutDefinitions already encode the orientation:
        //   A4P → "ISO A4 (210.00 x 297.00 MM)" (native portrait, 210x297)
        //   A4L → "ISO A4 (297.00 x 210.00 MM)" (native landscape, 297x210)
        // So PlotRotation is always Degrees000 (no rotation relative to paper).
        // Setting Degrees090 here would rotate the drawing 90° AGAINST the paper,
        // which is what caused the previous "orientation wrong" bug.
        psv.SetPlotRotation(layout,
            Autodesk.AutoCAD.DatabaseServices.PlotRotation.Degrees000);
        ed.WriteMessage($"\n    E) Orientation: {info.Orientation} (paper-native, no rotation)");

        // Step F: show plot styles (equivalent to LISP ToggleDisplayPlotStyle → always ON)
        layout.ShowPlotStyles = true;

        return $"{plotter} / {canonical}";
    }

    // Finds the best available PDF plotter.
    private static string ResolvePlotter(PlotSettingsValidator psv, Editor ed)
    {
        StringCollection devices = psv.GetPlotDeviceList();

        foreach (string? item in devices)
            if (item is string d && d.Equals(PreferredPlotter, StringComparison.OrdinalIgnoreCase))
                return d;

        foreach (string? item in devices)
            if (item is string d && d.Contains("PDF", StringComparison.OrdinalIgnoreCase))
                return d;

        foreach (string? item in devices)
            if (item is string d && !d.Equals("None", StringComparison.OrdinalIgnoreCase))
                return d;

        return PreferredPlotter;
    }

    // Finds the CANONICAL media name whose locale/display name matches the requested display name.
    // GetCanonicalMediaNameList returns internal names (e.g. "ISO_A4_(210.00_x_297.00_MM)")
    // GetLocaleMediaName converts each canonical → display name for comparison.
    private static string FindCanonicalByLocaleName(
        PlotSettingsValidator psv, Layout layout, string requestedLocale, Editor ed)
    {
        StringCollection canonicalList = psv.GetCanonicalMediaNameList(layout);

        // Pass 1: exact locale match
        foreach (string? item in canonicalList)
        {
            if (item is not string cn) continue;
            try
            {
                string locale = psv.GetLocaleMediaName(layout, cn);
                if (locale.Equals(requestedLocale, StringComparison.OrdinalIgnoreCase))
                    return cn;
            }
            catch { /* some canonical names can't be converted */ }
        }

        // Pass 2: also try direct canonical match (in case caller passed canonical)
        foreach (string? item in canonicalList)
            if (item is string cn && cn.Equals(requestedLocale, StringComparison.OrdinalIgnoreCase))
                return cn;

        // Pass 3: partial match on locale name (all words must appear)
        string[] parts = requestedLocale.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (string? item in canonicalList)
        {
            if (item is not string cn) continue;
            try
            {
                string locale = psv.GetLocaleMediaName(layout, cn);
                if (parts.All(p => locale.Contains(p, StringComparison.OrdinalIgnoreCase)))
                    return cn;
            }
            catch { }
        }

        // Pass 4: ISO size key fallback (e.g. "A4", "A3")
        string sizeKey = parts.FirstOrDefault(p => p.Length == 2 && p[0] == 'A' && char.IsDigit(p[1])) ?? "";
        if (!string.IsNullOrEmpty(sizeKey))
        {
            foreach (string? item in canonicalList)
                if (item is string cn && cn.Contains(sizeKey, StringComparison.OrdinalIgnoreCase))
                    return cn;
        }

        ed.WriteMessage($"\n    WARNING: no media match for '{requestedLocale}' — using as-is (may create custom paper)");
        return requestedLocale;
    }

    // Diagnostic: lists all plotters and their first 30 media names (both canonical + locale).
    // Run LISTMEDIA in AutoCAD.
    public static void ListAvailableMedia(Editor ed)
    {
        var psv = PlotSettingsValidator.Current;

        StringCollection devices = psv.GetPlotDeviceList();
        ed.WriteMessage("\n=== Available plotters ===");
        foreach (string? item in devices)
            if (item is string d) ed.WriteMessage($"\n  '{d}'");

        // Use a temp PlotSettings object to query media
        using var tempPs = new Layout();

        string preferred = PreferredPlotter;
        // Try to find DWG To PDF.pc3 or any PDF plotter
        foreach (string? item in devices)
            if (item is string d && d.Equals(PreferredPlotter, StringComparison.OrdinalIgnoreCase))
                { preferred = d; break; }

        psv.SetPlotConfigurationName(tempPs, preferred, null);
        psv.RefreshLists(tempPs);

        StringCollection media = psv.GetCanonicalMediaNameList(tempPs);
        ed.WriteMessage($"\n\n=== Media for '{preferred}' (canonical → locale) ===");
        int i = 0;
        foreach (string? item in media)
        {
            if (item is not string cn) continue;
            string locale = "(err)";
            try { locale = psv.GetLocaleMediaName(tempPs, cn); } catch { }
            ed.WriteMessage($"\n  canonical: '{cn}'");
            ed.WriteMessage($"\n    locale:   '{locale}'");
            if (++i >= 30) { ed.WriteMessage("\n  ... (truncated)"); break; }
        }
    }
}
