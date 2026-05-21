using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.Objects;
using ACadSharp.Tables;
using CSMath;

namespace DwgLayoutCreator;

public static class LayoutFactory
{
    // Plotter i style sheet hardcoded — u standalone modu nema PSV validacije,
    // ali AutoCAD ovo prepoznaje kad otvori DWG.
    private const string PlotterName = "DWG To PDF.pc3";
    private const string StyleSheet  = "monochrome.ctb";

    // Margins (mm): left, bottom, right, top. Korisnikova konfiguracija je 25mm uvez + 5mm ostalo.
    // U standalone modu nemamo PSV pa hardcode margine — AutoCAD će prepoznati svoj PMP pri otvaranju.
    private const double MarginLeft   = 25.0;
    private const double MarginBottom = 5.0;
    private const double MarginRight  = 5.0;
    private const double MarginTop    = 5.0;

    // Kreira standardni ISO layout (A4P/A4L/.../A0P).
    // Vraća kreirani Layout (s AssociatedBlock-om u koji se inserira sastavnica).
    public static Layout CreateStandardLayout(
        CadDocument doc,
        BlockScanResult result,
        PaperSizeInfo info)
    {
        // 1. Kreiraj Layout — ctor auto-kreira AssociatedBlock (paper space)
        var layout = new Layout(result.LayoutName);

        // 2. Plot settings (svi flat na Layout-u, naslijeđeno iz PlotSettings)
        layout.SystemPrinterName = PlotterName;
        layout.PaperSize         = info.PaperSize;
        layout.PaperUnits        = PlotPaperUnits.Millimeters;
        layout.PaperRotation     = PlotRotation.NoRotation;  // paper name već kodira orijentaciju
        layout.PaperWidth        = info.PhysW;
        layout.PaperHeight       = info.PhysH;
        layout.PlotType          = PlotType.LayoutInformation;
        layout.StandardScale     = 1.0;
        layout.NumeratorScale    = 1.0;
        layout.DenominatorScale  = 1.0;
        layout.StyleSheet        = StyleSheet;
        layout.UnprintableMargin = new PaperMargin(MarginLeft, MarginBottom, MarginRight, MarginTop);
        layout.PageName          = info.PaperSize;  // za prikaz u Page Setup Manager-u
        layout.Flags             = PlotFlags.UseStandardScale
                                 | PlotFlags.ShowPlotStyles      // = ToggleDisplayPlotStyle ON
                                 | PlotFlags.PlotPlotStyles
                                 | PlotFlags.PrintLineweights
                                 | PlotFlags.DrawViewportsFirst;

        // 3. Dodaj u dokument (BEZ ovog Layout ostaje "orphaned")
        doc.Layouts.Add(layout);

        // 4. Briši default viewport koji je auto-kreiran u AssociatedBlock-u
        var defaultVps = layout.AssociatedBlock.Entities.OfType<Viewport>().ToList();
        foreach (var vp in defaultVps)
            layout.AssociatedBlock.Entities.Remove(vp);

        // 5. Kreiraj novi viewport koji popunjava printable area
        AddPrintableAreaViewport(layout, info, result);

        return layout;
    }

    // Custom layout: kopija postojećeg "template" layouta CUSTOM1-8.
    // U standalone modu nemamo LayoutManager.CopyLayout — pa ručno kopiramo PlotSettings.
    public static Layout CreateCustomLayout(
        CadDocument doc,
        BlockScanResult result,
        Layout template)
    {
        var layout = new Layout(result.LayoutName);

        // Kopiraj plot settings iz template-a
        layout.SystemPrinterName = template.SystemPrinterName;
        layout.PaperSize         = template.PaperSize;
        layout.PaperUnits        = template.PaperUnits;
        layout.PaperRotation     = template.PaperRotation;
        layout.PaperWidth        = template.PaperWidth;
        layout.PaperHeight       = template.PaperHeight;
        layout.PlotType          = template.PlotType;
        layout.StandardScale     = template.StandardScale;
        layout.NumeratorScale    = template.NumeratorScale;
        layout.DenominatorScale  = template.DenominatorScale;
        layout.StyleSheet        = template.StyleSheet;
        layout.UnprintableMargin = template.UnprintableMargin;
        layout.Flags             = template.Flags;

        doc.Layouts.Add(layout);

        // Briši default viewport
        var defaultVps = layout.AssociatedBlock.Entities.OfType<Viewport>().ToList();
        foreach (var vp in defaultVps)
            layout.AssociatedBlock.Entities.Remove(vp);

        // Za custom layoute viewport veličina dolazi iz bounding boxa bloka
        double modelW = Math.Abs(result.MaxPt.X - result.MinPt.X);
        double modelH = Math.Abs(result.MaxPt.Y - result.MinPt.Y);

        var dummyInfo = new PaperSizeInfo(
            template.PaperSize, "Landscape",
            template.PaperWidth, template.PaperHeight, 0.0);
        AddPrintableAreaViewport(layout, dummyInfo, result);

        return layout;
    }

    // Dodaje viewport veličine printable area (paper minus margine).
    // U paper-space koordinatama, (0,0) je donji-lijevi printable area kut.
    public static void AddPrintableAreaViewport(
        Layout layout,
        PaperSizeInfo info,
        BlockScanResult result)
    {
        double vpW  = info.PhysW - MarginLeft - MarginRight;
        double vpH  = info.PhysH - MarginBottom - MarginTop;
        double vpCx = vpW / 2.0;
        double vpCy = vpH / 2.0;

        // Model space view: zoomira na block extents
        double modelW = Math.Max(result.MaxPt.X - result.MinPt.X, 1.0);
        double modelH = Math.Max(result.MaxPt.Y - result.MinPt.Y, 1.0);
        double vpAspect = vpW / vpH;
        double mAspect  = modelW / modelH;
        double viewH    = mAspect > vpAspect ? modelW / vpAspect : modelH;

        var vp = new Viewport
        {
            Center     = new XYZ(vpCx, vpCy, 0),
            Width      = vpW,
            Height     = vpH,
            ViewCenter = new XY(
                (result.MinPt.X + result.MaxPt.X) / 2.0,
                (result.MinPt.Y + result.MaxPt.Y) / 2.0),
            ViewHeight = viewH,
        };

        layout.AddViewport(vp);
    }

    // Vraća printable area (W, H) iz info i hardcoded margina.
    public static (double w, double h) GetPrintableArea(PaperSizeInfo info)
        => (info.PhysW - MarginLeft - MarginRight,
            info.PhysH - MarginBottom - MarginTop);
}
