using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LayoutCreator;

public static class LayoutBuilder
{
    public static void CreateStandardLayout(
        Database db,
        BlockScanResult result,
        string sastPath,
        Editor ed)
    {
        if (!LayoutDefinitions.StandardSizes.TryGetValue(result.BlockName, out var info))
            throw new ArgumentException($"No paper size definition for block: {result.BlockName}");

        ed.WriteMessage($"\n  [1/5] Importing title block...");
        TitleBlockInserter.EnsureImported(db, sastPath);

        ed.WriteMessage($"\n  [2/5] Creating layout '{result.LayoutName}'...");
        var lm       = LayoutManager.Current;
        ObjectId lid = lm.CreateLayout(result.LayoutName);
        if (lid.IsNull) throw new InvalidOperationException("CreateLayout returned Null.");

        // Switch to the new layout BEFORE configuring it (matches LISP: _.ctab layoutNumber)
        ed.WriteMessage($"\n  [3/5] Switching to layout...");
        lm.SetCurrentLayoutId(lid);

        ed.WriteMessage($"\n  [4/5] Configuring (page setup, viewport, title block)...");
        using var tr = db.TransactionManager.StartTransaction();

        var layout = (Layout)tr.GetObject(lid, OpenMode.ForWrite);

        // 4a: page setup
        try
        {
            string plotDesc = PageSetupConfigurator.Configure(layout, info, ed);
            ed.WriteMessage($"\n    → {plotDesc}");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    !!! PAGE SETUP FAILED: {ex.GetType().Name}: {ex.Message}");
            throw;
        }

        // 4b: erase AutoCAD's default entities
        try
        {
            int erased = EraseAllEntitiesInLayout(layout, tr);
            ed.WriteMessage($"\n    Erased {erased} entities.");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    !!! ERASE FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            throw;
        }

        // 4c: viewport
        try
        {
            CreateViewport(layout, tr, result.BoundingBox, info);
            ed.WriteMessage($"\n    Viewport created.");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    !!! VIEWPORT FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            throw;
        }

        // 4d: insert title block — insertX = right edge of PRINTABLE area minus RightOffset.
        // This auto-adapts to actual margins (e.g. 25mm binding) instead of using
        // the legacy LISP printPoint values (which were hardcoded for old margin scheme).
        try
        {
            var (printableW, _) = GetPrintableArea(layout, info);
            double insertX = printableW - info.RightOffset;
            ed.WriteMessage($"\n    Printable area: {printableW:F1}mm wide, insertX={insertX:F1}");

            var (brN, brL, brLi) = ModelSpaceScanner.ParseLayoutNumber(result.LayoutName);
            TitleBlockInserter.Insert(db, tr, layout, insertX, 1.0, brN, brL, brLi, sastPath);
            ed.WriteMessage($"\n    Title block inserted (BR_N='{brN}', BR_L='{brL}', BR_LI='{brLi}').");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    !!! TITLEBLOCK FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            throw;
        }

        Application.SetSystemVariable("PSLTSCALE", 0);

        try
        {
            tr.Commit();
            ed.WriteMessage($"\n    Transaction committed.");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\n    !!! COMMIT FAILED: {ex.GetType().Name}: {ex.Message}");
            throw;
        }

        ed.WriteMessage($"\n  [5/5] Done — {result.BlockName} '{result.LayoutName}' ({info.Orientation})");
    }

    public static void CreateCustomLayout(
        Database db,
        BlockScanResult result,
        string sastPath,
        Editor ed)
    {
        var lm = LayoutManager.Current;

        ObjectId templateId = lm.GetLayoutId(result.BlockName);
        if (templateId.IsNull)
        {
            ed.WriteMessage($"\n  Template '{result.BlockName}' not found — skipping.");
            return;
        }

        ed.WriteMessage($"\n  [1/5] Importing title block...");
        TitleBlockInserter.EnsureImported(db, sastPath);

        ed.WriteMessage($"\n  [2/5] Copying '{result.BlockName}' → '{result.LayoutName}'...");
        lm.CopyLayout(result.BlockName, result.LayoutName);
        ObjectId lid = lm.GetLayoutId(result.LayoutName);
        if (lid.IsNull) throw new InvalidOperationException($"CopyLayout: '{result.LayoutName}' not found.");

        ed.WriteMessage($"\n  [3/5] Switching to layout...");
        lm.SetCurrentLayoutId(lid);

        ed.WriteMessage($"\n  [4/5] Configuring...");
        using var tr = db.TransactionManager.StartTransaction();
        var layout = (Layout)tr.GetObject(lid, OpenMode.ForWrite);

        EraseAllEntitiesInLayout(layout, tr);

        // Custom layout: viewport sized from model bounding box (no ISO paper defined)
        CreateCustomViewport(layout, tr, result.BoundingBox);

        double modelWidth  = Math.Abs(result.BoundingBox.MaxPoint.X - result.BoundingBox.MinPoint.X);
        double scaledWidth = modelWidth / 10.0 / result.BlockScale;

        var (brN, brL, brLi) = ModelSpaceScanner.ParseLayoutNumber(result.LayoutName);
        TitleBlockInserter.Insert(db, tr, layout, scaledWidth, 1.0, brN, brL, brLi, sastPath);

        Application.SetSystemVariable("PSLTSCALE", 0);
        tr.Commit();

        ed.WriteMessage($"\n  [5/5] Done — custom '{result.LayoutName}'");
        lm.SetCurrentLayoutId(lid);
    }

    // Erase all paper-space entities EXCEPT the paper-boundary viewport (Number == 1).
    // The paper-boundary VP is always present and must not be erased.
    // Returns the number of entities actually erased.
    private static int EraseAllEntitiesInLayout(Layout layout, Transaction tr)
    {
        var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
        int erased = 0;

        foreach (ObjectId id in psBtr.Cast<ObjectId>().ToList())
        {
            DBObject obj;
            try { obj = tr.GetObject(id, OpenMode.ForWrite); }
            catch { continue; }

            // Skip the paper-boundary viewport (Number == 1 in paper space)
            if (obj is Viewport vp)
            {
                int num;
                try { num = vp.Number; } catch { num = -1; }
                if (num == 1) continue;
            }

            try
            {
                ((Entity)obj).Erase();
                erased++;
            }
            catch { /* some entities can't be erased — skip silently */ }
        }

        return erased;
    }

    // Creates a viewport sized to the printable area, computed from the LAYOUT's
    // ACTUAL margins (set by PSV in PageSetupConfigurator). This auto-adapts to
    // any custom plotter configuration (e.g. user's DWG To PDF MAREN.pmp with
    // 25mm binding margin on the left).
    //
    // In AutoCAD paper-space with PlotOrigin (0,0), the printable area spans
    // (0, 0) → (printableW, printableH) where printableW = paperW - leftMargin - rightMargin.
    // The viewport center is at (printableW/2, printableH/2) — center of printable area.
    private static void CreateViewport(
        Layout layout,
        Transaction tr,
        Extents3d modelExtents,
        PaperSizeInfo info)
    {
        var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

        var (vpW, vpH) = GetPrintableArea(layout, info);
        double vpCx = vpW / 2.0;
        double vpCy = vpH / 2.0;

        AppendViewport(psBtr, tr, vpW, vpH, vpCx, vpCy, modelExtents);
    }

    // Returns (printableWidth, printableHeight) in mm.
    // Reads from layout.PlotPaperSize / PlotPaperMargins (set by PSV).
    // Falls back to PaperSizeInfo hardcoded dimensions with 5mm margins if PSV values are 0.
    public static (double w, double h) GetPrintableArea(Layout layout, PaperSizeInfo info)
    {
        Point2d   paper   = layout.PlotPaperSize;
        Extents2d margins = layout.PlotPaperMargins;

        double paperW = paper.X > 0 ? paper.X : info.PhysW;
        double paperH = paper.Y > 0 ? paper.Y : info.PhysH;
        double mLeft   = margins.MinPoint.X > 0 ? margins.MinPoint.X : 5.0;
        double mBottom = margins.MinPoint.Y > 0 ? margins.MinPoint.Y : 5.0;
        double mRight  = margins.MaxPoint.X > 0 ? margins.MaxPoint.X : 5.0;
        double mTop    = margins.MaxPoint.Y > 0 ? margins.MaxPoint.Y : 5.0;

        return (paperW - mLeft - mRight, paperH - mBottom - mTop);
    }

    // For custom layouts: viewport sized from model bounding box.
    private static void CreateCustomViewport(
        Layout layout,
        Transaction tr,
        Extents3d modelExtents)
    {
        var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

        Point2d  paperSize = layout.PlotPaperSize;
        Extents2d margins  = layout.PlotPaperMargins;
        double vpW  = Math.Max(paperSize.X - margins.MinPoint.X - margins.MaxPoint.X, 10.0);
        double vpH  = Math.Max(paperSize.Y - margins.MinPoint.Y - margins.MaxPoint.Y, 10.0);
        double vpCx = margins.MinPoint.X + vpW / 2.0;
        double vpCy = margins.MinPoint.Y + vpH / 2.0;

        AppendViewport(psBtr, tr, vpW, vpH, vpCx, vpCy, modelExtents);
    }

    // SAFE pattern: create empty Viewport, append to DB, AddNewlyCreatedDBObject,
    // THEN set properties (some like .On require the entity to be in the database).
    private static void AppendViewport(
        BlockTableRecord psBtr,
        Transaction tr,
        double vpW, double vpH,
        double vpCx, double vpCy,
        Extents3d modelExtents)
    {
        double mW = Math.Max(modelExtents.MaxPoint.X - modelExtents.MinPoint.X, 1.0);
        double mH = Math.Max(modelExtents.MaxPoint.Y - modelExtents.MinPoint.Y, 1.0);
        double vpAspect = vpW / vpH;
        double mAspect  = mW  / mH;
        double viewH    = mAspect > vpAspect ? mW / vpAspect : mH;

        // 1. Create empty viewport
        var vp = new Viewport();

        // 2. Append to BTR + register with transaction FIRST
        psBtr.AppendEntity(vp);
        tr.AddNewlyCreatedDBObject(vp, true);

        // 3. Now safe to set properties
        vp.CenterPoint = new Point3d(vpCx, vpCy, 0.0);
        vp.Width       = vpW;
        vp.Height      = vpH;
        vp.ViewCenter  = new Point2d(
            (modelExtents.MinPoint.X + modelExtents.MaxPoint.X) / 2.0,
            (modelExtents.MinPoint.Y + modelExtents.MaxPoint.Y) / 2.0);
        vp.ViewHeight  = viewH;
        vp.On          = true;
    }
}
