#if CORECONSOLE
// Headless core console (accoreconsole) has no UI Application layer (acmgd).
// Bind the unqualified name `Application` to the core Application so the exact
// same source compiles for both the interactive plugin and the core build.
// See LayoutCrator/LayoutCreatorCore.
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(LayoutCreator.CreateLayoutCommand))]

namespace LayoutCreator;

public class CreateLayoutCommand
{
    // Equivalent to LISP c:createlayout
    // Scans model space for known paper-size blocks and creates layouts for each found.
    [CommandMethod("CREATELAYOUT", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db  = doc.Database;
        var ed  = doc.Editor;

        // Find sastAu.dwg via AutoCAD support paths (mirrors LISP findfile)
        string? sastPath = FindFile("sastAu.dwg");
        if (sastPath is null)
        {
            ed.WriteMessage("\nsastAu.dwg not found! Set up support path.");
            return;
        }

        var (created, _) = RunCreateLayouts(db, ed, sastPath);

#if !CORECONSOLE
        // Force regen (equivalent to LISP "regenall"). Interactive-only:
        // accoreconsole regenerates on open, and Document.SendStringToExecute
        // is part of the UI layer (acmgd) that the core build does not reference.
        doc.SendStringToExecute("regenall\n", false, false, false);
#endif

        ed.WriteMessage($"\nDone. {created} layout(s) created.");
    }

    // Shared layout-creation core, reused verbatim by both the interactive
    // CREATELAYOUT command and the headless CREATELAYOUTBATCH command
    // (LayoutCreatorCore). Scans model space and creates one layout per
    // numbered paper-size block; returns (created, candidates) so batch
    // callers can log and verify counts.
    //
    // `Application` resolves to the core Application in the core build (see the
    // CORECONSOLE alias at the top of this file), so SetSystemVariable — which
    // has the same signature on both the UI and core Application — runs headless.
    internal static (int created, int candidates) RunCreateLayouts(
        Database db, Editor ed, string sastPath)
    {
        // --- Phase 1: read model space (single read transaction) ---
        List<BlockScanResult> candidates;
        List<string> existingLayouts;

        using (var tr = db.TransactionManager.StartTransaction())
        {
            SetOrCreateLayer(db, tr, "MAREN SASTAVNICA");
            existingLayouts = ModelSpaceScanner.GetExistingLayoutNames(db, tr);
            candidates      = ModelSpaceScanner.Scan(db, tr);
            tr.Commit();
        }

        Application.SetSystemVariable("LTSCALE",   1.0);
        Application.SetSystemVariable("MSLTSCALE", 1);

        ed.WriteMessage($"\nFound {candidates.Count} block(s) matching layout names.");
        ed.WriteMessage($"\nExisting layouts: {string.Join(", ", existingLayouts)}");

        int created = 0;

        // --- Phase 2: create layouts (one transaction per layout) ---
        foreach (var result in candidates)
        {
            if (result.LayoutName.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage($"\n{result.BlockName}: layout not numbered (X), skipping.");
                continue;
            }

            if (existingLayouts.Contains(result.LayoutName, StringComparer.OrdinalIgnoreCase))
            {
                ed.WriteMessage($"\n{result.BlockName}: layout '{result.LayoutName}' already exists, skipping.");
                continue;
            }

            try
            {
                // LayoutBuilder manages its own transactions internally —
                // LayoutManager.CreateLayout uses an internal transaction and must
                // not be called inside an outer transaction (causes eNotInDatabase).
                if (LayoutDefinitions.IsCustomBlock(result.BlockName))
                    LayoutBuilder.CreateCustomLayout(db, result, sastPath, ed);
                else
                    LayoutBuilder.CreateStandardLayout(db, result, sastPath, ed);

                existingLayouts.Add(result.LayoutName);
                created++;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating layout '{result.LayoutName}': {ex.Message}");
            }
        }

        return (created, candidates.Count);
    }

    // Scans model space and prints every BlockReference with its effective name and BR_N attribute.
    // Run DIAGLAYOUTS in AutoCAD to verify the scanner sees the correct blocks.
    [CommandMethod("DIAGLAYOUTS", CommandFlags.Modal)]
    public void DiagLayouts()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db  = doc.Database;
        var ed  = doc.Editor;

        using var tr = db.TransactionManager.StartTransaction();
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        int total = 0, matched = 0;
        var knownNames = new HashSet<string>(LayoutDefinitions.AllBlockNames, StringComparer.OrdinalIgnoreCase);

        foreach (ObjectId id in ms)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not BlockReference bref) continue;
            total++;

            var dynBtr = (BlockTableRecord)tr.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForRead);
            string effName = dynBtr.Name;

            // Read ALL attributes so the diagnostic is not misleading
            var attrParts = new System.Collections.Generic.List<string>();
            string layoutValue = string.Empty;
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                var att = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                attrParts.Add($"{att.Tag}='{att.TextString}'");
                layoutValue = att.TextString.Trim(); // last wins (matches LISP behaviour)
                if (att.Tag.Equals("BR_N", StringComparison.OrdinalIgnoreCase))
                    layoutValue = att.TextString.Trim(); // prefer BR_N if found
            }
            string attrsDisplay = attrParts.Count > 0 ? string.Join(", ", attrParts) : "(no attributes)";
            string layoutDisplay = string.IsNullOrEmpty(layoutValue) ? "(empty)" : $"→ layoutName='{layoutValue}'";

            bool isKnown = knownNames.Contains(effName);
            if (isKnown) matched++;
            ed.WriteMessage($"\n  [{(isKnown ? "MATCH" : "skip ")}] '{effName}'  {attrsDisplay}  {layoutDisplay}  dynamic={bref.IsDynamicBlock}");
        }

        tr.Commit();
        ed.WriteMessage($"\nTotal BlockRefs: {total}, matched known names: {matched}");
    }

    // Adds a diagnostic command to list available PDF media names.
    // Run LISTMEDIA in AutoCAD to troubleshoot paper size mismatches.
    [CommandMethod("LISTMEDIA", CommandFlags.Modal)]
    public void ListMedia()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db  = doc.Database;
        var ed  = doc.Editor;

        PageSetupConfigurator.ListAvailableMedia(ed);
    }

    private static void SetOrCreateLayer(Database db, Transaction tr, string name)
    {
        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
        if (!lt.Has(name))
        {
            var layer = new LayerTableRecord { Name = name };
            lt.Add(layer);
            tr.AddNewlyCreatedDBObject(layer, true);
        }
        db.Clayer = lt[name];
    }

    // Searches AutoCAD support paths for the given filename.
    // Uses HostApplicationServices.FindFile — the direct .NET equivalent of LISP findfile.
    private static string? FindFile(string filename)
    {
        try
        {
            return HostApplicationServices.Current.FindFile(
                filename,
                Application.DocumentManager.MdiActiveDocument.Database,
                FindFileHint.Default);
        }
        catch
        {
            return null;
        }
    }
}
