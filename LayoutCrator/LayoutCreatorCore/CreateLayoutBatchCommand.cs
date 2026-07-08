using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(LayoutCreator.CreateLayoutBatchCommand))]

namespace LayoutCreator;

// Headless batch entry point invoked from accoreconsole scripts by the Python
// orchestrator (batch/batch_layouts.py, Korak 2). Reuses the exact same layout
// logic as the interactive CREATELAYOUT (CreateLayoutCommand.RunCreateLayouts);
// the only difference is how sastAu.dwg is located: the support-file search path
// in the core console is NOT the same as in full AutoCAD, so the orchestrator
// passes the title-block path via the LAYOUT_SAST_PATH environment variable,
// with a fallback to the normal support path.
//
// This file is compiled ONLY into LayoutCreatorCore (core console), never into
// the interactive UI plugin.
public class CreateLayoutBatchCommand
{
    [CommandMethod("CREATELAYOUTBATCH", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db  = doc.Database;
        var ed  = doc.Editor;

        string? sastPath = ResolveSastPath(db, ed);
        if (sastPath is null)
        {
            // Fatal for THIS file, but before any modification — orchestrator
            // treats a missing title block as ERR and leaves the original intact.
            ed.WriteMessage("\nERR| sastAu.dwg not found (LAYOUT_SAST_PATH unset/invalid and not on support path).");
            return;
        }
        ed.WriteMessage($"\nINFO| sastAu.dwg: {sastPath}");

        var (created, candidates) = CreateLayoutCommand.RunCreateLayouts(db, ed, sastPath);

        // Machine-parseable lines the orchestrator scrapes from stdout: one
        // CREATED| per new layout (verify pass confirms each name persisted),
        // then a RESULT| summary.
        foreach (string name in created)
            ed.WriteMessage($"\nCREATED| {name}");
        ed.WriteMessage($"\nRESULT| created={created.Count} candidates={candidates}");
    }

    // Resolution order: LAYOUT_SAST_PATH env var (set by the orchestrator) →
    // AutoCAD support search path (HostApplicationServices.FindFile).
    private static string? ResolveSastPath(Database db, Editor ed)
    {
        string? env = Environment.GetEnvironmentVariable("LAYOUT_SAST_PATH");
        if (!string.IsNullOrWhiteSpace(env))
        {
            if (System.IO.File.Exists(env)) return env;
            ed.WriteMessage($"\nWARN| LAYOUT_SAST_PATH set but file not found: {env}");
        }

        try
        {
            string found = HostApplicationServices.Current.FindFile(
                "sastAu.dwg", db, FindFileHint.Default);
            if (!string.IsNullOrEmpty(found)) return found;
        }
        catch { /* FindFile throws when not found — fall through to null */ }

        return null;
    }
}
