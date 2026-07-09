using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(LayoutCreator.SetFieldsBatchCommand))]

namespace LayoutCreator;

// Headless port SetFieldsValue.lsp: upisuje custom drawing properties
// (SummaryInfo) iz LAYOUT_BATCH_JSON["polja"]. Postojeci kljuc -> prepisi; novi
// -> dodaj. Kljucevi case-sensitivno kao u LISP-u (usporedba preko
// CustomPropertyTable.Contains). FIELD objekti koji referenciraju ta svojstva
// osvjeze se na UPDATEFIELD/REGEN (orkestrator zove _.UPDATEFIELD nakon svih
// izmjena — plan §3).
//
// Kompajlira se SAMO u LayoutCreatorCore (nikad u UI plugin).
public class SetFieldsBatchCommand
{
    [CommandMethod("SETFIELDSBATCH", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var ed = doc.Editor;

        BatchData? data = BatchDataFile.Load(ed);
        if (data is null) return;
        if (data.Polja is null)
        {
            ed.WriteMessage("\nERR| fields: nema kljuca 'polja' u LAYOUT_BATCH_JSON.");
            return;
        }

        // Mehanika iz plana §4.1: DatabaseSummaryInfoBuilder -> CustomPropertyTable
        // -> db.SummaryInfo.
        var builder = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
        System.Collections.IDictionary table = builder.CustomPropertyTable;

        int added = 0, overwritten = 0;
        foreach (var kv in data.Polja)
        {
            if (table.Contains(kv.Key))
            {
                table[kv.Key] = kv.Value;
                overwritten++;
            }
            else
            {
                table.Add(kv.Key, kv.Value);
                added++;
            }
        }
        db.SummaryInfo = builder.ToDatabaseSummaryInfo();

        // Sazetak (po retku nista — prebucno; VERIFYBATCH ionako dumpa VPROP|).
        ed.WriteMessage($"\nRESULT|fields|added={added} overwritten={overwritten}");
    }
}
