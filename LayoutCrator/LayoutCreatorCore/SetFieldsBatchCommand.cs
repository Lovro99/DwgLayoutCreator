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
        // -> db.SummaryInfo. VAZNO: cijeli upis mora doseci db.SummaryIntfo cak i
        // ako pojedini kljuc baci — inace jedna losa iznimka tiho ponisti SVE
        // (uzrok bug-a: SETFIELDSBATCH bi pukao prije db.SummaryInfo pa se nista
        // ne spremi, a accoreconsole samo predje na sljedecu komandu). Zato:
        //   - po-kljuc try/catch (jedan los kljuc ne rusi ostatak),
        //   - vanjski try/catch (setter/getter iznimke se vide kao ERR|, ne tiho),
        //   - self-verify u ISTOJ sesiji (present=P/Q) da odmah znamo je li upis
        //     uopce "sjeo" prije spremanja — ako je P=Q ovdje a VPROP prazan u
        //     verify passu, onda je rijec o problemu SPREMANJA, ne upisa.
        // Postavi radnu bazu na nasu za trajanje upisa — SETTITLESBATCH (cije se
        // izmjene POUZDANO spremaju) isto to radi; SummaryInfo upis se u nekim
        // core-console kontekstima ne serijalizira ako radna baza nije nasa.
        Database? prevWdb = HostApplicationServices.WorkingDatabase;
        HostApplicationServices.WorkingDatabase = db;

        int added = 0, overwritten = 0, failed = 0, present = 0;
        try
        {
            var builder = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
            System.Collections.IDictionary table = builder.CustomPropertyTable;

            foreach (var kv in data.Polja)
            {
                try
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
                catch (System.Exception exKey)
                {
                    failed++;
                    ed.WriteMessage($"\nERR| fields: kljuc '{kv.Key}': {exKey.Message}");
                }
            }

            db.SummaryInfo = builder.ToDatabaseSummaryInfo();

            // Self-verify: ponovno procitaj SummaryInfo i prebroji koliko je kljuceva
            // stvarno prisutno (u istoj sesiji, prije spremanja).
            var check = new DatabaseSummaryInfoBuilder(db.SummaryInfo).CustomPropertyTable;
            foreach (var kv in data.Polja)
                if (check.Contains(kv.Key)) present++;
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERR| fields: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            HostApplicationServices.WorkingDatabase = prevWdb;
        }

        // Prosireni RESULT (orkestrator i dalje parsira added/overwritten; failed/
        // present su dijagnostika). Ako je present < ukupno, upis nije "sjeo".
        ed.WriteMessage(
            $"\nRESULT|fields|added={added} overwritten={overwritten} " +
            $"failed={failed} present={present}/{data.Polja.Count}");
    }
}
