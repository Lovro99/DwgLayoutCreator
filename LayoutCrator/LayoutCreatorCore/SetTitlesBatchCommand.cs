using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(LayoutCreator.SetTitlesBatchCommand))]

namespace LayoutCreator;

// Headless port SetLayoutTitles.lsp: za svaki red iz LAYOUT_BATCH_JSON["naslovi"]
// nadi layout po imenu (case-insensitive), u njegovom paper space BTR-u prvi
// INSERT bloka "sastAu", i postavi atribute layout_title / layout_mjerilo
// (tag usporedba case-insensitive, kao vl-setattributevalue).
//
// Ishodi (vjerno LISP sazetku): ok / noblock (nema bloka ILI nema oba atributa)
// / nolayout. Kompajlira se SAMO u LayoutCreatorCore.
public class SetTitlesBatchCommand
{
    private const string BlockName = "sastAu";
    private const string TagTitle = "layout_title";
    private const string TagScale = "layout_mjerilo";

    [CommandMethod("SETTITLESBATCH", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var ed = doc.Editor;

        BatchData? data = BatchDataFile.Load(ed);
        if (data is null) return;
        if (data.Naslovi is null)
        {
            ed.WriteMessage("\nERR| titles: nema kljuca 'naslovi' u LAYOUT_BATCH_JSON.");
            return;
        }

        // AttributeReference.TextString postavljanje moze kroz AdjustAlignment
        // koristiti krivu bazu ako WorkingDatabase nije nasa (plan §4.2 gotcha).
        // Postavi ju za trajanje komande i vrati na kraju.
        Database? prevWdb = HostApplicationServices.WorkingDatabase;
        HostApplicationServices.WorkingDatabase = db;

        int ok = 0, noblock = 0, nolayout = 0;
        try
        {
            using var tr = db.TransactionManager.StartTransaction();
            var layoutDict = (DBDictionary)tr.GetObject(
                db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (NaslovRow row in data.Naslovi)
            {
                ObjectId layoutId = FindLayout(layoutDict, row.Layout);
                if (layoutId.IsNull)
                {
                    ed.WriteMessage($"\nTITLES|nolayout|{row.Layout}");
                    nolayout++;
                    continue;
                }

                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(
                    layout.BlockTableRecordId, OpenMode.ForRead);
                BlockReference? bref = FirstSastBlock(tr, btr);
                if (bref is null)
                {
                    ed.WriteMessage($"\nTITLES|noblock|{row.Layout}");
                    noblock++;
                    continue;
                }

                // LISP: (and (setattr title) (setattr mjerilo)) — kratki spoj:
                // ako naslovni tag ne postoji, mjerilo se ne pokusava. Ishod je
                // "ok" samo ako su OBA taga nadjena.
                bool bothSet = SetAttribute(tr, bref, TagTitle, row.Naslov)
                             && SetAttribute(tr, bref, TagScale, row.Mjerilo);
                if (bothSet)
                {
                    ed.WriteMessage($"\nTITLES|ok|{row.Layout}");
                    ok++;
                }
                else
                {
                    ed.WriteMessage($"\nTITLES|noblock|{row.Layout}");
                    noblock++;
                }
            }

            tr.Commit();
        }
        finally
        {
            HostApplicationServices.WorkingDatabase = prevWdb;
        }

        ed.WriteMessage($"\nRESULT|titles|ok={ok} noblock={noblock} nolayout={nolayout}");
    }

    // Nadi layout po imenu (case-insensitive — AutoCAD tako tretira imena layouta).
    private static ObjectId FindLayout(DBDictionary dict, string name)
    {
        foreach (DBDictionaryEntry entry in dict)
        {
            if (entry.Key.Equals(name, System.StringComparison.OrdinalIgnoreCase))
                return entry.Value;
        }
        return ObjectId.Null;
    }

    // Prvi INSERT bloka "sastAu" u paper space BTR-u (ime preko
    // DynamicBlockTableRecord-a radi robusnosti — plan §4.2).
    private static BlockReference? FirstSastBlock(Transaction tr, BlockTableRecord btr)
    {
        foreach (ObjectId id in btr)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not BlockReference bref) continue;
            var dynBtr = (BlockTableRecord)tr.GetObject(
                bref.DynamicBlockTableRecord, OpenMode.ForRead);
            if (dynBtr.Name.Equals(BlockName, System.StringComparison.OrdinalIgnoreCase))
                return bref;
        }
        return null;
    }

    // Postavi atribut po tagu (case-insensitive). Vrati true ako je tag nadjen.
    private static bool SetAttribute(Transaction tr, BlockReference bref,
                                     string tag, string value)
    {
        foreach (ObjectId attId in bref.AttributeCollection)
        {
            var att = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
            if (att.Tag.Equals(tag, System.StringComparison.OrdinalIgnoreCase))
            {
                att.TextString = value;
                return true;
            }
        }
        return false;
    }
}
