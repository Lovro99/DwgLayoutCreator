using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(LayoutCreator.VerifyBatchCommand))]

namespace LayoutCreator;

// Read-only dump stanja crteza za verifikaciju u orkestratoru (plan §4.6). NE
// mijenja nista i NE radi QSAVE. Zamjena za dosadasnji LISP verify pass (koji je
// citao ACAD_LAYOUT rjecnik u txt); sada sve ide kroz stdout markere:
//   VLAYOUT|<taborder>|<ime>            za svaki paper layout (sortirano po TabOrder)
//   VPROP|<kljuc>=<vrijednost>          za svaki custom drawing property
//   VTITLE|<layout>|<title>|<mjerilo>   za svaki layout koji ima sastAu blok
// Vrijednosti mogu sadrzavati '|' -> orkestrator parsira s ogranicenim brojem
// splitova. Kompajlira se SAMO u LayoutCreatorCore.
public class VerifyBatchCommand
{
    private const string BlockName = "sastAu";
    private const string TagTitle = "layout_title";
    private const string TagScale = "layout_mjerilo";

    [CommandMethod("VERIFYBATCH", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var ed = doc.Editor;

        using var tr = db.TransactionManager.StartTransaction();

        // Paper layouti (bez Modela) + TabOrder + paper space BTR.
        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
        var layouts = new List<LayoutInfo>();
        foreach (DBDictionaryEntry entry in layoutDict)
        {
            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
            if (layout.LayoutName.Equals("Model", System.StringComparison.OrdinalIgnoreCase))
                continue;
            layouts.Add(new LayoutInfo(layout.LayoutName, layout.TabOrder, layout.BlockTableRecordId));
        }
        layouts.Sort((a, b) => a.Tab.CompareTo(b.Tab));

        // VLAYOUT
        foreach (LayoutInfo l in layouts)
            ed.WriteMessage($"\nVLAYOUT|{l.Tab}|{l.Name}");

        // VPROP
        var builder = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
        foreach (System.Collections.DictionaryEntry de in builder.CustomPropertyTable)
            ed.WriteMessage($"\nVPROP|{de.Key}={de.Value}");

        // VTITLE (samo layouti koji imaju sastAu blok)
        foreach (LayoutInfo l in layouts)
        {
            var btr = (BlockTableRecord)tr.GetObject(l.BtrId, OpenMode.ForRead);
            if (TryReadTitle(tr, btr, out string title, out string scale))
                ed.WriteMessage($"\nVTITLE|{l.Name}|{title}|{scale}");
        }

        tr.Commit();
        ed.WriteMessage("\nRESULT|verify|done");
    }

    // Procitaj layout_title / layout_mjerilo iz prvog sastAu bloka. Vrati true
    // ako je sastAu blok pronaden (i tada su title/scale popunjeni; nedostajuci
    // tag ostaje prazan string).
    private static bool TryReadTitle(Transaction tr, BlockTableRecord btr,
                                     out string title, out string scale)
    {
        title = "";
        scale = "";
        foreach (ObjectId id in btr)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not BlockReference bref) continue;
            var dynBtr = (BlockTableRecord)tr.GetObject(
                bref.DynamicBlockTableRecord, OpenMode.ForRead);
            if (!dynBtr.Name.Equals(BlockName, System.StringComparison.OrdinalIgnoreCase))
                continue;
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                var att = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                if (att.Tag.Equals(TagTitle, System.StringComparison.OrdinalIgnoreCase))
                    title = att.TextString;
                else if (att.Tag.Equals(TagScale, System.StringComparison.OrdinalIgnoreCase))
                    scale = att.TextString;
            }
            return true;   // prvi sastAu blok
        }
        return false;
    }

    private readonly struct LayoutInfo
    {
        public readonly string Name;
        public readonly int Tab;
        public readonly ObjectId BtrId;

        public LayoutInfo(string name, int tab, ObjectId btrId)
        {
            Name = name;
            Tab = tab;
            BtrId = btrId;
        }
    }
}
