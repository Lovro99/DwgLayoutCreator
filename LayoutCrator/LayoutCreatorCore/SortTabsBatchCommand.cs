using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: CommandClass(typeof(LayoutCreator.SortTabsBatchCommand))]

namespace LayoutCreator;

// Headless numericko sortiranje layout tabova preko TabOrder (jezgra
// TabSortV2-2.lsp: vla-put-TabOrder) + komparator iz ExportLayoutsToExcel.lsp
// (LayoutNameOrder). Model ostaje TabOrder 0; valjani layouti sortirani dobivaju
// 1..N; nevaljani (npr. CUSTOM template layouti) idu IZA njih (N+1..) u zatecenom
// redoslijedu (stabilno). Kompajlira se SAMO u LayoutCreatorCore.
public class SortTabsBatchCommand
{
    [CommandMethod("SORTTABSBATCH", CommandFlags.Modal)]
    public void Execute()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var ed = doc.Editor;

        using var tr = db.TransactionManager.StartTransaction();
        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

        // Skupi paper layoute (bez Modela) + trenutni TabOrder.
        var layouts = new List<LayoutInfo>();
        foreach (DBDictionaryEntry entry in layoutDict)
        {
            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
            if (layout.LayoutName.Equals("Model", System.StringComparison.OrdinalIgnoreCase))
                continue;
            layouts.Add(new LayoutInfo(layout.LayoutName, entry.Value, layout.TabOrder));
        }

        // Valjani: stabilno po komparatoru (prvi pa drugi broj).
        var valid = layouts.Where(l => LayoutNameOrder.IsValidLayout(l.Name))
                           .OrderBy(l => LayoutNameOrder.SortKey(l.Name).first)
                           .ThenBy(l => LayoutNameOrder.SortKey(l.Name).second)
                           .ToList();
        // Nevaljani: u zatecenom medusobnom redoslijedu (po trenutnom TabOrderu).
        var other = layouts.Where(l => !LayoutNameOrder.IsValidLayout(l.Name))
                          .OrderBy(l => l.Tab)
                          .ToList();

        var ordered = new List<LayoutInfo>(valid);
        ordered.AddRange(other);

        int pos = 1;
        foreach (LayoutInfo l in ordered)
        {
            var layout = (Layout)tr.GetObject(l.Id, OpenMode.ForWrite);
            layout.TabOrder = pos;
            ed.WriteMessage($"\nTABORDER|{pos}|{l.Name}");
            pos++;
        }
        tr.Commit();

        ed.WriteMessage($"\nRESULT|sorttabs|sorted={valid.Count} other={other.Count}");
    }

    private readonly struct LayoutInfo
    {
        public readonly string Name;
        public readonly ObjectId Id;
        public readonly int Tab;

        public LayoutInfo(string name, ObjectId id, int tab)
        {
            Name = name;
            Id = id;
            Tab = tab;
        }
    }
}
