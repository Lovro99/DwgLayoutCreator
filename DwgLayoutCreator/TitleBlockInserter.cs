using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Objects;
using ACadSharp.Tables;
using CSMath;

namespace DwgLayoutCreator;

public static class TitleBlockInserter
{
    private const string TitleBlockLayer = "MAREN SASTAVNICA";

    // Importira sastAu.dwg block u target dokument ako ne postoji.
    // ACadSharp ne nudi direktni "Database.Insert(extDb)" kao AutoCAD .NET API —
    // ručno kopiramo BlockRecord iz extDb u doc.
    public static void EnsureImported(CadDocument doc, string sastPath)
    {
        string blockName = Path.GetFileNameWithoutExtension(sastPath); // "sastAu"

        // Već postoji u doc-u?
        if (doc.BlockRecords.Any(b => b.Name.Equals(blockName, StringComparison.OrdinalIgnoreCase)))
            return;

        // Čitaj external DWG
        CadDocument extDoc;
        using (var reader = new DwgReader(sastPath))
            extDoc = reader.Read();

        // Pristup: koristi extDoc.ModelSpace ENTITIES kao sadržaj novog BlockRecord-a "sastAu".
        // Razlog: u sastAu.dwg, sve geometrije sastavnice (linije, hatch, ATTDEFs) su u
        // model spaceu — to je standardni format "DWG kao blok".
        var newBlock = new BlockRecord(blockName);

        // Kopiraj sve entitete iz ext model space u novi block.
        // Direktno doc.Add(entity) odjavljuje s extDoc i registrira u doc.
        foreach (var ent in extDoc.ModelSpace.Entities.ToList())
        {
            // ACadSharp entiteti se mogu klonirati pa kopirati u novi dokument.
            var cloned = (Entity)ent.Clone();
            newBlock.Entities.Add(cloned);
        }

        doc.BlockRecords.Add(newBlock);
    }

    // Kreira Insert za sastAu u paper space layoutu i postavlja BR_N/BR_L/BR_LI atribute.
    public static void Insert(
        CadDocument doc,
        Layout layout,
        double insertX,
        double scale,
        string brN, string brL, string brLi,
        string sastPath)
    {
        string blockName = Path.GetFileNameWithoutExtension(sastPath);

        var blockRec = doc.BlockRecords.FirstOrDefault(b =>
            b.Name.Equals(blockName, StringComparison.OrdinalIgnoreCase));

        if (blockRec is null)
            throw new InvalidOperationException(
                $"Block '{blockName}' not in document. Call EnsureImported first.");

        // Layer "MAREN SASTAVNICA" — kreiraj ako ne postoji
        var layer = EnsureLayerExists(doc, TitleBlockLayer);

        // KORAK 1: Kreiraj Insert (ctor auto-kreira AttributeEntity-je iz BlockRecord.AttributeDefinitions)
        var ins = new Insert(blockRec)
        {
            InsertPoint = new XYZ(insertX, 0.0, 0.0),
            XScale      = scale,
            YScale      = scale,
            ZScale      = scale,
            Rotation    = 0.0,
            Layer       = layer,
        };

        // KORAK 2: WORKAROUND za ACadSharp 3.5.7 bug u DwgObjectWriter.writeCommonAttData:
        //          za atribute tipa MultiLine / ConstantMultiLine writer čita att.MText
        //          i zove writeEntityMode(att.MText). Ako je MText null (a za auto-kreirane
        //          attribute iz Insert(BlockRecord) ctora UVIJEK je null), pada s NRE u
        //          getEntMode jer null.Owner ne može.
        //          Block "sastAu" ima LAYOUT_TITLE kao ConstantMultiLine attribute.
        //          Najlakša fix: prebaci sve multiline atribute u SingleLine — vrijednost
        //          se ionako prikazuje isto, a writer ne tipa MText path. Alternativno bi
        //          se mogao kreirati prazan MText, ali to dodaje nepotreban entitet u DWG.
        foreach (var attr in ins.Attributes)
        {
            if (attr.AttributeType == AttributeType.MultiLine ||
                attr.AttributeType == AttributeType.ConstantMultiLine)
            {
                attr.AttributeType = AttributeType.SingleLine;
            }
        }

        // KORAK 3: dodaj Insert u paper space block — ovo postavlja Owner chain
        //          za auto-kreirane atribute (oni su djeca Insert-a, pa nasljeđuju kontekst).
        layout.AssociatedBlock.Entities.Add(ins);

        // KORAK 4: SAD postavi Layer + Value na svaki atribut
        foreach (var attr in ins.Attributes)
        {
            attr.Layer = layer;
            if (attr.Tag.Equals("BR_N",  StringComparison.OrdinalIgnoreCase)) attr.Value = brN;
            if (attr.Tag.Equals("BR_L",  StringComparison.OrdinalIgnoreCase)) attr.Value = brL;
            if (attr.Tag.Equals("BR_LI", StringComparison.OrdinalIgnoreCase)) attr.Value = brLi;
        }

        // KORAK 5: Sinkronizacija — ACadSharp preporuka
        ins.UpdateAttributes();
    }

    // Vraća postojeći ili novi Layer (uvijek validan, nikad null).
    private static Layer EnsureLayerExists(CadDocument doc, string name)
    {
        var existing = doc.Layers.FirstOrDefault(l =>
            l.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) return existing;

        var created = new Layer(name);
        doc.Layers.Add(created);
        return created;
    }
}
