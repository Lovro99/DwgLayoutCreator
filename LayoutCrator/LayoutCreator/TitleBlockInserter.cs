using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LayoutCreator;

public static class TitleBlockInserter
{
    // Imports the block definition from the external DWG into the database,
    // committed in its own standalone transaction.
    // Call this BEFORE opening any layout-configuration transaction,
    // because db.Insert() manages its own internal transaction and will cause
    // eNotInDatabase if called inside another open transaction.
    public static void EnsureImported(Database db, string sastPath)
    {
        string blockName = BlockNameFrom(sastPath);

        // Check without touching an open transaction
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has(blockName)) { tr.Commit(); return; }
            tr.Commit();
        }

        // Import in isolation — db.Insert manages its own internal transaction
        using var extDb = new Database(false, true);
        extDb.ReadDwgFile(sastPath, FileOpenMode.OpenForReadAndAllShare, false, null);
        db.Insert(blockName, extDb, true);
    }

    // Inserts a block reference in the given layout's paper space at (insertX, 0).
    // Assumes EnsureImported was already called (block definition is in db).
    // Sets BR_N, BR_L, BR_LI attributes on the inserted block.
    public static void Insert(
        Database db,
        Transaction tr,
        Layout layout,
        double insertX,
        double scale,
        string brN, string brL, string brLi,
        string sastPath)
    {
        string blockName = BlockNameFrom(sastPath);

        // Block definition must already exist (imported by EnsureImported)
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (!bt.Has(blockName))
            throw new InvalidOperationException(
                $"Block '{blockName}' not in database. EnsureImported must be called first.");

        ObjectId blockDefId = bt[blockName];
        var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

        // Create with constructor (point + blockDefId required), append to DB,
        // THEN set additional properties. Setting Layer (string name lookup) before
        // appending can cause eNotInDatabase.
        var bref = new BlockReference(new Point3d(insertX, 0.0, 0.0), blockDefId);
        psBtr.AppendEntity(bref);
        tr.AddNewlyCreatedDBObject(bref, true);

        bref.ScaleFactors = new Scale3d(scale, scale, scale);
        bref.Rotation     = 0.0;

        // Set layer only if it exists (avoid eNotInDatabase from missing layer lookup)
        var lt = (LayerTable)tr.GetObject(bref.Database.LayerTableId, OpenMode.ForRead);
        if (lt.Has("MAREN SASTAVNICA"))
            bref.Layer = "MAREN SASTAVNICA";

        // Synchronize attribute references — bref is now safely in the database
        var blockDef = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
        foreach (ObjectId defId in blockDef)
        {
            if (tr.GetObject(defId, OpenMode.ForRead) is not AttributeDefinition attDef) continue;
            if (attDef.Constant) continue;

            var attRef = new AttributeReference();
            attRef.SetAttributeFromBlock(attDef, bref.BlockTransform);
            bref.AttributeCollection.AppendAttribute(attRef);
            tr.AddNewlyCreatedDBObject(attRef, true);
        }

        SetAttribute(bref, tr, "BR_N",  brN);
        SetAttribute(bref, tr, "BR_L",  brL);
        SetAttribute(bref, tr, "BR_LI", brLi);
    }

    private static void SetAttribute(BlockReference bref, Transaction tr, string tag, string value)
    {
        foreach (ObjectId attId in bref.AttributeCollection)
        {
            var att = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
            if (att.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
            {
                att.TextString = value;
                return;
            }
        }
    }

    public static string BlockNameFrom(string sastPath) =>
        Path.GetFileNameWithoutExtension(sastPath); // "sastAu"
}
