using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LayoutCreator;

public record BlockScanResult(
    string BlockName,
    string LayoutName,
    Extents3d BoundingBox,
    double BlockScale
);

public static partial class ModelSpaceScanner
{
    [GeneratedRegex(@"\d+(\.\d+)?")]
    private static partial Regex NumberPattern();

    public static List<BlockScanResult> Scan(Database db, Transaction tr)
    {
        var results = new List<BlockScanResult>();
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        var allNames = new HashSet<string>(LayoutDefinitions.AllBlockNames, StringComparer.OrdinalIgnoreCase);

        foreach (ObjectId id in ms)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not BlockReference bref) continue;

            // For dynamic blocks, DynamicBlockTableRecord points to the source definition BTR
            // whose Name is the effective (visible) block name — replaces deprecated EffectiveName.
            var dynBtr = (BlockTableRecord)tr.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForRead);
            string effectiveName = dynBtr.Name;
            if (!allNames.Contains(effectiveName)) continue;

            string layoutName = ReadBrNAttribute(bref, tr);
            if (string.IsNullOrEmpty(layoutName)) continue;

            Extents3d ext;
            try { ext = bref.GeometricExtents; }
            catch { continue; }

            results.Add(new BlockScanResult(
                BlockName: effectiveName,
                LayoutName: layoutName,
                BoundingBox: ext,
                BlockScale: bref.ScaleFactors.X
            ));
        }

        return results;
    }

    private static string ReadBrNAttribute(BlockReference bref, Transaction tr)
    {
        // Matches LISP Group:tags:values behaviour:
        //   1. Return "BR_N" tag value if it exists (backward compat).
        //   2. Otherwise return the LAST attribute value found — LISP builds the list
        //      with (cons ...) so the last attribute ends up at (car lst), i.e. the
        //      first element returned, which is what (cadr (car ...)) reads.
        // This makes the scanner work regardless of the attribute tag name used
        // on the paper-size block (e.g. "ORDER_NUMBER", "BR_N", etc.).
        string lastValue = string.Empty;

        foreach (ObjectId attId in bref.AttributeCollection)
        {
            var att = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
            if (att.Tag.Equals("BR_N", StringComparison.OrdinalIgnoreCase))
                return att.TextString.Trim();
            lastValue = att.TextString.Trim(); // keep updating → ends up as the last one
        }

        return lastValue; // empty if no attributes at all
    }

    // Replicates LISP numbersFromString: extracts all numeric tokens from layout name.
    // "1-2-3" → ("1", "2", "3"), "1.5" → ("1.5"), "1" → ("1")
    public static (string brN, string brL, string brLi) ParseLayoutNumber(string layoutName)
    {
        var matches = NumberPattern().Matches(layoutName);
        string brN  = matches.Count > 0 ? matches[0].Value : layoutName;
        string brL  = matches.Count > 1 ? matches[1].Value : string.Empty;
        string brLi = matches.Count > 2 ? matches[2].Value : string.Empty;
        return (brN, brL, brLi);
    }

    public static List<string> GetExistingLayoutNames(Database db, Transaction tr)
    {
        var names = new List<string>();
        var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
        foreach (DBDictionaryEntry entry in dict)
            names.Add(entry.Key);
        return names;
    }
}
