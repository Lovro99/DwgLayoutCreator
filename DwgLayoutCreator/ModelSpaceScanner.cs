using System.Collections.Generic;
using System.Text.RegularExpressions;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.Objects;
using CSMath;

namespace DwgLayoutCreator;

// Bounding box izveden iz Insert entiteta u model spaceu.
public record BlockScanResult(
    string BlockName,
    string LayoutName,
    XY MinPt,
    XY MaxPt,
    double BlockScale);

public static partial class ModelSpaceScanner
{
    [GeneratedRegex(@"\d+(\.\d+)?")]
    private static partial Regex NumberPattern();

    public static List<BlockScanResult> Scan(CadDocument doc)
    {
        var results   = new List<BlockScanResult>();
        var knownSet  = new HashSet<string>(LayoutDefinitions.AllBlockNames, StringComparer.OrdinalIgnoreCase);

        foreach (var ent in doc.ModelSpace.Entities)
        {
            if (ent is not Insert ins) continue;

            string blockName = ins.Block?.Name ?? "";
            if (!knownSet.Contains(blockName)) continue;

            string layoutName = ReadLayoutName(ins);
            if (string.IsNullOrEmpty(layoutName)) continue;

            // Block bounding box u model spaceu: insert point + skalirane dimenzije bloka.
            // ACadSharp ne nudi GeometricExtents direktno; aproksimiramo iz Block.Extents.
            var (minPt, maxPt) = ComputeBoundingBox(ins);

            results.Add(new BlockScanResult(
                BlockName:   blockName,
                LayoutName:  layoutName,
                MinPt:       minPt,
                MaxPt:       maxPt,
                BlockScale:  ins.XScale
            ));
        }

        return results;
    }

    // Čita "layout broj" iz atributa Insert-a.
    // Prioritet: BR_N tag → posljednji atribut (matchira LISP Group:tags:values).
    // Ovo radi neovisno o tagu (ORDER_NUMBER, BR_N, itd.) — uvijek vrati VRIJEDNOST.
    private static string ReadLayoutName(Insert ins)
    {
        string lastValue = string.Empty;
        foreach (var attr in ins.Attributes)
        {
            if (attr.Tag.Equals("BR_N", StringComparison.OrdinalIgnoreCase))
                return (attr.Value ?? "").Trim();
            lastValue = (attr.Value ?? "").Trim();
        }
        return lastValue;
    }

    // Aproksimacija bounding boxa: uzima Block.BlockEntity (insert point) + scale * block extents.
    // Block.Extents nije pouzdano u ACadSharpu — koristimo InsertPoint kao centar i pretpostavku
    // (model je u istim jedinicama, blokovi su skalirani 1:1 najčešće).
    private static (XY min, XY max) ComputeBoundingBox(Insert ins)
    {
        // Fallback: ako ne možemo izračunati, koristimo insert point ± 100mm.
        var origin = new XY(ins.InsertPoint.X, ins.InsertPoint.Y);

        // Pokušaj izračunati iz block-records-a (sume Entities ekstrema).
        if (ins.Block is { } block)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            bool any = false;

            foreach (var e in block.Entities)
            {
                // ACadSharp Entity ima različite property-je po tipu — koristimo refleksiju kroz uobičajene
                // (Line.StartPoint/EndPoint, Insert.InsertPoint, itd.). Za jednostavnost,
                // uzmemo samo InsertPoint djece (radi za sad — bounding box je samo za viewport zoom).
                if (e is Insert sub)
                {
                    var p = sub.InsertPoint;
                    minX = Math.Min(minX, p.X); maxX = Math.Max(maxX, p.X);
                    minY = Math.Min(minY, p.Y); maxY = Math.Max(maxY, p.Y);
                    any = true;
                }
            }

            if (any)
            {
                // Transformiraj u world coordinates: dodaj insert point + scale
                double sx = ins.XScale, sy = ins.YScale;
                return (
                    new XY(origin.X + minX * sx, origin.Y + minY * sy),
                    new XY(origin.X + maxX * sx, origin.Y + maxY * sy)
                );
            }
        }

        // Fallback: ±50mm oko insert pointa (dovoljno za zoom)
        return (new XY(origin.X - 50, origin.Y - 50), new XY(origin.X + 50, origin.Y + 50));
    }

    // Replicira LISP numbersFromString: izvlači sve numeričke tokene iz layout imena.
    // "1-2-3" → ("1", "2", "3"), "1-1x1" → ("1", "1", "1")
    public static (string brN, string brL, string brLi) ParseLayoutNumber(string layoutName)
    {
        var matches = NumberPattern().Matches(layoutName);
        string brN  = matches.Count > 0 ? matches[0].Value : layoutName;
        string brL  = matches.Count > 1 ? matches[1].Value : string.Empty;
        string brLi = matches.Count > 2 ? matches[2].Value : string.Empty;
        return (brN, brL, brLi);
    }

    public static List<string> GetExistingLayoutNames(CadDocument doc)
    {
        var names = new List<string>();
        foreach (var layout in doc.Layouts)
            names.Add(layout.Name);
        return names;
    }
}
