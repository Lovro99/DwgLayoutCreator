using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DwgLayoutCreator;

// Standalone-pandam AutoCAD plugin LayoutDefinitions.
// PaperSize    = kanonsko ime (ACadSharp Layout.PaperSize) — identično AutoCAD .NET formatu
// PhysW/PhysH  = fizičke dimenzije u mm
// Orientation  = "Landscape" | "Portrait"
// RightOffset  = pomak sastavnice od desnog ruba printable area (0 = origin AT desni rub)
public record PaperSizeInfo(
    string PaperSize,
    string Orientation,
    double PhysW,
    double PhysH,
    double RightOffset);

public static class LayoutDefinitions
{
    // Kanonska imena (s donjim crtama) — POTVRĐENO identično AutoCAD .NET ISO listi
    public static readonly IReadOnlyDictionary<string, PaperSizeInfo> StandardSizes =
        new ReadOnlyDictionary<string, PaperSizeInfo>(new Dictionary<string, PaperSizeInfo>
        {
            ["A4L"] = new("ISO_A4_(297.00_x_210.00_MM)",  "Landscape", 297.0, 210.0, 0.0),
            ["A4P"] = new("ISO_A4_(210.00_x_297.00_MM)",  "Portrait",  210.0, 297.0, 0.0),
            ["A3L"] = new("ISO_A3_(420.00_x_297.00_MM)",  "Landscape", 420.0, 297.0, 0.0),
            ["A3P"] = new("ISO_A3_(297.00_x_420.00_MM)",  "Portrait",  297.0, 420.0, 0.0),
            ["A2L"] = new("ISO_A2_(594.00_x_420.00_MM)",  "Landscape", 594.0, 420.0, 0.0),
            ["A2P"] = new("ISO_A2_(420.00_x_594.00_MM)",  "Portrait",  420.0, 594.0, 0.0),
            ["A1L"] = new("ISO_A1_(841.00_x_594.00_MM)",  "Landscape", 841.0, 594.0, 0.0),
            ["A1P"] = new("ISO_A1_(594.00_x_841.00_MM)",  "Portrait",  594.0, 841.0, 0.0),
            ["A0L"] = new("ISO_A0_(1189.00_x_841.00_MM)", "Landscape",1189.0, 841.0, 0.0),
            ["A0P"] = new("ISO_A0_(841.00_x_1189.00_MM)", "Portrait",  841.0,1189.0, 0.0),
        });

    public static readonly IReadOnlyList<string> CustomBlockNames =
        ["CUSTOM1","CUSTOM2","CUSTOM3","CUSTOM4","CUSTOM5","CUSTOM6","CUSTOM7","CUSTOM8"];

    public static readonly IReadOnlyList<string> AllBlockNames =
        ["A4L","A4P","A3L","A3P","A2L","A2P","A1L","A1P","A0L","A0P",
         "CUSTOM1","CUSTOM2","CUSTOM3","CUSTOM4","CUSTOM5","CUSTOM6","CUSTOM7","CUSTOM8"];

    public static bool IsCustomBlock(string blockName) =>
        CustomBlockNames.Contains(blockName, StringComparer.OrdinalIgnoreCase);
}
