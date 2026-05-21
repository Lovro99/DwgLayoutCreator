using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace LayoutCreator;

// PaperName   = canonical display name passed to PSV (matched via GetLocaleMediaName)
// PrintPointX = LEGACY LISP value (kept for reference; v12+ computes insertX dynamically)
// Orientation = "Landscape" | "Portrait"
// PhysW/PhysH = physical ISO paper width × height in mm (fallback if PlotPaperSize unavailable)
// RightOffset = sastAu block origin distance from right edge of printable area (mm).
//               Defaults to 0 (block origin AT right edge of printable area).
//               Can be tuned per paper if needed (e.g. push block 5mm to the left).
public record PaperSizeInfo(
    string PaperName,
    double PrintPointX,
    string Orientation,
    double PhysW,
    double PhysH,
    double RightOffset);

public static class LayoutDefinitions
{
    public static readonly IReadOnlyDictionary<string, PaperSizeInfo> StandardSizes =
        new ReadOnlyDictionary<string, PaperSizeInfo>(new Dictionary<string, PaperSizeInfo>
        {
            ["A4L"] = new("ISO A4 (297.00 x 210.00 MM)", 287.0, "Landscape", 297.0, 210.0, 0.0),
            ["A4P"] = new("ISO A4 (210.00 x 297.00 MM)", 180.0, "Portrait",  210.0, 297.0, 0.0),
            ["A3L"] = new("ISO A3 (420.00 x 297.00 MM)", 390.0, "Landscape", 420.0, 297.0, 0.0),
            ["A3P"] = new("ISO A3 (297.00 x 420.00 MM)", 287.0, "Portrait",  297.0, 420.0, 0.0),
            ["A2L"] = new("ISO A2 (594.00 x 420.00 MM)", 564.0, "Landscape", 594.0, 420.0, 0.0),
            ["A2P"] = new("ISO A2 (420.00 x 594.00 MM)", 410.0, "Portrait",  420.0, 594.0, 0.0),
            ["A1L"] = new("ISO A1 (841.00 x 594.00 MM)", 811.0, "Landscape", 841.0, 594.0, 0.0),
            ["A1P"] = new("ISO A1 (594.00 x 841.00 MM)", 584.0, "Portrait",  594.0, 841.0, 0.0),
            ["A0L"] = new("ISO A0 (1189.00 x 841.00 MM)", 1159.0,"Landscape",1189.0, 841.0, 0.0),
            ["A0P"] = new("ISO A0 (841.00 x 1189.00 MM)", 831.0, "Portrait",  841.0,1189.0, 0.0),
        });

    public static readonly IReadOnlyList<string> CustomBlockNames =
        ["CUSTOM1","CUSTOM2","CUSTOM3","CUSTOM4","CUSTOM5","CUSTOM6","CUSTOM7","CUSTOM8"];

    public static readonly IReadOnlyList<string> AllBlockNames =
        ["A4L","A4P","A3L","A3P","A2L","A2P","A1L","A1P","A0L","A0P",
         "CUSTOM1","CUSTOM2","CUSTOM3","CUSTOM4","CUSTOM5","CUSTOM6","CUSTOM7","CUSTOM8"];

    public static bool IsCustomBlock(string blockName) =>
        CustomBlockNames.Contains(blockName, StringComparer.OrdinalIgnoreCase);
}
