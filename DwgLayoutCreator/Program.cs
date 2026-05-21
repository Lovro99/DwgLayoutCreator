/*
 * dwg_layout_creator — standalone DWG layout generator (ACadSharp, bez AutoCAD-a)
 *
 * Ekvivalent AutoCAD .NET plugina (LayoutCreator) — ali radi bez otvorenog AutoCAD-a.
 *
 * Poziv:
 *   dwg_layout_creator.exe <dwg_path> <sastau_dwg_path>
 *
 * Primjer:
 *   dwg_layout_creator.exe "C:\crtez.dwg" "\\server\sastavnice\sastAu.dwg"
 *
 * Stdout linije:
 *   INFO|<poruka>    — informacija (početak, korak)
 *   OK|<poruka>      — uspješno kreiran layout
 *   WARN|<poruka>    — preskočeno (već postoji, "X" oznaka, itd.)
 *   ERR|<poruka>     — greška na pojedinom layoutu (proces ipak nastavlja)
 *
 * Exit code: 0 = uspjeh (barem jedan layout kreiran), 1 = greška pri čitanju/pisanju
 */

using ACadSharp;
using ACadSharp.IO;
using DwgLayoutCreator;

// ── CLI args ────────────────────────────────────────────────────────────────
if (args.Length < 2)
{
    Console.Error.WriteLine(
        "Korištenje: dwg_layout_creator.exe <dwg_path> <sastau_dwg_path>");
    return 1;
}

string dwgPath  = args[0];
string sastPath = args[1];

if (!File.Exists(dwgPath))
{
    Console.Error.WriteLine($"DWG ne postoji: {dwgPath}");
    return 1;
}
if (!File.Exists(sastPath))
{
    Console.Error.WriteLine($"sastAu.dwg ne postoji: {sastPath}");
    return 1;
}

// ── Čitaj DWG ───────────────────────────────────────────────────────────────
CadDocument doc;
try
{
    Console.WriteLine($"INFO|Čitam {Path.GetFileName(dwgPath)}");
    using var reader = new DwgReader(dwgPath);
    doc = reader.Read();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Greška pri čitanju DWG: {ex.Message}");
    return 1;
}

// ── Skeniraj model space ────────────────────────────────────────────────────
var existingLayouts = ModelSpaceScanner.GetExistingLayoutNames(doc);
var candidates      = ModelSpaceScanner.Scan(doc);

Console.WriteLine($"INFO|Pronađeno {candidates.Count} blokova s layout imenima.");
Console.WriteLine($"INFO|Postojeći layouti: {string.Join(", ", existingLayouts)}");

if (candidates.Count == 0)
{
    Console.WriteLine("WARN|Nema kandidata — niti jedan blok A4L/A4P/.../CUSTOM ne odgovara.");
    return 0;
}

// ── Importiraj title block (jednom za sve layoute) ─────────────────────────
try
{
    Console.WriteLine($"INFO|Importiram title block {Path.GetFileName(sastPath)}");
    TitleBlockInserter.EnsureImported(doc, sastPath);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Greška pri importu title blocka: {ex.Message}");
    return 1;
}

// ── Kreiraj layoute ─────────────────────────────────────────────────────────
int created = 0;
int skipped = 0;
int errors  = 0;

foreach (var result in candidates)
{
    // Preskoči neoznačene
    if (result.LayoutName.Equals("X", StringComparison.OrdinalIgnoreCase))
    {
        Console.WriteLine($"WARN|{result.BlockName}: layout 'X' (neoznačen) — preskačem.");
        skipped++;
        continue;
    }

    // Preskoči postojeće
    if (existingLayouts.Contains(result.LayoutName, StringComparer.OrdinalIgnoreCase))
    {
        Console.WriteLine($"WARN|{result.BlockName}: layout '{result.LayoutName}' već postoji — preskačem.");
        skipped++;
        continue;
    }

    try
    {
        if (LayoutDefinitions.IsCustomBlock(result.BlockName))
        {
            // Treba postojati template layout s istim imenom (npr. "CUSTOM1")
            var template = doc.Layouts.FirstOrDefault(l =>
                l.Name.Equals(result.BlockName, StringComparison.OrdinalIgnoreCase));
            if (template is null)
            {
                Console.WriteLine($"WARN|Template layout '{result.BlockName}' ne postoji — preskačem '{result.LayoutName}'.");
                skipped++;
                continue;
            }

            var layout = LayoutFactory.CreateCustomLayout(doc, result, template);
            var (brN, brL, brLi) = ModelSpaceScanner.ParseLayoutNumber(result.LayoutName);

            double modelWidth = Math.Abs(result.MaxPt.X - result.MinPt.X);
            double scaledWidth = modelWidth / 10.0 / result.BlockScale;

            TitleBlockInserter.Insert(doc, layout, scaledWidth, 1.0, brN, brL, brLi, sastPath);
        }
        else
        {
            if (!LayoutDefinitions.StandardSizes.TryGetValue(result.BlockName, out var info))
            {
                Console.WriteLine($"ERR|Nema definicije za '{result.BlockName}' — preskačem.");
                errors++;
                continue;
            }

            var layout = LayoutFactory.CreateStandardLayout(doc, result, info);
            var (brN, brL, brLi) = ModelSpaceScanner.ParseLayoutNumber(result.LayoutName);

            var (printableW, _) = LayoutFactory.GetPrintableArea(info);
            double insertX = printableW - info.RightOffset;

            TitleBlockInserter.Insert(doc, layout, insertX, 1.0, brN, brL, brLi, sastPath);
        }

        existingLayouts.Add(result.LayoutName);
        created++;
        Console.WriteLine($"OK|Kreiran layout '{result.LayoutName}' ({result.BlockName})");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"ERR|Greška pri kreiranju '{result.LayoutName}': {ex.GetType().Name}: {ex.Message}");
        errors++;
    }
}

// ── Spremi DWG ──────────────────────────────────────────────────────────────
try
{
    Console.WriteLine($"INFO|Spremam {Path.GetFileName(dwgPath)}");
    doc.UpdateCollections(true, true);  // (createDictionaries, createDefaults)
    using var writer = new DwgWriter(dwgPath, doc);
    writer.Write();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Greška pri spremanju DWG: {ex.GetType().Name}: {ex.Message}");
    return 1;
}

// ── Sažetak ─────────────────────────────────────────────────────────────────
string level = errors > 0 ? "WARN" : "OK";
Console.WriteLine($"{level}|Gotovo — {created} kreirano, {skipped} preskočeno, {errors} grešaka.");
return errors > 0 && created == 0 ? 1 : 0;
