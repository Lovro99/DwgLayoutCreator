using System.Text.Json;
using Autodesk.AutoCAD.EditorInput;

namespace LayoutCreator;

// Jedan red naslova (sheet "Nacrti": A=layout, B=naslov, C=mjerilo).
internal sealed class NaslovRow
{
    public string Layout = "";
    public string Naslov = "";
    public string Mjerilo = "";
}

// Podaci koje orkestrator prosljeduje kroz LAYOUT_BATCH_JSON. Kljucevi su
// prisutni samo za ukljucene korake (plan §2).
internal sealed class BatchData
{
    // Redoslijed ocuvan (JSON object order) -> lista parova umjesto Dictionary.
    public List<KeyValuePair<string, string>>? Polja;
    public List<NaslovRow>? Naslovi;
}

// Cita LAYOUT_BATCH_JSON (env var -> put do UTF-8 JSON datoteke koju orkestrator
// zapise uz kopiju DWG-a). Struktura:
//   { "polja":   {"KLJUC": "vrijednost", ...},
//     "naslovi": [{"layout": "...", "naslov": "...", "mjerilo": "..."}, ...] }
// System.Text.Json je in-box u .NET 8 (bez NuGeta). Kompajlira se SAMO u
// LayoutCreatorCore.
internal static class BatchDataFile
{
    // Ucitaj batch podatke. Vrati null i ispisi ERR| ako env var/datoteka fali
    // ili se JSON ne moze parsirati.
    internal static BatchData? Load(Editor ed)
    {
        string? path = Environment.GetEnvironmentVariable("LAYOUT_BATCH_JSON");
        if (string.IsNullOrWhiteSpace(path))
        {
            ed.WriteMessage("\nERR| LAYOUT_BATCH_JSON nije postavljen.");
            return null;
        }
        if (!System.IO.File.Exists(path))
        {
            ed.WriteMessage($"\nERR| LAYOUT_BATCH_JSON datoteka ne postoji: {path}");
            return null;
        }

        try
        {
            byte[] bytes = System.IO.File.ReadAllBytes(path);
            using var doc = JsonDocument.Parse(bytes);
            JsonElement root = doc.RootElement;
            var data = new BatchData();

            if (root.TryGetProperty("polja", out JsonElement polja) &&
                polja.ValueKind == JsonValueKind.Object)
            {
                data.Polja = new List<KeyValuePair<string, string>>();
                foreach (JsonProperty prop in polja.EnumerateObject())
                    data.Polja.Add(new KeyValuePair<string, string>(
                        prop.Name, JsonToString(prop.Value)));
            }

            if (root.TryGetProperty("naslovi", out JsonElement naslovi) &&
                naslovi.ValueKind == JsonValueKind.Array)
            {
                data.Naslovi = new List<NaslovRow>();
                foreach (JsonElement el in naslovi.EnumerateArray())
                {
                    data.Naslovi.Add(new NaslovRow
                    {
                        Layout = GetStr(el, "layout"),
                        Naslov = GetStr(el, "naslov"),
                        Mjerilo = GetStr(el, "mjerilo"),
                    });
                }
            }

            return data;
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nERR| ne mogu procitati LAYOUT_BATCH_JSON: {ex.Message}");
            return null;
        }
    }

    private static string GetStr(JsonElement obj, string key)
    {
        if (obj.ValueKind == JsonValueKind.Object &&
            obj.TryGetProperty(key, out JsonElement v))
            return JsonToString(v);
        return "";
    }

    // Orkestrator normalizira vrijednosti u stringove; ovo je obrambeni sloj za
    // slucaj da JSON nosi broj/bool.
    private static string JsonToString(JsonElement v)
    {
        return v.ValueKind switch
        {
            JsonValueKind.String => v.GetString() ?? "",
            JsonValueKind.Number => v.GetRawText(),
            JsonValueKind.True => "TRUE",
            JsonValueKind.False => "FALSE",
            JsonValueKind.Null => "",
            _ => v.GetRawText(),
        };
    }
}
