using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Zekon;
public record PackageManifest(int ReleaseNumber, string Version, string AddinFile, Dictionary<string, string> Sha256);
public record InstallState(string Current, string? Previous);
public static class Package
{
    public const string ReleasesUrl = "https://github.com/ZekonGJ/zekon-excel-tools/releases/latest";
    public static readonly string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ZekonTools");
    public static string ReleaseFolder(string version)
    {
        if (!Regex.IsMatch(version, @"\A[0-9]+\.[0-9]+\.[0-9]+(?:-[A-Za-z0-9.]+)?\z")) throw new InvalidDataException("Niepoprawny numer wersji.");
        return Path.Combine(Root, "releases", version);
    }
    public static PackageManifest Extract(Stream stream, string directory)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var manifestEntry = zip.GetEntry("manifest.json") ?? throw new InvalidDataException("Brak manifestu.");
        if (manifestEntry.Length > 65536) throw new InvalidDataException("Zbyt duzy manifest.");
        using var ms = manifestEntry.Open();
        var m = JsonSerializer.Deserialize<PackageManifest>(ms) ?? throw new InvalidDataException("Niepoprawny manifest.");
        ReleaseFolder(m.Version);
        if (m.ReleaseNumber < 1) throw new InvalidDataException("Niepoprawny numer wydania.");
        if (m.AddinFile != "ZekonTools.xlam" || m.Sha256.Count != 2 || !m.Sha256.ContainsKey(m.AddinFile) || !m.Sha256.ContainsKey("logo.bmp")) throw new InvalidDataException("Niepoprawna lista plikow.");
        if (zip.Entries.Count != 3) throw new InvalidDataException("Nieoczekiwane pliki w paczce.");
        Directory.CreateDirectory(directory);
        foreach (var item in m.Sha256)
        {
            if (item.Value.Length != 64 || !item.Value.All(Uri.IsHexDigit)) throw new InvalidDataException("Niepoprawna suma kontrolna.");
            var matches = zip.Entries.Where(e => e.FullName == item.Key).ToArray();
            if (matches.Length != 1 || matches[0].Length > 32 * 1024 * 1024) throw new InvalidDataException("Niepoprawny rozmiar lub powtorzona nazwa pliku.");
            string file = Path.Combine(directory, item.Key);
            matches[0].ExtractToFile(file, false);
            using var input = File.OpenRead(file);
            if (!Convert.ToHexString(SHA256.HashData(input)).Equals(item.Value, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Plik nie przeszedl kontroli integralnosci: " + item.Key);
        }
        File.WriteAllText(Path.Combine(directory, "manifest.json"), JsonSerializer.Serialize(m));
        return m;
    }
    public static PackageManifest VerifyInstalled(string version)
    {
        string dir = ReleaseFolder(version);
        var m = JsonSerializer.Deserialize<PackageManifest>(File.ReadAllText(Path.Combine(dir, "manifest.json"))) ?? throw new InvalidDataException("Brak manifestu wersji.");
        if (m.Version != version || m.AddinFile != "ZekonTools.xlam" || m.Sha256.Count != 2 || !m.Sha256.ContainsKey("logo.bmp") || !m.Sha256.ContainsKey(m.AddinFile)) throw new InvalidDataException("Niepoprawny manifest wersji.");
        foreach (var item in m.Sha256)
        {
            using var input = File.OpenRead(Path.Combine(dir, item.Key));
            if (!Convert.ToHexString(SHA256.HashData(input)).Equals(item.Value, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Zmieniony plik zainstalowanej wersji: " + item.Key);
        }
        return m;
    }
    public static bool IsOwnedEntry(string value)
    {
        var match = Regex.Match(value, "^\\s*(?:/R\\s+)?\"([^\"]+)\"\\s*$", RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        string path;
        try { path = Path.GetFullPath(match.Groups[1].Value); } catch { return false; }
        return path.StartsWith(Path.GetFullPath(Root) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            && Path.GetFileName(path).StartsWith("ZekonTools", StringComparison.OrdinalIgnoreCase)
            && Path.GetExtension(path).Equals(".xlam", StringComparison.OrdinalIgnoreCase);
    }
    public static string[] PlanEntries(IEnumerable<string> existing, string? newPath)
    {
        var values = existing.Where(v => !IsOwnedEntry(v)).ToList();
        if (newPath != null) values.Add("/R \"" + newPath + "\"");
        return values.ToArray();
    }
}
