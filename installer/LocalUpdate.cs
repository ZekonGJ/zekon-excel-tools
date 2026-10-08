using System.IO.Compression;
using System.Text.Json;
namespace Zekon;

public static class LocalUpdate
{
    private const int Limit = 32 * 1024 * 1024;
    public static byte[] ReadLatest(string directory)
    {
        // No current-directory or mapped-drive assumptions. Only completed package names.
        var paths = Directory.GetFiles(directory, "*.zekonupdate", SearchOption.TopDirectoryOnly);
        if (paths.Length == 0) throw new IOException("Brak paczki .zekonupdate obok ZekonSetup.exe. Popros administratora o umieszczenie aktualizacji w tym folderze.");
        var candidates = new List<(string Path, PackageManifest Manifest)>();
        foreach (string path in paths)
        {
            using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (input.Length > Limit) throw new InvalidDataException("Zbyt duza paczka: " + Path.GetFileName(path));
            candidates.Add((path, ReadManifest(input)));
        }
        int newest = candidates.Max(x => x.Manifest.ReleaseNumber);
        var latest = candidates.Where(x => x.Manifest.ReleaseNumber == newest).ToArray();
        if (latest.Length != 1) throw new InvalidDataException("Kilka paczek ma ten sam najnowszy numer wydania. Administrator musi pozostawic jedna paczke tego wydania.");
        // Read the complete package locally before installation or changing registration.
        using var source = new FileStream(latest[0].Path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var copy = new MemoryStream();
        byte[] buffer = new byte[81920];
        int n;
        while ((n = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            if (copy.Length + n > Limit) throw new InvalidDataException("Zbyt duza paczka aktualizacji.");
            copy.Write(buffer, 0, n);
        }
        copy.Position = 0;
        var actual = ReadManifest(copy);
        var expected = latest[0].Manifest;
        if (actual.Version != expected.Version || actual.ReleaseNumber != expected.ReleaseNumber || actual.Sha256.Any(kv => !expected.Sha256.TryGetValue(kv.Key, out var hash) || !string.Equals(hash, kv.Value, StringComparison.OrdinalIgnoreCase)))
            throw new IOException("Paczka zmienila sie podczas odczytu. Uruchom instalator ponownie po zakonczeniu kopiowania.");
        return copy.ToArray();
    }
    private static PackageManifest ReadManifest(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, true);
        var entries = zip.Entries.Where(x => x.FullName == "manifest.json").ToArray();
        if (zip.Entries.Count != 3 || entries.Length != 1 || entries[0].Length > 65536) throw new InvalidDataException("Niepoprawna paczka aktualizacji.");
        using var input = entries[0].Open();
        var m = JsonSerializer.Deserialize<PackageManifest>(input) ?? throw new InvalidDataException("Brak manifestu.");
        Package.ReleaseFolder(m.Version);
        if (m.ReleaseNumber < 1 || m.AddinFile != "ZekonTools.xlam" || m.Sha256 == null || m.Sha256.Count != 2 || !m.Sha256.ContainsKey("ZekonTools.xlam") || !m.Sha256.ContainsKey("logo.bmp")) throw new InvalidDataException("Niepoprawny manifest.");
        foreach (var kv in m.Sha256)
            if (kv.Value == null || kv.Value.Length != 64 || !kv.Value.All(Uri.IsHexDigit) || zip.Entries.Count(e => e.FullName == kv.Key) != 1) throw new InvalidDataException("Niepoprawna zawartosc manifestu.");
        return m;
    }
}
