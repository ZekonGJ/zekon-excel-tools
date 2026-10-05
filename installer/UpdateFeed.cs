using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
namespace Zekon;
public record UpdateOffer(int SchemaVersion, int ReleaseNumber, string Version, string Sha256);
public static class UpdateFeed
{
    public const string BaseUrl = "https://raw.githubusercontent.com/ZekonGJ/zekon-excel-tools/main/updates/";
    private const int Limit = 32 * 1024 * 1024;
    public static UpdateOffer Parse(string json)
    {
        if (json.Length > 16384) throw new InvalidDataException("Zbyt duzy katalog aktualizacji.");
        var offer = JsonSerializer.Deserialize<UpdateOffer>(json) ?? throw new InvalidDataException("Brak informacji o najnowszej wersji.");
        if (offer.SchemaVersion != 1 || offer.ReleaseNumber < 1) throw new InvalidDataException("Nieobslugiwany katalog aktualizacji.");
        Package.ReleaseFolder(offer.Version);
        if (offer.Sha256 == null || offer.Sha256.Length != 64 || !offer.Sha256.All(Uri.IsHexDigit)) throw new InvalidDataException("Niepoprawna suma paczki aktualizacji.");
        return offer;
    }
    public static void Verify(UpdateOffer offer, byte[] bytes)
    {
        if (bytes.Length > Limit || !Convert.ToHexString(SHA256.HashData(bytes)).Equals(offer.Sha256, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Pobrana paczka nie zgadza sie z katalogiem aktualizacji.");
        using var stream = new MemoryStream(bytes);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = zip.GetEntry("manifest.json") ?? throw new InvalidDataException("Brak manifestu paczki.");
        if (entry.Length > 65536) throw new InvalidDataException("Zbyt duzy manifest paczki.");
        using var input = entry.Open();
        var manifest = JsonSerializer.Deserialize<PackageManifest>(input) ?? throw new InvalidDataException("Niepoprawny manifest paczki.");
        if (manifest.Version != offer.Version || manifest.ReleaseNumber != offer.ReleaseNumber)
            throw new InvalidDataException("Pobrano inna wersje niz wskazana w katalogu aktualizacji.");
    }
    public static async Task<byte[]> DownloadAsync(HttpClient client)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        using var request = new HttpRequestMessage(HttpMethod.Get, BaseUrl + "latest.json?check=" + Guid.NewGuid().ToString("N"));
        request.Headers.CacheControl = new CacheControlHeaderValue { NoCache = true, NoStore = true };
        using var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
        response.EnsureSuccessStatusCode();
        byte[] catalog = await ReadLimited(response, 16384, timeout.Token);
        var offer = Parse(System.Text.Encoding.UTF8.GetString(catalog));
        string url = BaseUrl + "ZekonTools_" + offer.Version + ".zekonupdate";
        using var package = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
        package.EnsureSuccessStatusCode();
        byte[] bytes = await ReadLimited(package, Limit, timeout.Token);
        Verify(offer, bytes);
        return bytes;
    }
    private static async Task<byte[]> ReadLimited(HttpResponseMessage response, int limit, CancellationToken token)
    {
        if (response.Content.Headers.ContentLength > limit) throw new InvalidDataException("Zbyt duzy plik aktualizacji.");
        using var input = await response.Content.ReadAsStreamAsync(token);
        using var output = new MemoryStream();
        byte[] buffer = new byte[81920];
        int n;
        while ((n = await input.ReadAsync(buffer.AsMemory(), token)) > 0)
        {
            if (output.Length + n > limit) throw new InvalidDataException("Zbyt duzy plik aktualizacji.");
            output.Write(buffer, 0, n);
        }
        return output.ToArray();
    }
}
