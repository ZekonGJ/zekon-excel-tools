using Zekon;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
static void Check(bool condition, string name) { if (!condition) throw new Exception(name); }
string oldPath = Path.Combine(Package.Root, "1.0.0-rc2", "ZekonTools_1.0.0-rc2.xlam");
string nextPath = Path.Combine(Package.ReleaseFolder("1.0.0-rc3"), "ZekonTools.xlam");
string other = "/R \"C:\\OtherCompany\\keep.xlam\"";
var plan = Package.PlanEntries(new[] { other, "/R \"" + oldPath + "\"" }, nextPath);
Check(plan.Length == 2 && plan[0] == other && plan[1].Contains(nextPath), "Migrate only owned add-ins");
Check(Package.PlanEntries(plan, null).SequenceEqual(new[] { other }), "Detach preserves foreign add-ins");
Check(!Package.IsOwnedEntry("/R \"" + Package.Root + "-other/ZekonTools.xlam\""), "Directory boundary");
try { Package.ReleaseFolder("../outside"); throw new Exception("Traversal accepted"); } catch (InvalidDataException) { }
static MemoryStream MakePackage(bool corrupt = false, bool extra = false, string version = "1.0.0-rc2", int release = 2)
{
    byte[] addin = new byte[] {1,2,3}, logo = new byte[] {4,5};
    var manifest = new PackageManifest(release,version,"ZekonTools.xlam",new() {
        ["ZekonTools.xlam"] = Convert.ToHexString(SHA256.HashData(addin)), ["logo.bmp"] = Convert.ToHexString(SHA256.HashData(logo)) });
    var stream = new MemoryStream();
    using (var zip = new ZipArchive(stream,ZipArchiveMode.Create,true))
    {
        using (var writer = new StreamWriter(zip.CreateEntry("manifest.json").Open())) writer.Write(JsonSerializer.Serialize(manifest));
        using (var entry = zip.CreateEntry("ZekonTools.xlam").Open()) entry.Write(corrupt ? new byte[] {9} : addin);
        using (var entry = zip.CreateEntry("logo.bmp").Open()) entry.Write(logo);
        if (extra) zip.CreateEntry("../escape");
    }
    stream.Position=0; return stream;
}
foreach (var transport in new[] {"embedded", "external-file"})
foreach (var mode in new[] {"valid","tampered","extra"})
{
    string stage = Path.Combine(Path.GetTempPath(), "zekon-test-" + Guid.NewGuid());
    string updateFile = stage + ".zekonupdate";
    try
    {
        using var generated = MakePackage(mode=="tampered",mode=="extra");
        if (transport == "external-file") File.WriteAllBytes(updateFile, generated.ToArray());
        using Stream payload = transport == "external-file"
            ? new FileStream(updateFile, FileMode.Open, FileAccess.Read, FileShare.Read)
            : generated;
        try { Package.Extract(payload,stage); Check(mode=="valid","Invalid package accepted"); }
        catch (InvalidDataException) { Check(mode!="valid","Valid package rejected"); }
    }
    finally { if(Directory.Exists(stage)) Directory.Delete(stage,true); if(File.Exists(updateFile)) File.Delete(updateFile); }
}
Console.WriteLine("PASS: registration plan, ownership boundaries, traversal, hash validation, payload allowlist for embedded and external update files.");

// The same updater discovers a newer feed without a new executable or installed state.
foreach (var release in new[] { 3, 4 })
{
    string version = release == 3 ? "1.0.0-rc2a" : "1.0.0-rc2b";
    using var generated = MakePackage(version: version, release: release);
    byte[] bytes = generated.ToArray();
    var offer = new UpdateOffer(1, release, version, Convert.ToHexString(SHA256.HashData(bytes)));
    using var client = new System.Net.Http.HttpClient(new FeedHandler(offer, bytes));
    Check(UpdateFeed.DownloadAsync(client).GetAwaiter().GetResult().SequenceEqual(bytes), "Latest feed must determine installed package");
    try { UpdateFeed.Verify(offer with { Version = "9.0.0" }, bytes); throw new Exception("Wrong manifest version accepted"); } catch (InvalidDataException) { }
    bytes[0] ^= 1;
    try { UpdateFeed.Verify(offer, bytes); throw new Exception("Corrupted download accepted"); } catch (InvalidDataException) { }
}
try { UpdateFeed.Parse(JsonSerializer.Serialize(new UpdateOffer(1, 3, "../escape", new string('A',64)))); throw new Exception("Feed traversal accepted"); } catch (InvalidDataException) { }
using (var offline = new System.Net.Http.HttpClient(new FeedHandler(null, Array.Empty<byte>())))
{
    try { UpdateFeed.DownloadAsync(offline).GetAwaiter().GetResult(); throw new Exception("Offline install should fail, not use stale fallback"); } catch (System.Net.Http.HttpRequestException) { }
}
Console.WriteLine("PASS: latest-feed versions, download hash, version consistency, traversal, offline failure without stale fallback.");
// Network-folder deployment uses the supplied executable folder, not working directory.
string folder = Path.Combine(Path.GetTempPath(), "zekon-local-" + Guid.NewGuid());
Directory.CreateDirectory(folder);
try
{
    try { LocalUpdate.ReadLatest(folder); throw new Exception("Missing local package accepted"); } catch (IOException) { }
    using var older = MakePackage(version: "1.0.0-old", release: 8);
    using var newer = MakePackage(version: "1.0.0-new", release: 9);
    File.WriteAllBytes(Path.Combine(folder,"z-old.zekonupdate"), older.ToArray());
    File.WriteAllBytes(Path.Combine(folder,"a-new.zekonupdate"), newer.ToArray());
    File.WriteAllText(Path.Combine(folder,"newest.zekonupdate.part"), "unfinished transfer");
    var bytes = LocalUpdate.ReadLatest(folder);
    Check(bytes.SequenceEqual(newer.ToArray()), "Choose manifest release number, not filename or current folder");
    File.Delete(Path.Combine(folder,"a-new.zekonupdate"));
    using (var cached = new MemoryStream(bytes))
    {
        string stage = Path.Combine(folder,"stage");
        Check(Package.Extract(cached,stage).ReleaseNumber == 9,"Local copy survives source disappearance");
        Directory.Delete(stage,true);
    }
    File.WriteAllBytes(Path.Combine(folder,"a-new.zekonupdate"), newer.ToArray());
    File.WriteAllBytes(Path.Combine(folder,"duplicate.zekonupdate"), newer.ToArray());
    try { LocalUpdate.ReadLatest(folder); throw new Exception("Ambiguous release accepted"); } catch (InvalidDataException) { }
    File.Delete(Path.Combine(folder,"duplicate.zekonupdate"));
    using (var locked = new FileStream(Path.Combine(folder,"a-new.zekonupdate"),FileMode.Open,FileAccess.ReadWrite,FileShare.None))
    {
        try { LocalUpdate.ReadLatest(folder); throw new Exception("Locked package accepted"); } catch (IOException) { }
    }
    using var bad = MakePackage(corrupt:true,version:"1.0.0-new",release:9);
    File.WriteAllBytes(Path.Combine(folder,"a-new.zekonupdate"),bad.ToArray());
    using var corrupt = new MemoryStream(LocalUpdate.ReadLatest(folder));
    try { Package.Extract(corrupt,Path.Combine(folder,"badstage")); throw new Exception("Tampered local package accepted"); } catch (InvalidDataException) { }
    File.WriteAllText(Path.Combine(folder,"a-new.zekonupdate"),"partial zip");
    try { LocalUpdate.ReadLatest(folder); throw new Exception("Broken package silently downgraded"); } catch (InvalidDataException) { }
}
finally { Directory.Delete(folder,true); }
Console.WriteLine("PASS: local folder, manifest selection, partial files, cached bytes, duplicate release, file lock, corruption, no stale fallback.");
sealed class FeedHandler(UpdateOffer? offer, byte[] package) : System.Net.Http.HttpMessageHandler
{
    protected override Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request, CancellationToken token)
    {
        if (offer == null) throw new System.Net.Http.HttpRequestException("Offline test");
        string expected = UpdateFeed.BaseUrl + "ZekonTools_" + offer.Version + ".zekonupdate";
        bool catalog = request.RequestUri!.AbsoluteUri.StartsWith(UpdateFeed.BaseUrl + "latest.json?check=", StringComparison.Ordinal);
        if (!catalog && request.RequestUri.AbsoluteUri != expected) throw new Exception("Unexpected update URL");
        return Task.FromResult(new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK) {
            Content = catalog ? new System.Net.Http.StringContent(JsonSerializer.Serialize(offer)) : new System.Net.Http.ByteArrayContent(package)
        });
    }
}

