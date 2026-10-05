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
static MemoryStream MakePackage(bool corrupt = false, bool extra = false)
{
    byte[] addin = new byte[] {1,2,3}, logo = new byte[] {4,5};
    var manifest = new PackageManifest(2,"1.0.0-rc2","ZekonTools.xlam",new() {
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
