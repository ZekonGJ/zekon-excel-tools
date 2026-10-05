using Microsoft.Win32;
using System.Text.RegularExpressions;
namespace Zekon;
public sealed class ExcelRegistration
{
    private record Snapshot(string Key, Dictionary<string, string> Entries);
    private readonly List<Snapshot> snapshots = new();
    private static bool IsOpen(string name) => Regex.IsMatch(name, @"\AOPEN(?:[1-9][0-9]*)?\z", RegexOptions.IgnoreCase);
    private static int Order(string name) => name.Length == 4 ? 0 : int.Parse(name[4..]);
    private static Dictionary<string, string> Read(RegistryKey key) => key.GetValueNames().Where(IsOpen).OrderBy(Order).ToDictionary(n => n, n => key.GetValue(n) as string ?? throw new InvalidDataException("Niepoprawny typ wpisu Excel OPEN."));
    private static void Write(RegistryKey key, IEnumerable<string> entries)
    {
        foreach (var name in key.GetValueNames().Where(IsOpen)) key.DeleteValue(name, false);
        int i = 0;
        foreach (var entry in entries) { key.SetValue(i == 0 ? "OPEN" : "OPEN" + i, entry, RegistryValueKind.String); i++; }
    }
    public void Apply(string? newPath)
    {
        if (System.Diagnostics.Process.GetProcessesByName("EXCEL").Length != 0) throw new InvalidOperationException("Excel zostal uruchomiony. Zamknij go i ponow operacje.");
        // Excel 2016, 2019, 2021, 2024 and Microsoft 365 use 16.0.
        var versions = new[] { "16.0", "15.0", "14.0" }.Where(v =>
        {
            using var k = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{v}\Excel");
            return k != null;
        }).ToArray();
        if (versions.Length == 0) throw new InvalidOperationException("Najpierw uruchom i zamknij Excel na tym koncie, a potem ponow instalacje.");
        try
        {
            foreach (var version in versions)
            {
                string path = $@"Software\Microsoft\Office\{version}\Excel\Options";
                using var key = Registry.CurrentUser.CreateSubKey(path, true);
                var old = Read(key);
                snapshots.Add(new(path, old));
                Write(key, Package.PlanEntries(old.Values, newPath));
            }
        }
        catch { Rollback(); throw; }
    }
    public void Rollback()
    {
        var failures = new List<Exception>();
        foreach (var s in snapshots.AsEnumerable().Reverse())
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(s.Key, true);
                foreach (var name in key.GetValueNames().Where(IsOpen)) key.DeleteValue(name, false);
                foreach (var item in s.Entries) key.SetValue(item.Key, item.Value, RegistryValueKind.String);
            }
            catch (Exception e) { failures.Add(e); }
        }
        snapshots.Clear();
        if (failures.Count > 0) throw new AggregateException("Nie udalo sie w pelni przywrocic rejestracji. Zachowaj log instalatora.", failures);
    }
}
