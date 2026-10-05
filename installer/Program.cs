using System.Diagnostics;
using System.Net.Http;
using System.Text.Json;
namespace Zekon;
internal static class Program
{
    [STAThread] static void Main()
    {
        ApplicationConfiguration.Initialize();
        using var mutex = new Mutex(true, "Local\\ZekonToolsInstaller", out bool alone);
        if (!alone) { MessageBox.Show("Instalator ZEKON jest juz uruchomiony."); return; }
        Application.Run(new SetupWindow());
    }
}
public sealed class SetupWindow : Form
{
    private readonly Label status = new() { AutoSize = false, Width = 540, Height = 110 };
    private readonly Button install = new() { Text = "Zainstaluj / Aktualizuj online", Width = 235, Height = 42 };
    private readonly Button rollback = new() { Text = "Przywroc poprzednia wersje", Width = 220, Height = 34 };
    private readonly Label installedVersion = new() { AutoSize = false, Width = 540, Height = 24, Location = new Point(26, 72) };
    private bool working;
    private readonly string stateFile = Path.Combine(Package.Root, "installation.json");
    public SetupWindow()
    {
        Text = "ZEKON - instalator / aktualizator 1.2.0"; Width = 610; Height = 425;
        StartPosition = FormStartPosition.CenterScreen; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;
        Font = new Font("Segoe UI", 10); BackColor = Color.White;
        var title = new Label { Text = "ZEKON | Narzedzia Excel", Font = new Font("Segoe UI", 22, FontStyle.Bold), ForeColor = Color.FromArgb(189,32,45), AutoSize = true, Location = new Point(24,24) };
        Controls.Add(title);
        installedVersion.Font = new Font("Segoe UI", 11, FontStyle.Bold); Controls.Add(installedVersion);
        status.Location = new Point(26,102); status.Height = 94; Controls.Add(status);
        install.Location = new Point(26,205); install.BackColor = Color.FromArgb(189,32,45); install.ForeColor = Color.White;
        rollback.Location = new Point(26,265);
        var updates = new Button { Text = "Aktualizuj z pliku...", Location = new Point(274,205), Width = 250, Height = 42 };
        var remove = new Button { Text = "Odlacz dodatek", Location = new Point(274,265), Width = 250, Height = 34 };
        Controls.AddRange(new Control[] {install,rollback,updates,remove});
        install.Click += (_,_) => Run(InstallLatest); rollback.Click += (_,_) => Run(Rollback, "Przywrócono poprzednią wersję");
        updates.Click += (_,_) => SelectUpdate();
        remove.Click += (_,_) => { if (MessageBox.Show("Odlaczyc dodatek ZEKON od Excela? Pliki i ustawienia pozostana.", "ZEKON", MessageBoxButtons.YesNo) == DialogResult.Yes) Run(Detach, "Dodatek odłączony"); };
        FormClosing += (_,e) => { if (working) e.Cancel = true; };
        RefreshState();
    }
    private void SelectUpdate()
    {
        if (working) return;
        using var dialog = new OpenFileDialog
        {
            Title = "Wybierz paczke aktualizacji ZEKON",
            Filter = "Aktualizacja ZEKON (*.zekonupdate)|*.zekonupdate",
            CheckFileExists = true, Multiselect = false
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        string path = dialog.FileName;
        Run(() =>
        {
            using var payload = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            InstallPayload(payload);
        });
    }
    private InstallState? State() => File.Exists(stateFile) ? JsonSerializer.Deserialize<InstallState>(File.ReadAllText(stateFile)) : null;
    private void RefreshInstalledVersion()
    {
        try { installedVersion.Text = "Zainstalowany dodatek: " + (State()?.Current ?? "brak"); }
        catch { installedVersion.Text = "Zainstalowany dodatek: nie mozna odczytac wersji"; }
    }
    private void RefreshState()
    {
        RefreshInstalledVersion();
        try { var s = State(); status.Text = s == null ? "Zapisz dokumenty i zamknij Excel.\nInstalacja dla biezacego uzytkownika, bez uprawnien administratora." : "Zainstalowana wersja: " + s.Current + "\nPrzed aktualizacja zapisz dokumenty i zamknij Excel."; rollback.Enabled = s?.Previous != null; }
        catch { status.Text = "Nie mozna odczytac stanu instalacji. Zachowaj plik installation.json do diagnostyki."; install.Enabled = false; rollback.Enabled = false; }
    }
    private async void Run(Action action, string successTitle = "Aktualizacja ukończona")
    {
        if (working) return;
        if (Process.GetProcessesByName("EXCEL").Length != 0) { MessageBox.Show("Zapisz dokumenty i zamknij wszystkie procesy Excel. Instalator nie zamyka ich automatycznie.","ZEKON"); return; }
        working = true; foreach (Control c in Controls) if (c is Button) c.Enabled = false;
        status.Text = "Trwa instalacja / aktualizacja. Prosze czekac...";
        try
        {
            await Task.Run(action);
            RefreshInstalledVersion();
            var current = State();
            string details = current == null
                ? "Dodatek ZEKON został odłączony od Excela. Pliki i ustawienia zostały zachowane."
                : "Zainstalowana wersja dodatku: " + current.Current + "\nMożesz teraz uruchomić Excel.";
            status.Text = successTitle + ".\n" + details;
            MessageBox.Show(this, details, "ZEKON — " + successTitle,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception e)
        {
            Directory.CreateDirectory(Package.Root);
            string log = Path.Combine(Package.Root, "installer-error.log");
            try { File.WriteAllText(log, DateTimeOffset.Now + "\n" + e); } catch { }
            status.Text = "Operacja nie zostala zakonczona.\n" + e.Message;
            MessageBox.Show(e.Message + "\n\nLog: " + log, "ZEKON", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { RefreshInstalledVersion(); working = false; foreach (Control c in Controls) if (c is Button) c.Enabled = true; try { rollback.Enabled = State()?.Previous != null; } catch { rollback.Enabled = false; } }
    }
    private void SaveState(InstallState state)
    {
        Directory.CreateDirectory(Package.Root);
        string tmp = stateFile + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(state));
        File.Move(tmp, stateFile, true);
    }
    private void InstallLatest()
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        byte[] bytes;
        try { bytes = UpdateFeed.DownloadAsync(client).GetAwaiter().GetResult(); }
        catch (Exception e) when (e is HttpRequestException || e is TaskCanceledException)
        {
            throw new IOException("Nie mozna pobrac najnowszej wersji. Sprawdz internet lub uzyj Aktualizuj z pliku. Nie zainstalowano starszej wersji.", e);
        }
        using var payload = new MemoryStream(bytes);
        InstallPayload(payload);
    }
    private void InstallPayload(Stream payload)
    {
        Directory.CreateDirectory(Package.Root);
        string stage = Path.Combine(Package.Root, "staging-" + Guid.NewGuid().ToString("N"));
        try
        {
            var manifest = Package.Extract(payload, stage);
            var previous = State();
            if (previous != null && Package.VerifyInstalled(previous.Current).ReleaseNumber > manifest.ReleaseNumber) throw new InvalidOperationException("To starsza wersja dodatku. Do powrotu uzyj Przywroc poprzednia wersje.");
            string destination = Package.ReleaseFolder(manifest.Version);
            if (Directory.Exists(destination))
            {
                var installed = Package.VerifyInstalled(manifest.Version);
                if (installed.Sha256.Any(kv => !manifest.Sha256.TryGetValue(kv.Key, out var hash) || hash != kv.Value)) throw new InvalidDataException("Pod tym numerem wersji znajduje sie inna zawartosc. Wymagany nowy numer wydania.");
            }
            else { Directory.CreateDirectory(Path.GetDirectoryName(destination)!); Directory.Move(stage, destination); }
            var registration = new ExcelRegistration();
            registration.Apply(Path.Combine(destination, manifest.AddinFile));
            try { SaveState(new(manifest.Version, previous?.Current == manifest.Version ? previous.Previous : previous?.Current)); }
            catch { registration.Rollback(); throw; }
            // A browser link supports both private and public repositories without storing tokens on PCs.
            try
            {
                string shortcuts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "ZEKON");
                Directory.CreateDirectory(shortcuts);
                File.WriteAllText(Path.Combine(shortcuts,"Aktualizacje ZEKON.url"), "[InternetShortcut]\r\nURL=" + Package.ReleasesUrl + "\r\n");
            }
            catch { /* Installation remains valid when Start Menu is managed by IT. */ }
        }
        finally { if (Directory.Exists(stage)) Directory.Delete(stage, true); }
    }
    private void Rollback()
    {
        var old = State() ?? throw new InvalidOperationException("Brak instalacji.");
        string target = old.Previous ?? throw new InvalidOperationException("Brak poprzedniej wersji zarzadzanej przez ten instalator.");
        var m = Package.VerifyInstalled(target);
        var registration = new ExcelRegistration();
        registration.Apply(Path.Combine(Package.ReleaseFolder(target), m.AddinFile));
        try { SaveState(new(target, old.Current)); } catch { registration.Rollback(); throw; }
    }
    private void Detach()
    {
        var registration = new ExcelRegistration(); registration.Apply(null);
        try { if (File.Exists(stateFile)) File.Delete(stateFile); } catch { registration.Rollback(); throw; }
    }
}
