# Narzędzia ZEKON — dodatek Excel i instalator Windows

Projekt: https://github.com/ZekonGJ/zekon-excel-tools

## Stan na 2026-10-05

Przeniesiono aktualne źródła VBA **1.0.0-rc2** i dodano kod natywnego instalatora
**ZekonSetup.exe** oraz GitHub Actions. Nie jest to jeszcze zatwierdzone wydanie
produkcyjne. Otrzymano XLAM rc2 oraz raport PASS z natywnego Excela (2026-10-05 08:58:38).
Instalator z tym dodatkiem wymaga jeszcze prób instalacji na stanowisku Windows/Excel.
Dostarczony wcześniej XLAM rc1c nie jest używany jako zamiennik.

Lokalnie wykonano testy Python blokujące wydanie niewłaściwych plików.
Kompilacja C#, testy planu rejestracji oraz testy pakietu są zdefiniowane w CI.
Testy interfejsu instalatora, wpisów rejestru i działania dodatku wymagają Windows
z desktopowym Excelem. Zielone CI bez tych prób nie jest potwierdzeniem poprawnej
instalacji na stanowisku produkcyjnym.

## Praca użytkownika po zatwierdzeniu wydania

1. Zapisz dokumenty i zamknij Excel.
2. Pobierz `ZekonSetup.exe` z firmowego wydania i uruchom.
3. Kliknij **Zainstaluj / Aktualizuj** i otwórz Excel.

Kolejne wersje instaluje się tym samym sposobem. W menu Start powstaje skrót
**ZEKON → Aktualizacje ZEKON** do ostatniego stabilnego wydania GitHub.
Można też rozesłać jeden EXE wszystkim stanowiskom — instalacja działa offline.
W repozytorium prywatnym do pobierania z GitHub potrzebne jest uprawnione konto;
instalator nie zapisuje tokenów GitHub i nie udostępnia prywatnego kodu publicznie.
Nie ma automatycznego pobierania ani instalowania w tle.

Instalator jest samodzielną aplikacją .NET 10 dla Windows x64, również dla
32-bitowego Excela na 64-bitowym Windows. Nie wymaga osobnej instalacji .NET,
PowerShella, dostępu do projektu VBA ani interfejsu Office Interop/Excel COM.
Rejestruje gotowy XLAM we wpisach startowych Excela dla bieżącego użytkownika.
Nie ustawia zaufanych lokalizacji, nie wyłącza ochrony makr i nie zmienia polityk IT.

## Aktualizacja i wycofanie

- Każda wersja trafia do `%LOCALAPPDATA%\ZekonTools\releases\<wersja>`.
- Istniejąca wersja nie jest nadpisywana inną zawartością.
- Procesy Excel muszą być zamknięte. Instalator ich nie zabija.
- Inne dodatki i profile ZEKON pozostają zachowane.
- Po aktualizacji przycisk **Przywróć poprzednią wersję** przełącza do poprzedniej
  wersji zarządzanej przez ten instalator. Pierwsza migracja ze starych instalatorów
  rc1/rc2 nie zapewnia automatycznego rollbacku do ich plików; pozostają na dysku.
- **Odłącz dodatek** usuwa jego rejestrację; zachowuje pliki i ustawienia użytkownika.
- W razie błędu rejestracji wykonywana jest próba odtworzenia wpisów sprzed operacji.
  Jeśli także ona się nie powiedzie, komunikat wskazuje błąd i lokalny log.
- Stan: `%LOCALAPPDATA%\ZekonTools\installation.json`.
  Błędy: `%LOCALAPPDATA%\ZekonTools\installer-error.log`.

## Przygotowanie kolejnej wersji — tylko osoba utrzymująca projekt

1. Zmień kod VBA, nadaj nowy numer wersji i uzupełnij `addin/HISTORIA_WERSJI.md`.
   Zachowaj zgodność APP_VERSION, tytułu, generatorów i skryptu budowania.
2. Zbuduj dodatek w kontrolowanym Windows z licencjonowanym Excelem.
   `addin/Build-Zekon.ps1` to dotychczasowa pomocnicza metoda deweloperska;
   komputery użytkowników jej nie uruchamiają. Alternatywnie generator
   `addin/tools/generate_builder.py` tworzy instalator VBA dla lokalnego Excela.
3. Przejdź testy natywne oraz sprawdź rzeczywiste dane. Zachowaj plik XLAM i
   `test-result.txt` z tego samego budowania. Podpisz projekt VBA, jeśli firma
   korzysta z certyfikatu. Nie zmieniaj kodu po podpisaniu.
4. Dodaj do `release-input/` pliki `ZekonTools.xlam` i `test-result.txt`.
   Nie dodawaj prywatnych skoroszytów, haseł, certyfikatów z kluczem ani tokenów.
5. GitHub **Actions → Build complete ZekonSetup → Run workflow**.
   Podaj wersję i rosnący `release_number`; nie używaj ponownie numerów.
6. Pobierz artefakt `ZekonSetup-review`. Przetestuj instalację, migrację,
   aktualizację, wycofanie i odłączenie na testowym stanowisku Windows/Excel.
7. Podpisz EXE firmowym certyfikatem podpisu kodu i znacznikiem czasu, jeżeli
   jest dostępny. Po podpisaniu oblicz ponownie SHA-256; hash z CI dotyczy pliku
   przed podpisaniem. Klucz pozostaje poza repozytorium.
8. Utwórz zatwierdzone wydanie GitHub z EXE, sumą kontrolną i opisem zmian.
   Stabilne wydanie pojawi się pod adresem `/releases/latest`.
   Nie publikuj wersji do użytkowania przed zakończeniem testów.

GitHub Actions buduje instalator; nie zakłada obecności licencjonowanego Excela
na standardowych runnerach. W przyszłości budowanie VBA można przenieść na
wydzielony runner Windows/Excel kontrolowany przez firmę. Nie jest on skonfigurowany.

## Podpisy i ostrzeżenia Windows

Obecny kod nie ma certyfikatu wydawcy. Brak PowerShella nie oznacza braku alertów.
Podpis instalatora, reputacja SmartScreen i polityki stanowiska są niezależne od
samego GitHuba. Podpis VBA jest oddzielny od podpisu EXE. Nie gwarantujemy instalacji
bez ostrzeżeń; zatwierdzenie wydawcy przez IT może być nadal potrzebne.

## Testy

```
python -m unittest discover -s tests -v
dotnet run --project installer.Tests/Installer.Tests.csproj --configuration Release
dotnet build installer/ZekonSetup.csproj --configuration Release
```

`dotnet build` bez payloadu służy tylko sprawdzeniu kompilacji. Takiego pliku nie
wolno dystrybuować. `dotnet publish` i workflow wydania wymagają payloadu z XLAM.

## Źródła techniczne

- https://learn.microsoft.com/en-us/platform/support/ (polityki platform Microsoft)
- https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core
- https://support.microsoft.com/en-us/excel/excel-com-add-ins-and-automation-add-ins
- https://docs.github.com/en/actions/concepts/workflows-and-actions/workflow-artifacts
