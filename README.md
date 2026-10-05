# Narzędzia ZEKON — dodatek Excel i instalator Windows

Projekt: https://github.com/ZekonGJ/zekon-excel-tools

## Poprawka rozbijania z filtrem: 1.0.0-rc2a

Źródła zawierają poprawkę kopiowania przy aktywnym filtrze. Pakiet developerski
`support/ZEKON_BUDUJ_RC2A_B04.zip` buduje tę wersję i uruchamia nowe testy Excel.
Otrzymano XLAM rc2a i raport PASS z 2026-10-05 10:36:21, w tym nowe testy
rozbijania pod filtrem. Wszystkie 8 modułów zgodne ze źródłami; pakiet release-input
zawiera rc2a. Instalator wymaga jeszcze sprawdzenia aktualizacji na stanowisku użytkownika.

## Stan rc2 na 2026-10-05

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

## Aktualizator 1.1.0 i osobne paczki dodatku

Jednorazowo pobierz nowy ZekonSetup.exe (aktualizator 1.1.0). Dotychczasowy EXE
nie ma obsługi zewnętrznych paczek. Zachowaj nowy program w stałym miejscu.
Numer programu aktualizującego i numer dodatku są niezależne.

Kolejne aktualizacje dodatku:
1. Pobierz mały plik `ZekonTools_<wersja>.zekonupdate`. Nie rozpakowuj go.
2. Zapisz dokumenty i zamknij Excel.
3. Uruchom zachowany `ZekonSetup.exe`, wybierz **Aktualizuj z pliku...** i wskaż paczkę.
4. Otwórz Excel i sprawdź numer wersji w panelu ZEKON.

Przycisk **Zainstaluj dołączony dodatek** instaluje wersję wbudowaną w EXE;
nie służy do wczytywania kolejnych aktualizacji. Nowe EXE jest potrzebne tylko
przy zmianie samego aktualizatora lub formatu paczek, nie przy zwykłej poprawce VBA.
Plik aktualizacji można przekazać na inne stanowiska i zastosować offline.
Paczka zawiera XLAM, logo oraz manifest z numerem i sumami kontrolnymi. Jest
sprawdzana tym samym mechanizmem co pakiet wbudowany. Sumy kontrolne wykrywają
uszkodzenia, nie zastępują podpisu wydawcy. Pobieraj paczki z firmowego źródła.

Program działa na Windows x64, również przy 32-bitowym Excelu. Nie wymaga
PowerShella ani osobnego instalowania .NET. Nie zmienia zabezpieczeń makr.

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
