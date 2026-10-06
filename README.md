# Narzędzia ZEKON — dodatek Excel i instalator Windows

Projekt: https://github.com/ZekonGJ/zekon-excel-tools

## Przygotowywana poprawka: 1.0.0-rc2d — eksporty ocynku

Status: źródła i testy przygotowane; wymagany natywny XLAM i raport Windows Excel.
Katalog aktualizacji nadal wskazuje przetestowaną rc2c. Nie opublikowano rc2d.

„Eksport: ocynk” i „Sztuki + ocynk” zapisują `wzor.xls` w wybranym folderze,
w formacie Excel 97–2003 i układzie dostarczonego `wzor1.xls`. Wbudowany wzorzec
zawiera tylko nagłówki: usunięto przykładowe dane i metadane dokumentu. Zachowano
rekordy formatowania i ustawienia wydruku. Dane zaczynają się od wiersza 2.
Kolumna H „Zekon Unterlieferanten” pozostaje pusta pod nagłówkiem.
Istniejący `wzor.xls` nie jest nadpisywany. Przed rozbiciem sprawdzany jest limit
65535 pozycji XLS. Oba eksporty malowania zachowują wcześniejsze zachowanie.

Próba deweloperska B07: `support/ZEKON_BUDUJ_RC2D_B07.zip`, makro
`ZEKON_BUDUJ_RC2D_B07`. Buduje nowy XLAM i uruchamia test zapisu oraz ponownego
otwarcia `wzor.xls`, w tym sekwencji rozbijania pod filtrem. To narzędzie budowania
jednego wydania, nie instalator stanowiskowy. Publikacja po PASS: wydanie 6,
`ZekonTools_1.0.0-rc2d.zekonupdate`, bez wymiany aktualizatora 1.2.0.

Zgłoszony alarm AVG dotyczący instalatora pozostaje niewyjaśniony. Ta poprawka
eksportu nie jest weryfikacją bezpieczeństwa EXE ani rozwiązaniem tego alarmu.

## Aktualny dodatek: 1.0.0-rc2c

Opublikowany XLAM zawiera wszystkie dotychczasowe poprawki, rozbijanie pod filtrem
oraz stały napis wersji pobierany z APP_VERSION. Raport Excel z 2026-10-05 11:23:04
potwierdza PASS, także test wersji panelu. Porównano wszystkie osiem modułów VBA
z aktualnymi źródłami. Aktualizator online wskazuje rc2c (wydanie 5).

## Aktualizator 1.2.0 — pierwsza instalacja zawsze z aktualnego katalogu

Przy pierwszym uruchomieniu na nowym komputerze i przy kolejnej aktualizacji
używaj **Zainstaluj / Aktualizuj online**. Program pobiera `updates/latest.json`
z tego repozytorium, a następnie wskazaną wersją paczkę `.zekonupdate`.
Nie zawiera wbudowanej starej wersji dodatku. Brak internetu lub błąd pobierania
kończy operację komunikatem; nie powoduje instalacji starej wersji.

Na 2026-10-05 katalog wskazuje **1.0.0-rc2c**, ze zweryfikowanym XLAM i raportem
Excel PASS. Aktualizator 1.2.0 pobiera tę paczkę bez wymiany EXE.

Po udanej operacji wyświetla się osobne okno potwierdzenia i numer dodatku.
Numer jest też stale widoczny w aktualizatorze, również podczas pobierania.
Numer programu aktualizującego (1.2.0) jest niezależny od numeru dodatku.

Przy braku internetu można użyć **Aktualizuj z pliku...** i wskazać wcześniej
pobraną paczkę. Nie rozpakowuj pliku `.zekonupdate`. Taki wybór instaluje wersję
z wybranej paczki, a nie automatycznie najnowszą wersję internetową.

Zapisz dokumenty i zamknij Excel przed operacją. Program działa na Windows x64,
również z 32-bitowym Excelem, nie wymaga PowerShella ani instalowania .NET.
Sumy kontrolne wykrywają uszkodzenia; nie są podpisem wydawcy.

### Publikacja kolejnej wersji dodatku

1. Zbuduj i przetestuj nowy XLAM w Excelu. Sprawdź zgodność ze źródłami.
2. Użyj `tools/package_release.py` z nową wersją i rosnącym numerem wydania.
3. Dodaj wynik jako `updates/ZekonTools_<wersja>.zekonupdate`, bez nadpisywania starszych paczek.
4. Zapisz `updates/proof_<wersja>.json` z sumą XLAM i rzeczywistym raportem testów.
5. W tej samej zmianie ustaw `updates/latest.json`: SchemaVersion=1, ReleaseNumber,
   Version oraz Sha256 całej paczki. Sprawdź `python tools/verify_update_feed.py`.
6. Od tej chwili ten sam aktualizator pobierze nową wersję na wszystkich stanowiskach,
   kiedy użytkownik kliknie przycisk online. Nie jest to automatyczna instalacja w tle.

Nowe EXE jest potrzebne przy zmianach samego aktualizatora, nie przy zwykłej poprawce VBA.
Przejście ze starszego aktualizatora wymaga jednorazowego zastąpienia EXE wersją 1.2.0.

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

## Budowanie aktualizatora

GitHub Actions buduje program Windows oraz sprawdza opublikowany katalog paczek.
`tools/verify_update_feed.py` weryfikuje SHA256 paczki i jej zawartości, wersję
XLAM oraz powiązanie z zapisanym raportem Excel. Publikowanie kolejnego dodatku
nie wymaga przebudowania aktualizatora. Kontrolowany Windows z Excelem nadal
jest potrzebny do zbudowania i natywnego przetestowania zmienionego VBA.

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

Aktualizator 1.2.0 nie osadza XLAM. Budowanie wydania wymaga poprawnego katalogu
opublikowanych, przetestowanych paczek. Testy UI i instalacji na stanowisku Windows
pozostają odrębnym etapem odbioru.

## Źródła techniczne

- https://learn.microsoft.com/en-us/platform/support/ (polityki platform Microsoft)
- https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core
- https://support.microsoft.com/en-us/excel/excel-com-add-ins-and-automation-add-ins
- https://docs.github.com/en/actions/concepts/workflows-and-actions/workflow-artifacts
