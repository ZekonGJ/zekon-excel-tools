# Historia wersji Narzędzi ZEKON

Historia projektu, prowadzona na życzenie użytkownika od 2026-10-02.
Opisy wcześniejszych wydań odtworzono z dostępnego kodu i przebiegu pracy.
Nie stanowią deklaracji, że wszystkie wydania były poprawne lub przetestowane.

## Zasady kolejnych aktualizacji

- Przy każdej zmianie dostarczanego dodatku nadaj nowy numer wersji.
- Zachowaj zgodność numeru w APP_VERSION, tytule okna, lewym panelu,
  instalatorach, katalogu instalacji i dokumentacji.
- Dopisz tutaj cel zmiany, zmienione zachowanie, zakres weryfikacji
  oraz znane ograniczenia. Nie przepisuj historii tak, jakby późniejsza
  poprawka istniała we wcześniejszym wydaniu.
- Wygeneruj ponownie instalatory, sumy kontrolne i pakiet.
- Użytkownik chce przechowywania opisów w projekcie, bez prezentowania
  historii w rozmowie. Pokaż ją dopiero na jego prośbę.
- Nie deklaruj testów Windows Excel, jeżeli nie zostały wykonane.

## 1.0.0-rc2 — 2026-10-02

- Rozbijanie: jeden zapis bloku kopii na wiersz źródłowy, zbiorcze
  ustawienie ilości i numeracji; odczyt ilości do tablicy.
- Zaokrąglanie: zapis ciągłych grup zmienianych komórek do 2048 naraz;
  pomijane formuły, ukryte wiersze i teksty nie są nadpisywane.
- Wyszukiwanie: tablice wartości obu kolumn. Eksport: jednorazowy
  odczyt zakresu mapowanych kolumn do pamięci.
- Panel wstrzymuje automatyczne przeliczanie podczas operacji i przywraca
  wcześniejszy tryb po zakończeniu lub błędzie. Sekwencja rozbijania
  i eksportu przelicza arkusz przed odczytem danych eksportowych.
- Pasek postępu, nazwa etapu, procent etapu, licznik oraz czas pracy;
  odświeżanie ograniczone do około 5 razy na sekundę.
- Przerwij i zamknięcie okna zgłaszają żądanie zatrzymania. Zatrzymanie
  następuje między porcjami pracy; wynik może pozostać częściowy.
- Kontrolki ustawień są blokowane podczas pracy, aby DoEvents nie
  umożliwiało zmiany parametrów lub powtórnego uruchomienia operacji.
- INSTALUJ.cmd / Instaluj-Zekon.ps1: wspólna ścieżka instalacji i aktualizacji,
  testy przed rejestracją, przełączanie z poprzednich wersji w katalogu ZEKON,
  próba przywrócenia poprzedniej rejestracji po błędzie.
- Pakiet źródłowy buduje XLAM lokalnie; następnie tworzy gotowy ZIP
  dystrybucyjny. Inne komputery nie wymagają importu BAS ani dostępu
  do projektu VBA. Makra i PowerShell nadal podlegają politykom firmy.
- Dodane testy formuł względnych i wysokości kopiowanych wierszy,
  granicy porcji zaokrągleń, przerwania i przywrócenia przeliczania.
- Weryfikacja tutaj: statyczna spójność kodu, kontrolek, generatorów,
  źródeł osadzonych i pakietu. Brak runtime Windows/Excel/PowerShell.
  Brak pomiarów czasów i potwierdzenia pełnych testów rc2 na Windows.
- Ograniczenia: wstawianie wierszy nadal odbywa się per wiersz źródłowy;
  pojedynczy zapis/otwarcie/przeliczanie w Excelu może blokować odświeżanie
  i obsługę Przerwij do końca tej czynności. Brak obietnicy przyspieszenia xN.

## 1.0.0-rc1f — 2026-10-02

- Poprawiono błędne oczekiwanie w teście zachowania zakresu filtra.
- Diagnostyka użytkownika (Excel 16.0, build 20430) wykazała, że Excel
  już przy tworzeniu filtra zmienia syntetyczny zakres A7:BD10 na B7:B10,
  ponieważ tylko kolumna B zawiera dane. SearchPositions zachował B7:B10.
- Potwierdzony wynik tego przypadku w rc1e: braki 0; ukryte wiersze
  8/9/10 = False/True/False, zgodnie z oczekiwaniem.
- Test rc1f porównuje zakres po wyszukiwaniu z rzeczywistym zakresem
  utworzonym przez Excel przed wyszukiwaniem. Dodatkowo sprawdza wiersz
  początku 7 i liczbę wierszy 4. Nadal sprawdza wynik i widoczność wierszy.
- Komunikat niezgodności zawiera teraz oczekiwany i otrzymany zakres.
- Nie zmieniono logiki makr użytkowych; wersja i nazwy instalacyjne rc1f.
- Wykonano weryfikację statyczną i zgodność kodu osadzonego w instalatorze.
  Pełne testy rc1f w Windows Excel oczekują wykonania podczas instalacji.

## 1.0.0-rc1e — 2026-10-02

- Naprawa instalatora po zgłoszeniu błędu 1004 na etapie SaveAs XLAM.
- Potwierdzony błąd implementacji: enumeracja Workbooks nie obejmuje
  otwartych dodatków, więc dotychczasowa próba zamknięcia starego dodatku
  nie gwarantowała usunięcia konfliktu nazw. Sam zrzut nie potwierdza,
  że był to jedyny powód błędu zapisu u użytkownika.
- Plik wynikowy ma teraz nazwę ZekonTools_1.0.0-rc1e.xlam; następne
  wydania także mają mieć własną nazwę pliku z numerem wersji.
- Usunięto próbę zamykania starego dodatku przed budowaniem. Pozostaje
  dostępny; użytkownik przełącza wersje po pomyślnym zbudowaniu i testach.
- Instalator naprawczy ma unikalną nazwę modułu ZekonInstallerRC1E
  i makra AktualizujZekonRC1E, aby nie mylić go z poprzednimi importami.
- Tytuł w ustawieniach dodatków zawiera numer wersji. Komunikat błędu
  nie sugeruje już automatycznie problemu z uprawnieniami do VBA.
- Logika makr pozostaje bez zmian względem rc1d.
- Weryfikacja: statyczna kontrola generatorów, nazwy pliku, osadzonych
  źródeł i panelu. Nie wykonano testów Windows Excel w tym środowisku.
- Źródło: https://learn.microsoft.com/en-us/office/vba/api/excel.application.workbooks

## 1.0.0-rc1d — 2026-10-02

- AUTO dla początku danych we wszystkich funkcjach: zaokrąglanie,
  wyszukiwanie, rozbijanie, oba eksporty i sekwencje rozbijania z eksportem.
- Osobne wykrywanie początku listy źródłowej i zestawienia docelowego.
- Rozpoznawanie na podstawie tabeli lub filtra obejmującego kolumnę,
  znanych nagłówków w pierwszych 50 niepustych komórkach, a następnie
  prostych list bez nagłówka. Ręczny numer pozostaje opcją.
- Niejednoznaczny układ zatrzymuje automatyczne rozpoznawanie.
- Migracja profili: stare zapisane numery początku danych nie zastępują AUTO.
- Wiersz zapisu nowego eksportu jest ustalany z formatu: malowanie 3,
  ocynk 2; pole to jest ukryte w panelu.
- Obsługa istniejącego filtra z pustym wierszem nagłówka oraz nagłówka
  oddzielonego pustymi wierszami od danych. W przesłanym BT-List filtr
  A7:BD14047 daje początek 8; numer 8 nie jest stałą algorytmu.
- Wstępna walidacja wyszukiwania przed utworzeniem kopii roboczej.
- Po zaokrąglaniu i rozbijaniu aktywowany jest przetwarzany arkusz.
- Instalator naprawczy zawiera pełny aktualny kod, niezależny od kodu
  w starym rozpakowanym folderze. Nadal korzysta z logo z tego folderu.
- Numer wersji widoczny w tytule okna i lewym panelu.
- Weryfikacja wykonana: kontrola statyczna pakietu, zgodność osadzonych
  źródeł instalatora i analiza struktury dostarczonego XLSX.
- Dodano testy natywne początków 1, 4, 5, 8 i 12, filtra, nagłówków,
  tabeli i niejednoznaczności. Instalator uruchamia je w Excelu użytkownika.
- Brak potwierdzenia wykonania testów rc1d w Windows Excel na dzień wpisu.
- Ograniczenia: heurystyka nie rozumie dowolnego układu danych;
  rozpoznawanie początku nie zmienia reguł końca danych ani ograniczeń
  obsługi tabel poszczególnych modułów.

## 1.0.0-rc1c — 2026-10-02

- Poprawiono zakres AutoFilter i numer pola filtra względem początku zakresu.
- Jedno kryterium przekazywane skalarnie, wiele jako tablica tekstowa;
  obsłużono także brak dopasowań.
- Oznaczanie braków formatowaniem warunkowym bez funkcji zależnych
  od języka Excela; usuwanie własnych nieaktualnych oznaczeń.
- Dodano testy jednego, wielu i zera dopasowań.
- Dostarczony przez użytkownika test-result.txt z 2026-10-02 11:56:02
  zawiera PASS dla testów tej wersji. Kod dostarczonego XLAM porównano
  ze źródłami. Były to testy na danych syntetycznych i inicjalizacja panelu,
  nie pełna weryfikacja pracy na rzeczywistym skoroszycie.
- Pozostały problemy ujawnione później: wymóg ręcznego początku danych,
  pusty wiersz nagłówka filtra i tworzenie kopii przed częścią walidacji.

## 1.0.0-rc1b — 2026-10-02

- Poprawiono budowanie XLAM po SaveAs: sprawdzenie istnienia pliku,
  zamknięcie obiektu źródłowego, ponowne otwarcie zapisanego dodatku,
  kontrola FullName i uruchomienie testów z właściwego pliku.
- Usunięto założenie, że zapis XLAM zmienia nazwę i tożsamość aktualnego
  obiektu skoroszytu w oczekiwany sposób.
- Potwierdzono statycznie kolejność operacji. Brak osobnego potwierdzenia
  testów Windows Excel tego wydania w dostępnej dokumentacji.

## 1.0.0-rc1a — 2026-10-02

- Korekty wywołań makr przez pełną ścieżkę dodatku oraz zgodności
  generatora BAS i kontrolek panelu.
- Weryfikacja statyczna; brak potwierdzenia testów Windows Excel tego
  wydania w dostępnej dokumentacji.

## 1.0.0-rc1 — 2026-10-02

- Pierwszy pakiet modułowego dodatku: wspólne funkcje, zaokrąglanie,
  wyszukiwanie i raport braków, rozbijanie, eksport malowania i ocynku.
- Panel z logo ZEKON, profile ustawień, praca na kopii i wybór widocznych
  wierszy; źródła VBA i narzędzia budowania XLAM w lokalnym Excelu.
- Wydanie początkowe wymagało kolejnych poprawek instalacji i działania.

## Infrastruktura instalatora EXE — 2026-10-05, wersja robocza

Przeniesiono źródła rc2 bez zmiany ich logiki. Dodano repozytorium, natywny
instalator Windows, planowanie rejestracji z zachowaniem obcych wpisów,
kontrolę integralności, wersjonowane katalogi, rollback kolejnych instalacji,
skrót aktualizacji do GitHub Releases i konfigurację Actions.
Brak podpisów i aktualnego binarnego XLAM blokuje gotowe wydanie produkcyjne.
Nie potwierdzono jeszcze kompilacji Windows ani instalacji w Excelu.
