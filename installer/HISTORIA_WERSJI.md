# Instalator 1.3.0 — 08.10.2026

- Aktualizacja z paczki .zekonupdate obok uruchomionego EXE, bez zależności od litery dysku lub katalogu roboczego.
- Automatyczne rozpoczęcie po otwarciu; Excel musi być zamknięty przez użytkownika.
- Wybór najwyższego ReleaseNumber z manifestu; odrzucenie konfliktu numerów i uszkodzonych paczek. Bez automatycznego pobierania online.
- Odczyt całej paczki do pamięci lokalnej przed instalacją; sprawdzanie sum plików i zachowanie poprzedniej wersji.
- Istniejące rejestrowanie dodatku i możliwość wycofania zachowane. Dodatek R05 bez zmian.

Administrator: umieść ZekonSetup.exe i paczkę .zekonupdate w tym samym folderze. Nową paczkę kopiuj najpierw z rozszerzeniem .part i zmień nazwę dopiero po zakończeniu kopiowania. Użytkownik uruchamia EXE z udziału sieciowego; nie kopiuje samego EXE na pulpit. Skrót powinien wskazywać EXE na udziale.

Testy automatyczne Windows obejmują wybór wersji, brak paczki, duplikaty, zablokowany plik, uszkodzenia i pracę na pełnej lokalnej kopii. Rzeczywisty udział SMB oraz uruchomienie okna na stanowisku użytkownika wymagają sprawdzenia w firmie. Instalator i VBA pozostają niepodpisane cyfrowo.
