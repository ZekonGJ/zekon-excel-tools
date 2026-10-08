# Pakiety produkcyjne — zapis postępu

Baza: `1b135f6ccfa46ff21e57437f11eb519d5571a71e`, opublikowana wersja 1.0.0-rc2e, aktualizator ReleaseNumber 7.

## 2026-10-08: odzyskanie po awarii

Środowisko wykonawcze ponownie działa. Katalog roboczy tej rozmowy został jednak wyczyszczony. Wczorajsze niezatwierdzone moduły nie znalazły się w zdalnym repozytorium. Ich kod jest odtwarzany z zachowanego zapisu rozmowy, bez zmiany uzgodnionych zasad. Zachowały się oryginalne materiały wejściowe. Osobny katalog poprzedniego projektu zawiera niezwiązane z tym etapem zmiany i pozostaje nietknięty.

Potwierdzone teraz: czysta baza repozytorium; 9 testów Python PASS; opublikowany kanał rc2e PASS.

Nie wykonano: kompilacji nowych modułów VBA, uruchomienia nowych makr, testów Excel 2016, zbudowania nowego XLAM, wygenerowania i odebrania nowych pakietów produkcyjnych. Wyników zapisanych w źródłowych XLSX nie traktujemy jako przeliczenia nowych formuł.

Etapy będą zatwierdzane na osobnej gałęzi rozwojowej. Nie zmieniać `updates/latest.json`, opublikowanych paczek ani instalatora produkcyjnego bez natywnej walidacji.

## Ustalenia obowiązujące

- Norma w `Czasy` dotyczy ilości całego wiersza.
- Przy PCS = 10 przyrosty 0,35 + 0,65 oznaczają jedną sztukę wykonaną i dziewięć pozostałych. Przydział wykonanej normy: norma wiersza / PCS × przyrost.
- Składanie: jawny współczynnik każdego bloku; zachować korekty 0,6 oraz 0,7.
- Spawanie: norma główna + wsporniki; dodatkowe pozostają osobno. Zachować obecny układ listy części.
- Malowanie: ręczne przyrosty zmianowe zastępują odziedziczone powiązania RF; zachować historyczne dane.
- Potwierdzony przepływ między plikami: liczba części z normowania składania do normowania malowania. Korekt ręcznych nie nadpisywać.
- Dane produkcyjne i oryginalne skoroszyty pozostają poza publicznym repozytorium.

## Pozostałe prace

Odtworzenie i przegląd modułów; kontrola kompletności zakładek/formularzy; generowanie i aktualizacja wszystkich czterech plików; synchronizacja; testy przerwanego zapisu; przykłady i wygląd; dokumentacja; natywne bramki Windows/Excel 2016. Sam brak Excela nie blokuje powyższych prac źródłowych i analiz danych.
