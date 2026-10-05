# ZEKON — zasady utrzymania

- Zachowuj istniejące moduły VBA i nie zmieniaj zachowania bez związku z zadaniem.
- Każda dostarczana zmiana dodatku dostaje nowy numer; opisy w HISTORIA_WERSJI.md.
  Użytkownik nie chce całej historii w rozmowie, chyba że o nią poprosi.
- Nie deklaruj testów Windows/Excel na podstawie statycznego sprawdzenia kodu.
- Nie publikuj instalatora z nieaktualnym lub syntetycznym XLAM.
- Testy muszą pozostać aktywne; błędne oczekiwanie można poprawić na podstawie dowodu.
- Instalacja na stanowisku użytkownika nie może wymagać PowerShella, VBA buildera
  ani zmiany zabezpieczeń. Nie dodawaj tokenów GitHub do instalatora.
- Nie zamykaj wymuszenie Excela, nie usuwaj skoroszytów użytkownika i obcych dodatków.
- Nie dodawaj certyfikatów prywatnych, tokenów ani danych produkcyjnych do repozytorium.
- Do zmian logiki rejestracji/pakowania dopisz test odpowiedniego ryzyka.
- Przed produkcyjnym wydaniem wymagane są próby Windows/Excel i jawny status podpisu.
