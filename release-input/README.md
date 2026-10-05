Tutaj umieść `ZekonTools.xlam` i `test-result.txt` z tego samego, sprawdzonego
budowania aktualnej wersji w lokalnym Excelu. Nie dodawaj plików użytkowników.
Nie zastępuj brakujących plików starym XLAM z rc1c.

Ten katalog celowo nie zawiera binarnego dodatku. Dostępny wcześniej plik
użytkownika pochodził z rc1c, a aktualne źródła są rc2. Pipeline blokuje
utworzenie kompletnego instalatora bez odpowiedniego artefaktu i raportu PASS.

Kontrole wersji i raportu zapobiegają pomyłkom; nie dowodzą zgodności binarnego
VBA ze źródłami. Osoba wydająca potwierdza pochodzenie obu plików z tego samego
budowania i zatwierdza wynik na rzeczywistym skoroszycie.
