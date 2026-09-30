# wum-tracker

Konwerter planu zajęć WUM (plik Excel z arkuszem „PLAN ZAJĘĆ”) do kalendarza `.ics`
dla wybranej grupy dziekańskiej.

## Użycie

```bash
pip install -r requirements.txt
python wum_tracker.py                                  # domyślnie: grupa 8, podgrupa 8b
python wum_tracker.py --group "grupa 14" --subgroup a  # tylko zajęcia podgrupy 14a
python wum_tracker.py --input data/plan.xlsx --output plan_zajec.ics
```

Bez `--input` skrypt bierze najnowszy plik `licencjat*.xlsx` (także z katalogu `data/`).
Parametry można też ustawić zmiennymi środowiskowymi (np. w GitHub Actions):
`WUM_INPUT_FILE`, `WUM_TARGET_GROUP`, `WUM_TARGET_SUBGROUP`, `WUM_OUTPUT_FILE`.

Kody wyjścia: `0` – OK, `1` – nie znaleziono zajęć (plik `.ics` nie jest nadpisywany),
`2` – błąd krytyczny.

## Założenia

- Zakresy dat (`05.10. - 25.01.`) są rozwijane co 7 dni z pominięciem świąt i przerwy
  świątecznej 21.12.2026–03.01.2027 (`NO_CLASS_DATES`, `WINTER_BREAK` w skrypcie).
- „cały semestr” = każdy dany dzień tygodnia między `SEMESTER_START` a `SEMESTER_END`.
- Sala dopisana w komórce grupy ma pierwszeństwo przed informacją z wiersza 6.

## GitHub Actions

Workflow `.github/workflows/generate-calendar.yml` generuje `plan_zajec.ics` i commituje go
do gałęzi `main`, gdy:
- na `main` trafi zmiana w `data/` (np. nowy plik z planem), skrypcie lub zależnościach,
- uruchomisz go ręcznie (Actions → „Generuj kalendarz” → *Run workflow*),
- codziennie rano, **tylko** jeśli ustawiono zmienną `WUM_PLAN_URL`.

Commit powstaje tylko wtedy, gdy plan faktycznie się zmienił. Plik `.ics` jest też
dostępny jako artefakt uruchomienia.

Zmienne repozytorium (Settings → Secrets and variables → Actions → *Variables*), wszystkie opcjonalne:

| Zmienna | Przykład | Znaczenie |
|---|---|---|
| `WUM_TARGET_GROUP` | `grupa 8` | grupa dziekańska |
| `WUM_TARGET_SUBGROUP` | `b` | podgrupa |
| `WUM_PLAN_URL` | `https://…/plan.xlsx` | pobieraj plan z tego adresu zamiast z `data/` |

Aktualizacja planu: wrzuć nowy plik do `data/` (np. `licencjat-i-rok-piel.-15.10.2026.xlsx`).
Skrypt wybiera plik z najnowszą datą w nazwie.

### Subskrypcja kalendarza

Dodaj w Kalendarzu Google (*Inne kalendarze → Z adresu URL*) lub w Apple/Outlook:

```
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/plan_zajec.ics
```

Ten adres działa tylko dla **publicznego** repozytorium. Kalendarz Google odświeża
subskrypcje co kilka–kilkanaście godzin.
