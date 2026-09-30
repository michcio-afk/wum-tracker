# wum-tracker

Konwerter planu zajęć WUM (plik Excel z arkuszem „PLAN ZAJĘĆ”) do kalendarza `.ics`
dla wybranej grupy dziekańskiej.

## Użycie

```bash
pip install -r requirements.txt
python wum_tracker.py                                  # grupa 14, wszystkie podgrupy
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
