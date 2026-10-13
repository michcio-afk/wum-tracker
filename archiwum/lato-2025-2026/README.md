# Semestr letni 2025/2026 – grupa 8, podgrupy b

Jednorazowo wygenerowane z pliku
`pielegniarstwo-licencjat-i-rok-lato-2025-2026-28.05.2026.xlsx` (plan z 28.05.2026):

```bash
python wum_tracker.py \
  -i archiwum/lato-2025-2026/pielegniarstwo-licencjat-i-rok-lato-2025-2026-28.05.2026.xlsx \
  -g "grupa 8" -s b --export-bring \
  -o archiwum/lato-2025-2026/grupa-08_b/wszystko.ics --calendars-dir archiwum/lato-2025-2026/grupa-08_b
```

| Plik | Zawartość |
|---|---|
| `grupa-08_b/wszystko.ics` | wszystkie zajęcia (79 wydarzeń, 16.02–9.06.2026) |
| `grupa-08_b/cwiczenia.ics` | ćwiczenia i zajęcia praktyczne (70) – czerwony |
| `grupa-08_b/seminaria.ics` | seminaria (9) – zielony |

Zaimportuj **albo** `wszystko.ics`, **albo** dwa osobne pliki.

Uwagi:

- Wykładów w tym planie nie ma (arkusz „WYKŁADY”: wszystkie wykłady na platformie e-learningowej).
- Pominięte przy rozwijaniu zakresów dat: przerwy 2–8.04. i 27.04.–1.05.2026 oraz święta.
- Zajęcia praktyczne z PP i ćwiczenia z PP nie mają sali w planie (adresy placówek zna
  Zakład Propedeutyki Pielęgniarstwa) – w opisie jest notatka z legendy.
- Interna: poniedziałek 13.04. w CSM, pozostałe dni w Klinice I (Nowogrodzka 59) – wg legendy planu.
- Grafik prowadzących PP z `data/prowadzacy-pp.csv` dotyczy semestru zimowego, więc tu go nie użyto.
- Ten folder nie jest aktualizowany przez workflow.
