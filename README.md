# wum-tracker

Konwerter planu zajęć WUM (plik Excel z arkuszami „PLAN ZAJĘĆ” i „WYKŁADY”) do kalendarzy
`.ics` dla wybranej grupy dziekańskiej. Arkusz „Arkusz1” jest pomocniczy i jest pomijany.

## Użycie

```bash
pip install -r requirements.txt
python wum_tracker.py                                  # domyślnie: grupa 8, podgrupa 8b
python wum_tracker.py --group "grupa 14" --subgroup a  # tylko zajęcia podgrupy 14a
python wum_tracker.py --input data/plan.xlsx --output plan_zajec.ics --calendars-dir kalendarze
python wum_tracker.py --no-lectures                    # bez arkusza z wykładami
```

Bez `--input` skrypt bierze najnowszy plik `licencjat*.xlsx` (także z katalogu `data/`).
Parametry można też ustawić zmiennymi środowiskowymi (np. w GitHub Actions):
`WUM_INPUT_FILE`, `WUM_TARGET_GROUP`, `WUM_TARGET_SUBGROUP`, `WUM_OUTPUT_FILE`, `WUM_CALENDARS_DIR`.

Pliki wynikowe:

| Plik | Kolor | Zawartość |
|---|---|---|
| `kalendarze/cwiczenia.ics` | czerwony | ćwiczenia ze wszystkich przedmiotów, zajęcia praktyczne, OSCE |
| `kalendarze/seminaria.ics` | zielony | seminaria ze wszystkich przedmiotów |
| `kalendarze/wyklady.ics` | niebieski | wykłady stacjonarne i online (💻) |
| `kalendarze/zaliczenia.ics` | fioletowy | zaliczenia, kolokwia, egzaminy (gdy pojawią się w planie) |
| `plan_zajec.ics` | – | wszystko razem w jednym kalendarzu |

Subskrybuj **albo** cztery osobne kalendarze, **albo** `plan_zajec.ics`, inaczej wydarzenia się zdublują.

Kody wyjścia: `0` – OK, `1` – nie znaleziono zajęć (plik `.ics` nie jest nadpisywany),
`2` – błąd krytyczny.

## Założenia

- Zakresy dat (`05.10. - 25.01.`) są rozwijane co 7 dni z pominięciem świąt i przerwy
  świątecznej 21.12.2026–03.01.2027 (`NO_CLASS_DATES`, `WINTER_BREAK` w skrypcie).
- „cały semestr” = każdy dany dzień tygodnia między `SEMESTER_START` a `SEMESTER_END`.
- Sala dopisana w komórce grupy ma pierwszeństwo przed informacją z wiersza 6.
- Przedmiot bez słowa „ćwiczenia/seminaria” w nazwie (np. „Język angielski”, „Anatomia NZZA”)
  jest traktowany jako ćwiczenia, zgodnie z legendą pod planem.
- Wykład bez podanej formy jest traktowany jako stacjonarny (z notatką w opisie).

## Wygląd wydarzeń

Tytuł ma wzór `emoji FORMA · Przedmiot: temat`, bo w widoku tygodnia w telefonie widać
ok. 20 znaków, np. `🩺 ĆW · PP`, `🦴 ĆW · Anatomia`, `🩺 WYK · PP 💻` (💻 = online).
Zaliczenia mają na początku `⚠️` (zaliczenie, kolokwium, wejściówka) albo `🎯` (egzamin, OSCE).

Opis zawsze zaczyna się od tego samego szablonu (puste pozycje są pomijane):

```
👥 Grupa: 8b
🎒 Przynieś: strój, identyfikator, spięte włosy, krótkie paznokcie
🔗 Teams / e-learning: link
```

Niżej: pełna nazwa z planu, prowadzący, liczba spotkań, terminy i źródło w arkuszu.
W polu „Miejsce” jest adres z salą, np. `Ciołka 27, sala 204, Warszawa`, więc mapa i czas
dojazdu działają same. Linki do Teams (w planie ich nie ma) wpisz w `LINKS` w skrypcie;
trafią do pola URL i do opisu.

Do ćwiczeń i zaliczeń dołączone są dwa przypomnienia: dzień wcześniej o 18:00
(z listą rzeczy do zabrania) i 60 minut przed zajęciami. Wydarzenia z adresem mają
włączone automatyczne liczenie czasu dojazdu w Kalendarzu Apple.

Emoji, kolory, kategorie, przypomnienia i listy „Przynieś” ustawia się w sekcji
„WYGLĄD KALENDARZA” na początku skryptu.

## GitHub Actions

Workflow `.github/workflows/generate-calendar.yml` generuje `plan_zajec.ics` oraz
`kalendarze/*.ics` i commituje je do gałęzi `main`, gdy:
- na `main` trafi zmiana w `data/` (np. nowy plik z planem), skrypcie lub zależnościach,
- uruchomisz go ręcznie (Actions → „Generuj kalendarz” → *Run workflow*),
- codziennie rano, **tylko** jeśli ustawiono zmienną `WUM_PLAN_URL`.

Commit powstaje tylko wtedy, gdy plan faktycznie się zmienił. Pliki `.ics` są też
dostępne jako artefakt uruchomienia.

Zmienne repozytorium (Settings → Secrets and variables → Actions → *Variables*), wszystkie opcjonalne:

| Zmienna | Przykład | Znaczenie |
|---|---|---|
| `WUM_TARGET_GROUP` | `grupa 8` | grupa dziekańska |
| `WUM_TARGET_SUBGROUP` | `b` | podgrupa |
| `WUM_PLAN_URL` | `https://…/plan.xlsx` | pobieraj plan z tego adresu zamiast z `data/` |

Aktualizacja planu: wrzuć nowy plik do `data/` (np. `licencjat-i-rok-piel.-15.10.2026.xlsx`).
Skrypt wybiera plik z najnowszą datą w nazwie.

### Subskrypcja kalendarzy

Adresy (działają tylko dla **publicznego** repozytorium):

```
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/kalendarze/cwiczenia.ics
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/kalendarze/seminaria.ics
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/kalendarze/wyklady.ics
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/kalendarze/zaliczenia.ics
https://raw.githubusercontent.com/michcio-afk/wum-tracker/main/plan_zajec.ics   # albo wszystko razem
```

**Apple (iPhone / Mac):** Ustawienia → Kalendarz → Konta → Dodaj konto → Inne →
Dodaj subskrybowany kalendarz (na Macu: Plik → Nowa subskrypcja kalendarza). Kolor zwykle
ustawia się sam; jeśli nie, wybierz go ręcznie. Na Macu **odznacz „Usuń: Alerty”**,
inaczej przypomnienia z pliku zostaną pominięte.

**Google:** Inne kalendarze → Z adresu URL, każdy adres osobno. Google nie czyta kolorów
ani przypomnień z subskrybowanych plików: kolor każdego kalendarza ustaw ręcznie
(⋮ przy nazwie kalendarza), a przypomnienia w jego ustawieniach. Subskrypcje odświeżają się
co kilka–kilkanaście godzin.
