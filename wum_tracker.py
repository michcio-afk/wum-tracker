#!/usr/bin/env python3
"""
WUM-Tracker - konwerter planu zajęć WUM (Excel) do kalendarza iCalendar (.ics).

Plik Excel ma bardzo nieregularną strukturę (arkusz "PLAN ZAJĘĆ"):
  * kolumna A (od wiersza 7)  - 15-minutowe sloty czasowe, np. "8.15 - 8.30",
  * wiersz 3                  - dni tygodnia (komórki scalone poziomo),
  * wiersz 4                  - nazwa przedmiotu (zwykle scalona poziomo),
  * wiersz 5                  - daty zajęć ("05.10. - 25.01.", "12.10., 19.10.",
                                "06.10. i 13.10.", "cały semestr", ...),
  * wiersz 6                  - sala / notatki,
  * siatka od C7              - nazwy grup w komórkach scalonych pionowo
                                (wysokość scalenia = czas trwania zajęć).

Ponieważ niemal wszystko opiera się na scalonych komórkach, używamy openpyxl
(data_only=True) i sami rozwiązujemy wartości scaleń - pandas tego nie potrafi.

Użycie:
    python wum_tracker.py [--input PLIK.xlsx] [--group "grupa 14"]
                          [--subgroup a] [--output plan_zajec.ics]

Każdy parametr można też ustawić zmienną środowiskową (wygodne w GitHub
Actions): WUM_INPUT_FILE, WUM_TARGET_GROUP, WUM_TARGET_SUBGROUP, WUM_OUTPUT_FILE.
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

import openpyxl
import pytz
from icalendar import Calendar, Event
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# KONFIGURACJA
# ---------------------------------------------------------------------------

# Grupa dziekańska, dla której generujemy kalendarz.
TARGET_GROUP = "grupa 14"

# Podgrupa ("a", "b", "c") - ćwiczenia odbywają się w podgrupach 8/12-osobowych.
# None = weź zajęcia całej grupy ORAZ wszystkich jej podgrup (podgrupa trafi
# do tytułu wydarzenia, np. "Anatomia NZZA [14a]").
TARGET_SUBGROUP: str | None = None

# Domyślny plik wejściowy; gdy nie istnieje, bierzemy najnowszy pasujący *.xlsx.
INPUT_FILE = "licencjat-i-rok-piel.xlsx"
INPUT_GLOB_PATTERNS = ("licencjat*.xlsx", "data/licencjat*.xlsx", "*.xlsx", "data/*.xlsx")

OUTPUT_FILE = "plan_zajec.ics"
SHEET_NAME = "PLAN ZAJĘĆ"

# Układ arkusza (indeksy 1-based, jak w Excelu).
ROW_DAYS = 3
ROW_SUBJECT = 4
ROW_DATES = 5
ROW_ROOM = 6
FIRST_GRID_ROW = 7
FIRST_GRID_COL = 3
TIME_COL = 1

TIMEZONE = pytz.timezone("Europe/Warsaw")

# Rok akademicki: miesiące >= 8 należą do roku startowego, pozostałe do kolejnego
# (10, 11, 12 -> 2026; 01, 02 -> 2027).
ACADEMIC_YEAR_START = 2026
FIRST_MONTH_OF_ACADEMIC_YEAR = 8

# Granice semestru - używane dla wpisu "cały semestr".
SEMESTER_START = date(2026, 10, 1)
SEMESTER_END = date(2027, 1, 31)

# Dni bez zajęć, pomijane przy rozwijaniu ZAKRESÓW dat ("05.10. - 25.01.").
# Daty wpisane w Excelu jawnie (pojedynczo) nie są filtrowane.
# Przerwa świąteczna ustalona na podstawie list dat w samym planie
# (np. poniedziałki: 14.12. -> 04.01.) - zweryfikuj z kalendarzem akademickim.
NO_CLASS_DATES: set[date] = {
    date(2026, 11, 1),   # Wszystkich Świętych
    date(2026, 11, 11),  # Narodowe Święto Niepodległości
    date(2026, 12, 24),  # Wigilia
    date(2026, 12, 25),  # Boże Narodzenie
    date(2026, 12, 26),  # Drugi dzień Bożego Narodzenia
    date(2027, 1, 1),    # Nowy Rok
    date(2027, 1, 6),    # Trzech Króli
}
WINTER_BREAK = (date(2026, 12, 21), date(2027, 1, 3))

WEEKDAYS_PL = {
    "PONIEDZIAŁEK": 0,
    "WTOREK": 1,
    "ŚRODA": 2,
    "CZWARTEK": 3,
    "PIĄTEK": 4,
    "SOBOTA": 5,
    "NIEDZIELA": 6,
}

log = logging.getLogger("wum-tracker")


# ---------------------------------------------------------------------------
# NARZĘDZIA TEKSTOWE
# ---------------------------------------------------------------------------

def normalize_ws(text: object) -> str:
    """Zamienia dowolne białe znaki (\\n, \\xa0, tabulatory, wielokrotne spacje)
    na pojedynczą spację i obcina brzegi. None -> ''."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def split_lines(text: object) -> list[str]:
    """Dzieli tekst komórki na niepuste, oczyszczone linie."""
    if text is None:
        return []
    return [normalize_ws(line) for line in str(text).splitlines() if normalize_ws(line)]


def cell_ref(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def is_no_class_day(d: date) -> bool:
    return d in NO_CLASS_DATES or WINTER_BREAK[0] <= d <= WINTER_BREAK[1]


# ---------------------------------------------------------------------------
# DOSTĘP DO KOMÓREK SCALONYCH
# ---------------------------------------------------------------------------

class MergedSheet:
    """Opakowanie arkusza, które rozumie komórki scalone.

    W openpyxl wartość scalonego obszaru siedzi tylko w lewej górnej komórce,
    pozostałe są puste. Budujemy indeks (wiersz, kolumna) -> zakres scalenia,
    żeby w O(1) odczytać wartość dowolnej komórki "tak jak widzi ją człowiek".
    """

    def __init__(self, ws: Worksheet):
        self.ws = ws
        self._index: dict[tuple[int, int], object] = {}
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._index[(r, c)] = rng
        log.info("Zindeksowano %d scalonych zakresów.", len(ws.merged_cells.ranges))

    def merged_range(self, row: int, col: int):
        """Zwraca zakres scalenia zawierający komórkę albo None."""
        return self._index.get((row, col))

    def bounds(self, row: int, col: int) -> tuple[int, int, int, int]:
        """(min_row, max_row, min_col, max_col) - dla zwykłej komórki to ona sama."""
        rng = self.merged_range(row, col)
        if rng is None:
            return row, row, col, col
        return rng.min_row, rng.max_row, rng.min_col, rng.max_col

    def value(self, row: int, col: int):
        """Wartość komórki z uwzględnieniem scaleń."""
        rng = self.merged_range(row, col)
        if rng is not None:
            return self.ws.cell(rng.min_row, rng.min_col).value
        return self.ws.cell(row, col).value

    def value_in_block(self, row: int, col: int, block_min_col: int):
        """Jak value(), ale gdy komórka jest pusta, szuka najbliższej niepustej
        wartości na lewo - nie dalej niż do początku bloku przedmiotu
        (block_min_col). Zwraca (wartość, kolumna_źródłowa)."""
        for c in range(col, block_min_col - 1, -1):
            v = self.value(row, c)
            if normalize_ws(v):
                return v, c
        return None, None


# ---------------------------------------------------------------------------
# PARSOWANIE CZASU
# ---------------------------------------------------------------------------

_TIME_SLOT_RE = re.compile(r"(\d{1,2})[.:](\d{2})\s*[-–—]\s*(\d{1,2})[.:](\d{2})")


def parse_time_slot(text: object) -> tuple[str, str] | None:
    """'8.15 - 8.30' -> ('08:15', '08:30'). Zwraca None, gdy to nie jest slot."""
    m = _TIME_SLOT_RE.search(normalize_ws(text))
    if not m:
        return None
    h1, m1, h2, m2 = (int(x) for x in m.groups())
    if not (0 <= h1 < 24 and 0 <= h2 < 24 and 0 <= m1 < 60 and 0 <= m2 < 60):
        return None
    return f"{h1:02d}:{m1:02d}", f"{h2:02d}:{m2:02d}"


def build_time_map(ws: Worksheet) -> dict[int, tuple[str, str]]:
    """Mapuje numer wiersza siatki na (początek, koniec) slotu z kolumny A.
    Wiersze bez poprawnego slotu (notatki pod planem) są pomijane."""
    time_map: dict[int, tuple[str, str]] = {}
    for row in range(FIRST_GRID_ROW, ws.max_row + 1):
        slot = parse_time_slot(ws.cell(row, TIME_COL).value)
        if slot:
            time_map[row] = slot
    if time_map:
        first, last = min(time_map), max(time_map)
        log.info(
            "Sloty czasowe: wiersze %d-%d (%s - %s), %d slotów.",
            first, last, time_map[first][0], time_map[last][1], len(time_map),
        )
    return time_map


# ---------------------------------------------------------------------------
# PARSOWANIE DAT (wiersz 5)
# ---------------------------------------------------------------------------

# "05.10. - 25.01.", "13.10 - 15.12.", "17.11. -15.12." itp.
_DATE_RANGE_RE = re.compile(
    r"(?<!\d)(\d{1,2})\.(\d{1,2})\.?\s*[-–—]\s*(\d{1,2})\.(\d{1,2})\.?(?!\d)"
)
# Pojedyncza data "12.10." lub "05.10" (bez końcowej kropki).
_SINGLE_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.?(?!\d)")
_WHOLE_SEMESTER_RE = re.compile(r"ca[łl]y\s+semestr", re.IGNORECASE)


def make_date(day: int, month: int) -> date | None:
    """Tworzy datę z uwzględnieniem przełomu roku akademickiego."""
    year = ACADEMIC_YEAR_START if month >= FIRST_MONTH_OF_ACADEMIC_YEAR else ACADEMIC_YEAR_START + 1
    try:
        return date(year, month, day)
    except ValueError:
        log.warning("Nieprawidłowa data: %02d.%02d - pomijam.", day, month)
        return None


def weekly_dates(start: date, end: date, context: str) -> list[date]:
    """Daty co 7 dni od start do end włącznie, z pominięciem dni wolnych."""
    result: list[date] = []
    current = start
    while current <= end:
        if is_no_class_day(current):
            log.info("    %s: pomijam %s (dzień wolny / przerwa).", context, current.strftime("%d.%m.%Y"))
        else:
            result.append(current)
        current += timedelta(days=7)
    return result


def parse_dates(raw: object, weekday: int | None, context: str = "") -> list[date]:
    """Parsuje zawartość wiersza 5 na listę konkretnych dat.

    Obsługiwane formaty (także w dowolnej kombinacji):
      * zakres "DD.MM. - DD.MM."  -> daty co 7 dni,
      * lista "DD.MM., DD.MM." / "DD.MM. i DD.MM." / pojedyncza data,
      * "cały semestr"           -> każdy `weekday` między SEMESTER_START a SEMESTER_END.
    """
    text = normalize_ws(raw)
    if not text:
        return []

    found: set[date] = set()

    if _WHOLE_SEMESTER_RE.search(text):
        if weekday is None:
            log.warning("%s: 'cały semestr', ale nie znam dnia tygodnia - pomijam.", context)
        else:
            first = SEMESTER_START + timedelta(days=(weekday - SEMESTER_START.weekday()) % 7)
            found.update(weekly_dates(first, SEMESTER_END, context))
        text = _WHOLE_SEMESTER_RE.sub(" ", text)

    # Najpierw zakresy (i wycinamy je z tekstu), potem pozostałe pojedyncze daty.
    for d1, m1, d2, m2 in _DATE_RANGE_RE.findall(text):
        start, end = make_date(int(d1), int(m1)), make_date(int(d2), int(m2))
        if start is None or end is None:
            continue
        if start > end:
            log.warning("%s: odwrócony zakres dat '%s' - zamieniam kolejność.", context, text)
            start, end = end, start
        found.update(weekly_dates(start, end, context))
    text = _DATE_RANGE_RE.sub(" ", text)

    for d, m in _SINGLE_DATE_RE.findall(text):
        parsed = make_date(int(d), int(m))
        if parsed:
            found.add(parsed)

    if not found:
        log.warning("%s: nie rozpoznano żadnej daty w '%s'.", context, normalize_ws(raw))

    if weekday is not None:
        for d in sorted(found):
            if d.weekday() != weekday:
                log.warning(
                    "%s: data %s wypada w inny dzień tygodnia niż kolumna - sprawdź plan.",
                    context, d.strftime("%d.%m.%Y (%A)"),
                )
    return sorted(found)


# ---------------------------------------------------------------------------
# DOPASOWANIE GRUPY
# ---------------------------------------------------------------------------

_GROUP_PREFIX_RE = re.compile(r"grupa\s*", re.IGNORECASE)
# "14", "14a", "14 a" - litera podgrupy opcjonalna.
_GROUP_TOKEN_RE = re.compile(r"(?<!\d)(\d+)\s*([a-z])?(?![\da-z])", re.IGNORECASE)


def parse_target(group: str, subgroup: str | None) -> tuple[str, str | None]:
    """'grupa 14' -> ('14', None); 'grupa 14a' -> ('14', 'a')."""
    m = _GROUP_TOKEN_RE.search(_GROUP_PREFIX_RE.sub("", normalize_ws(group)))
    if not m:
        raise ValueError(f"Nie rozpoznano numeru grupy w '{group}'")
    number, letter = m.group(1), (m.group(2) or "").lower() or None
    if subgroup:
        letter = subgroup.strip().lower()
    return number, letter


def match_group_cell(text: object, number: str, subgroup: str | None) -> list[str] | None:
    """Sprawdza, czy komórka siatki dotyczy docelowej grupy.

    Zwykłe `'grupa 14' in text` jest błędne: łapie też 'grupa 14a' (inna
    podgrupa), a 'grupa 1' łapie 'grupa 10'...'grupa 15'. Dlatego parsujemy
    PIERWSZĄ linię komórki na tokeny (kolejne linie to zwykle sala).

    Zwraca listę dopasowanych oznaczeń (np. ['14'] lub ['14a']) albo None.
    """
    lines = split_lines(text)
    if not lines or not _GROUP_PREFIX_RE.match(lines[0]):
        return None
    tokens = _GROUP_TOKEN_RE.findall(_GROUP_PREFIX_RE.sub("", lines[0]))
    matched: list[str] = []
    for num, letter in tokens:
        if num != number:
            continue
        letter = letter.lower()
        # Zajęcia całej grupy (bez litery) dotyczą każdej podgrupy.
        if not letter or subgroup is None or letter == subgroup:
            matched.append(f"{num}{letter}")
    return matched or None


# ---------------------------------------------------------------------------
# SALA / LOKALIZACJA (wiersz 6)
# ---------------------------------------------------------------------------

# Linie opisujące liczbę spotkań/godzin - to nie jest lokalizacja.
_NOT_LOCATION_RE = re.compile(r"spotka|godz|^i \d+ min|^\d+ min", re.IGNORECASE)
_GROUP_LIST_LINE_RE = re.compile(r"^grup[ay]?\b[:\s]*(.*)$", re.IGNORECASE)


def extract_location(raw: object, group_number: str) -> str:
    """Wyciąga z wiersza 6 samą lokalizację.

    Przykład: '15 spotkań po 4 godz.\\nPracownie do ćwiczeń,\\nul. Ciołka 27'
           -> 'Pracownie do ćwiczeń, ul. Ciołka 27'
    Obsługuje też sale przypisane do list grup, np.
      'grupy: 4, 12, 14, 10 - sala 104, ul. Ciołka' -> 'sala 104, ul. Ciołka'
    """
    lines = split_lines(raw)
    kept: list[str] = []
    group_specific: str | None = None
    for line in lines:
        gm = _GROUP_LIST_LINE_RE.match(line)
        if gm:
            head, _, room = gm.group(1).partition(" - ")
            numbers = re.findall(r"\d+", head)
            if group_number in numbers and room:
                group_specific = room.strip()
            continue
        if _NOT_LOCATION_RE.search(line):
            continue
        kept.append(line.strip(" ,;"))
    if group_specific:
        return group_specific
    return ", ".join(part for part in kept if part)


# ---------------------------------------------------------------------------
# GŁÓWNA LOGIKA
# ---------------------------------------------------------------------------

@dataclass
class Lesson:
    subject: str
    start: str               # "HH:MM"
    end: str                 # "HH:MM"
    dates: list[date]
    location: str
    groups: list[str]        # np. ['14'] albo ['14a']
    source: str              # np. "N33:N44"
    notes: list[str] = field(default_factory=list)


def build_weekday_map(sheet: MergedSheet, max_col: int) -> dict[int, int]:
    """Kolumna -> dzień tygodnia (0=pon). Nagłówki w wierszu 3 są scalone,
    ale nie zawsze pokrywają cały blok dnia, więc propagujemy ostatni
    znaleziony dzień w prawo."""
    result: dict[int, int] = {}
    current: int | None = None
    for col in range(FIRST_GRID_COL, max_col + 1):
        label = normalize_ws(sheet.value(ROW_DAYS, col)).upper()
        if label in WEEKDAYS_PL:
            if current != WEEKDAYS_PL[label]:
                log.info("Dzień %-13s od kolumny %s", label, get_column_letter(col))
            current = WEEKDAYS_PL[label]
        if current is not None:
            result[col] = current
    return result


def find_lessons(ws: Worksheet, group_number: str, subgroup: str | None) -> list[Lesson]:
    sheet = MergedSheet(ws)
    time_map = build_time_map(ws)
    if not time_map:
        raise RuntimeError("Nie znaleziono slotów czasowych w kolumnie A - zmienił się format pliku?")
    weekday_map = build_weekday_map(sheet, ws.max_column)
    last_time_row = max(time_map)

    lessons: list[Lesson] = []
    for row in range(FIRST_GRID_ROW, last_time_row + 1):
        for col in range(FIRST_GRID_COL, ws.max_column + 1):
            raw = ws.cell(row, col).value
            if not isinstance(raw, str):
                continue
            groups = match_group_cell(raw, group_number, subgroup)
            if not groups:
                continue

            min_row, max_row, min_col, _ = sheet.bounds(row, col)
            # Komórka w środku scalenia nie ma wartości, ale na wszelki wypadek
            # przetwarzamy tylko lewy górny róg zakresu.
            if (row, col) != (min_row, min_col):
                continue
            src = cell_ref(min_row, min_col)
            if max_row != min_row:
                src += f":{cell_ref(max_row, min_col)}"
            ctx = f"[{src}]"

            if min_row not in time_map:
                log.warning("%s: brak slotu czasu w kolumnie A dla wiersza %d - pomijam.", ctx, min_row)
                continue
            end_row = max(r for r in time_map if r <= max_row)
            start_time, end_time = time_map[min_row][0], time_map[end_row][1]

            # Metadane z kolumny komórki (z rozwiązywaniem scaleń poziomych).
            subject = normalize_ws(sheet.value(ROW_SUBJECT, col)).rstrip(" ,;:") or "Zajęcia"
            _, _, block_min, _ = sheet.bounds(ROW_SUBJECT, col)

            notes: list[str] = []
            dates_raw, dates_col = sheet.value_in_block(ROW_DATES, col, block_min)
            if dates_col is not None and dates_col != col:
                notes.append(f"daty przejęte z kolumny {get_column_letter(dates_col)}")
                log.warning("%s: pusta komórka dat - użyto kolumny %s.", ctx, get_column_letter(dates_col))
            room_raw, _ = sheet.value_in_block(ROW_ROOM, col, block_min)

            weekday = weekday_map.get(col)
            dates = parse_dates(dates_raw, weekday, ctx)

            # Sala: dopisek w samej komórce grupy (np. 'grupa 14\n5. DE 003 ...')
            # ma pierwszeństwo przed ogólną informacją z wiersza 6.
            cell_extra = ", ".join(split_lines(raw)[1:])
            location = cell_extra or extract_location(room_raw, group_number)

            log.info(
                "%s %-45s %s-%s  grupa=%-8s dat=%2d  sala=%s",
                ctx, subject[:45], start_time, end_time, ",".join(groups), len(dates),
                location or "(brak)",
            )
            if not dates:
                log.warning("%s: brak dat dla '%s' (wiersz 5: %r) - pomijam.", ctx, subject, dates_raw)
                continue

            lessons.append(Lesson(
                subject=subject,
                start=start_time,
                end=end_time,
                dates=dates,
                location=location,
                groups=groups,
                source=src,
                notes=notes + ([f"Informacje: {normalize_ws(room_raw)}"] if normalize_ws(room_raw) else [])
                + [f"Daty w planie: {normalize_ws(dates_raw)}"],
            ))
    return lessons


# ---------------------------------------------------------------------------
# GENEROWANIE ICS
# ---------------------------------------------------------------------------

def to_local_dt(day: date, hhmm: str) -> datetime:
    hour, minute = (int(x) for x in hhmm.split(":"))
    # pytz wymaga localize() - samo tzinfo=... dałoby błędny offset (LMT).
    return TIMEZONE.localize(datetime(day.year, day.month, day.day, hour, minute))


def build_calendar(lessons: list[Lesson], group_label: str, show_subgroup: bool) -> tuple[Calendar, int]:
    cal = Calendar()
    cal.add("prodid", "-//WUM-Tracker//plan zajec//PL")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    cal.add("x-wr-calname", f"WUM - {group_label}")
    cal.add("x-wr-timezone", TIMEZONE.zone)

    stamp = datetime.now(timezone.utc)
    seen: set[tuple] = set()
    count = 0
    for lesson in lessons:
        tag = "/".join(lesson.groups)
        summary = f"{lesson.subject} [{tag}]" if show_subgroup and any(g[-1].isalpha() for g in lesson.groups) else lesson.subject
        for day in lesson.dates:
            key = (day, lesson.start, lesson.end, summary)
            if key in seen:
                log.info("Duplikat %s %s %s - pomijam.", summary, day, lesson.start)
                continue
            seen.add(key)

            event = Event()
            # Deterministyczny UID -> przy ponownym imporcie/subskrypcji
            # kalendarz aktualizuje wydarzenia zamiast je dublować.
            uid_src = f"{group_label}|{day.isoformat()}|{lesson.start}|{lesson.end}|{summary}"
            event.add("uid", hashlib.sha1(uid_src.encode("utf-8")).hexdigest() + "@wum-tracker")
            event.add("dtstamp", stamp)
            event.add("dtstart", to_local_dt(day, lesson.start))
            event.add("dtend", to_local_dt(day, lesson.end))
            event.add("summary", summary)
            if lesson.location:
                event.add("location", lesson.location)
            event.add("description", "\n".join([f"Grupa: {tag}", *lesson.notes, f"Źródło: {lesson.source}"]))
            cal.add_component(event)
            count += 1

    # Dołącz definicję VTIMEZONE (lepsza zgodność z Outlookiem itp.).
    if hasattr(cal, "add_missing_timezones"):
        cal.add_missing_timezones()
    return cal, count


# ---------------------------------------------------------------------------
# WEJŚCIE / WYJŚCIE
# ---------------------------------------------------------------------------

def resolve_input(path: str | None) -> str:
    """Zwraca ścieżkę pliku Excel: podaną jawnie albo najnowszy pasujący."""
    if path:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Plik wejściowy nie istnieje: {path}")
        return path
    if os.path.isfile(INPUT_FILE):
        return INPUT_FILE
    for pattern in INPUT_GLOB_PATTERNS:
        candidates = [p for p in glob.glob(pattern) if not os.path.basename(p).startswith("~$")]
        if candidates:
            chosen = max(candidates, key=os.path.getmtime)
            log.info("Nie podano pliku - używam najnowszego pasującego: %s", chosen)
            return chosen
    raise FileNotFoundError("Nie znaleziono pliku .xlsx z planem zajęć (użyj --input).")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Konwertuje plan zajęć WUM (xlsx) do pliku .ics")
    parser.add_argument("--input", "-i", default=os.environ.get("WUM_INPUT_FILE") or None,
                        help="plik Excel z planem (domyślnie: najnowszy licencjat*.xlsx)")
    parser.add_argument("--group", "-g", default=os.environ.get("WUM_TARGET_GROUP") or TARGET_GROUP,
                        help=f"grupa dziekańska (domyślnie: '{TARGET_GROUP}')")
    parser.add_argument("--subgroup", "-s", default=os.environ.get("WUM_TARGET_SUBGROUP") or TARGET_SUBGROUP,
                        help="podgrupa a/b/c (domyślnie: wszystkie)")
    parser.add_argument("--output", "-o", default=os.environ.get("WUM_OUTPUT_FILE") or OUTPUT_FILE,
                        help=f"plik wynikowy (domyślnie: {OUTPUT_FILE})")
    parser.add_argument("--verbose", "-v", action="store_true", help="więcej logów")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stdout,
    )

    try:
        number, subgroup = parse_target(args.group, args.subgroup)
        group_label = f"grupa {number}{subgroup or ''}"
        log.info("=== WUM-Tracker: %s (podgrupa: %s) ===", group_label, subgroup or "wszystkie")

        input_path = resolve_input(args.input)
        log.info("Wczytuję %s ...", input_path)
        wb = openpyxl.load_workbook(input_path, data_only=True)
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.worksheets[0]
            log.warning("Brak arkusza '%s' - używam '%s'.", SHEET_NAME, ws.title)
        log.info("Arkusz '%s': %d wierszy x %d kolumn.", ws.title, ws.max_row, ws.max_column)

        lessons = find_lessons(ws, number, subgroup)
        log.info("Znaleziono %d bloków zajęć dla %s.", len(lessons), group_label)
        if not lessons:
            # Nie nadpisujemy poprzedniego pliku pustym kalendarzem - w CI
            # to prawie na pewno oznacza zmianę formatu pliku albo złą grupę.
            log.error("Brak zajęć dla %s - nie zapisuję pliku .ics.", group_label)
            return 1

        cal, count = build_calendar(lessons, group_label, show_subgroup=subgroup is None)
        with open(args.output, "wb") as fh:
            fh.write(cal.to_ical())
        log.info("Zapisano %d wydarzeń do %s.", count, args.output)

        # Krótkie podsumowanie tygodnia w logu CI.
        for lesson in sorted(lessons, key=lambda l: (l.dates[0].weekday(), l.start)):
            log.info(
                "  %-10s %s-%s  %-45s %-6s %d terminów",
                ["pon", "wt", "śr", "czw", "pt", "sob", "nd"][lesson.dates[0].weekday()],
                lesson.start, lesson.end, lesson.subject[:45], "/".join(lesson.groups), len(lesson.dates),
            )
        return 0
    except Exception:
        log.exception("Błąd krytyczny")
        return 2


if __name__ == "__main__":
    sys.exit(main())
