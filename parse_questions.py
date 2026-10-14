#!/usr/bin/env python3
"""
Парсер сборников задач («ЧАСТЬ 2 / В1 ... / ЧАСТЬ 3 / C1 ...») в Markdown и Excel.

Конвейер:
    PDF --(OCR-движок)--> «сырой» md --(regex)--> questions.md + questions.xlsx

Движки OCR (--engine):
    marker     бесплатно, локально, формулы в LaTeX   (pip install marker-pdf)   <- основной
    mathpix    платный API, ключи MATHPIX_APP_ID / MATHPIX_APP_KEY
    tesseract  бесплатно, формулы без LaTeX            (apt install tesseract-ocr tesseract-ocr-rus poppler-utils)
    auto       (по умолчанию) marker -> mathpix (если есть ключи) -> tesseract

Примеры:
    python parse_questions.py input.pdf -o out                  # PDF, движок выберется сам
    python parse_questions.py input.pdf -o out --engine marker
    python parse_questions.py input.md  -o out                  # готовый md (Mathpix или marker)
    python parse_questions.py out/raw.md -o out --ocr             # повторный разбор уже распознанного текста
    python parse_questions.py input.pdf -o out --engine tesseract --fixes fixes.txt --tess-lang rus+eng+ell

Зависимость для всех режимов: openpyxl.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# --------------------------------------------------------------------------- #
#  Регулярные выражения
# --------------------------------------------------------------------------- #

# Заголовок части. Строка должна НАЧИНАТЬСЯ с него (так «часть 2 пути» внутри
# предложения не сработает). Допустимые обёртки:
#   Mathpix:  \title{ ЧАСТЬ 2 }   \section*{ЧАСТЬ 3}
#   marker:   ## ЧАСТЬ 2   **ЧАСТЬ 2**
RE_PART = re.compile(
    r"^[ \t]*"
    r"(?:\\(?:title|section\*?)\{\s*|[#*_]+[ \t]*)?"
    r"(?:ЧАСТЬ|Часть)\s+(?P<part>\d+)"
    r"[ \t]*[*_]*[ \t]*\}?",
    re.MULTILINE,
)

# Номер вопроса в начале строки: B1, В1 (кирилл.), Β1 (греч.), C1, С1 (кирилл.).
# Допускаем мусор вокруг: «**B1**», «- B1», «[ B1 |», «B1.», «B1)» и буквы-двойники
# вместо цифр (S, O, I, З) – их расшифровывает resolve_number().
RE_QUESTION = re.compile(
    r"^[ \t>#*_\[|\-]*"
    r"(?P<num>[BВΒCС]\s?[\dSOIЗ]{1,3})(?![A-Za-zА-Яа-я])"
    r"[ \t]*[.)|\]*_:]*[ \t]*"
    r"(?=\S)",
    re.MULTILINE,
)

# «Рамка» вопроса из OCR: «[οἱ |  текст». Номер не распознан, но разделитель «|» остался.
# Используется только в режиме OCR; номер тогда выводится по порядку.
RE_QUESTION_FRAME = re.compile(
    r"^[ \t\[]*(?P<num>[^\s|\]]{1,3})[ \t]*\|[ \t]*(?=\S)",
    re.MULTILINE,
)

# Картинка Markdown: ![](url) или ![alt](url)
RE_IMAGE = re.compile(r"!\[[^\]]*\]\((?P<url>[^)\s]+)\)")

# Служебное из marker: <span id="page-0-0"></span>,  <math ...>...</math>
RE_SPAN = re.compile(r"<span\b[^>]*></span>")
RE_MATH = re.compile(r"<math\b[^>]*>(.*?)</math>", re.DOTALL)

RE_MANY_NL = re.compile(r"\n{4,}")
RE_SOFT_WRAP = re.compile(r"(?<!\n)\n(?!\n)")      # одиночный перенос строки
RE_TAIL_JUNK = re.compile(r"[\s#*_>|\-]+$")        # «## », «**» в конце куска


# --------------------------------------------------------------------------- #
#  Исправление типичных ошибок OCR
# --------------------------------------------------------------------------- #
# Встроенные (общие) правила «регулярка -> замена». Правила под конкретный скан
# лежат в fixes.txt (--fixes). Формат файла:  регулярка => замена   (# – комментарий)
DEFAULT_FIXES = [
    (r"(?<![\w])[AАΛ]\s?t(?=\s*[=,])", "Δt"),            # «А t = 10»  -> «Δt = 10»
    (r"(?<![\w/])[BВ](?=\s*=\s*\d)", "R"),               # «В = 100»   -> «R = 100»
    (r"(?<![\w])[аА](?=\s*=\s*\d)", "a"),                # кириллическая «а = 0,3» -> латинская
    (r"(?<![\w])Т(?=\s*=)", "T"),                         # кириллическая «Т =» -> латинская
    (r"(?<=м/с)\?", "²"),                                  # «м/с?» -> «м/с²»
    (r"\bΜΜ\b", "мм"),                                    # греческие «ΜΜ» -> «мм»
    (r"\bPuc\.", "Рис."),
    (r"(?<=\d)(?=(?:мин|мм|км|см|м/с|рад/с)(?![а-яё]))", " "),  # «1мин» -> «1 мин»
    (r"(?m)^\|[ \t]*", ""),                                # одинокая «|» в начале строки
    (r"(?m)^[ \t]*[.|]+[ \t]*$", ""),                      # строки из одной точки / палки
]


def load_fixes(path: Path | None) -> list[tuple[str, str]]:
    rules = list(DEFAULT_FIXES)
    if path:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=>" not in line:
                continue
            pat, rep = line.split("=>", 1)
            rules.append((pat.strip(), rep.strip()))
    return rules


def apply_fixes(text: str, rules: list[tuple[str, str]]) -> str:
    for pat, rep in rules:
        text = re.sub(pat, rep, text)
    return text


# Серия вопросов по букве (латиница, кириллица и греческая Β – одна серия)
LETTER_SERIES = {"B": "B", "В": "B", "Β": "B", "C": "C", "С": "C"}
# Какая серия в какой части (структура ЕГЭ-подобных сборников: часть 2 -> B, часть 3 -> C)
PART_SERIES = {"1": "A", "2": "B", "3": "C"}
# Буквы-двойники в номере и какие цифры они могут означать
LOOKALIKE = {"S": "85", "O": "0", "I": "1", "З": "3"}
RE_NUMBER = re.compile(r"^([BВΒCС])\s?([\dSOIЗ]{1,3})$")


def resolve_number(raw: str, last: dict, part: str, latin: bool) -> str:
    """
    Приводит номер к виду «B8».
      * «BS» -> B5 или B8 – по тому, какой номер ожидается по порядку;
      * непрочитанный номер («οἱ») -> следующий по порядку в серии этой части.
    """
    raw = re.sub(r"\s+", "", raw)
    m = RE_NUMBER.match(raw)
    if m:
        letter, rest = m.groups()
        series = LETTER_SERIES[letter]
        expected = last.get(series, 0) + 1
        options = [""]
        for ch in rest:
            options = [o + d for o in options for d in LOOKALIKE.get(ch, ch)]
        pick = next((o for o in options if o.isdigit() and int(o) == expected), options[0])
    else:                                    # номер не прочитан – берём по порядку
        series = PART_SERIES.get(part, "B")
        letter = series
        pick = str(last.get(series, 0) + 1)
    if pick.isdigit():
        last[series] = int(pick)
    return (series if latin else letter) + pick


@dataclass
class Question:
    part: str
    number: str
    text: str
    figure: str = ""


# --------------------------------------------------------------------------- #
#  Подготовка текста и парсинг
# --------------------------------------------------------------------------- #
def normalize_marker(md: str) -> str:
    """Приводит разметку marker к виду, близкому к Mathpix."""
    md = RE_SPAN.sub("", md)
    md = RE_MATH.sub(lambda m: f"${m.group(1).strip()}$", md)
    return md


def clean_ocr_body(text: str) -> str:
    """Для OCR: склеить переносы слов («поло-\\nвину»), все переносы строк -> пробел."""
    text = re.sub(r"(?<=[а-яёА-ЯЁ])-[ \t]*\n\s*(?=[а-яё])", "", text)
    text = re.sub(r"\s*\n\s*", " ", text)
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def parse_markdown(md: str, unwrap: bool = False, ocr: bool = False) -> list[Question]:
    """
    Разбивает текст на вопросы, запоминая текущую «ЧАСТЬ».
      unwrap – склеить одиночные переносы строк внутри вопроса;
      ocr    – режим для распознанного текста: искать «рамочные» маркеры, выводить
               нечитаемые номера по порядку, склеивать переносы слов, буквы -> латиница.
    """
    md = normalize_marker(md)

    events = []  # (начало, конец, тип, значение)
    for m in RE_PART.finditer(md):
        events.append((m.start(), m.end(), "part", m.group("part")))
    for m in RE_QUESTION.finditer(md):
        events.append((m.start(), m.end(), "q", m.group("num")))
    if ocr:
        for m in RE_QUESTION_FRAME.finditer(md):
            events.append((m.start(), m.end(), "q", m.group("num")))
    events.sort(key=lambda e: e[0])          # сортировка устойчивая: при равном начале первым остаётся RE_QUESTION

    cleaned, seen = [], set()
    for ev in events:
        if ev[0] in seen:
            continue
        seen.add(ev[0])
        cleaned.append(ev)
    events = cleaned

    questions: list[Question] = []
    current_part = ""
    last: dict = {}
    for i, (start, end, kind, value) in enumerate(events):
        if kind == "part":
            current_part = value
            continue
        number = resolve_number(value, last, current_part, latin=ocr)
        body_end = events[i + 1][0] if i + 1 < len(events) else len(md)
        body = md[end:body_end]

        img = RE_IMAGE.search(body)
        figure = img.group("url") if img else ""
        body = RE_IMAGE.sub("", body)
        body = RE_TAIL_JUNK.sub("", body).strip()
        if ocr:
            body = clean_ocr_body(body)
        elif unwrap:
            body = RE_SOFT_WRAP.sub(" ", body)

        questions.append(Question(current_part, number, body, figure))
    return questions


def check_sequence(questions: list[Question]) -> None:
    """Предупреждает, если в нумерации есть пропуски (значит, что-то не распозналось)."""
    by_series: dict[str, list[int]] = {}
    for q in questions:
        by_series.setdefault(LETTER_SERIES.get(q.number[0], q.number[0]), []).append(int(q.number[1:]))
    for series, nums in by_series.items():
        missing = sorted(set(range(1, max(nums) + 1)) - set(nums))
        if missing:
            print(f"ВНИМАНИЕ: в серии {series} нет номеров {missing}", file=sys.stderr)
        if len(nums) != len(set(nums)):
            print(f"ВНИМАНИЕ: в серии {series} есть повторяющиеся номера", file=sys.stderr)


# --------------------------------------------------------------------------- #
#  Вывод
# --------------------------------------------------------------------------- #
def write_markdown(questions: list[Question], path: Path) -> None:
    lines = ["# Задачи", ""]
    last_part = None
    for q in questions:
        if q.part != last_part:
            lines += [f"## Часть {q.part}", ""]
            last_part = q.part
        lines += [f"### {q.number}", "", RE_MANY_NL.sub("\n\n", q.text)]
        if q.figure:
            lines += ["", f"![]({q.figure})"]
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_excel(questions: list[Question], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Часть", "Номер вопроса", "Вопрос", "Рисунок"])
    for q in questions:
        ws.append([q.part, q.number, q.text, q.figure or None])

    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial")
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCD", (8, 16, 100, 45)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


# --------------------------------------------------------------------------- #
#  OCR-движки: PDF -> md
# --------------------------------------------------------------------------- #
def pdf_to_md_marker(pdf: Path, workdir: Path) -> str:
    """marker (Surya). Запускаем CLI marker_single; --force_ocr нужен для сканов."""
    exe = shutil.which("marker_single")
    if not exe:
        raise RuntimeError("marker_single не найден. Установите: pip install marker-pdf")
    out_dir = workdir / "marker_out"
    subprocess.run(
        [exe, str(pdf), "--output_dir", str(out_dir), "--force_ocr"],
        check=True,
    )
    mds = sorted(out_dir.rglob("*.md"))
    if not mds:
        raise RuntimeError(f"marker не создал .md в {out_dir}")
    return mds[0].read_text(encoding="utf-8")


def pdf_to_md_mathpix(pdf: Path, workdir: Path) -> str:
    """Mathpix /v3/pdf. Нужны MATHPIX_APP_ID и MATHPIX_APP_KEY, pip install requests."""
    import json
    import time

    import requests

    headers = {"app_id": os.environ["MATHPIX_APP_ID"], "app_key": os.environ["MATHPIX_APP_KEY"]}
    opts = {"conversion_formats": {"md": True}, "math_inline_delimiters": ["$", "$"]}
    with open(pdf, "rb") as f:
        r = requests.post("https://api.mathpix.com/v3/pdf", headers=headers,
                          data={"options_json": json.dumps(opts)}, files={"file": f})
    r.raise_for_status()
    pdf_id = r.json()["pdf_id"]
    while True:
        st = requests.get(f"https://api.mathpix.com/v3/pdf/{pdf_id}", headers=headers).json()
        if st.get("status") == "completed":
            break
        if st.get("status") == "error":
            raise RuntimeError(st)
        time.sleep(2)
    return requests.get(f"https://api.mathpix.com/v3/pdf/{pdf_id}.md", headers=headers).text


TESS_LANG = "rus+eng"


def pdf_to_md_tesseract(pdf: Path, workdir: Path) -> str:
    """Растеризация + Tesseract (rus+eng). Формулы будут обычным текстом."""
    pages = []
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", "300", "-png", str(pdf), f"{tmp}/p"], check=True)
        for img in sorted(Path(tmp).glob("p*.png")):
            out = subprocess.run(
                ["tesseract", str(img), "-", "-l", TESS_LANG, "--psm", "4"],
                check=True, capture_output=True, text=True,
            ).stdout
            pages.append(out)
    return "\n".join(pages)


def pick_engine(requested: str) -> str:
    if requested != "auto":
        return requested
    if shutil.which("marker_single"):
        return "marker"
    if os.environ.get("MATHPIX_APP_ID") and os.environ.get("MATHPIX_APP_KEY"):
        return "mathpix"
    return "tesseract"


ENGINES = {"marker": pdf_to_md_marker, "mathpix": pdf_to_md_mathpix, "tesseract": pdf_to_md_tesseract}


# --------------------------------------------------------------------------- #
def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path, help="PDF или готовый Markdown")
    ap.add_argument("-o", "--out", type=Path, default=Path("."), help="папка для результатов")
    ap.add_argument("--engine", choices=["auto", *ENGINES], default="auto", help="OCR-движок для PDF")
    ap.add_argument("--unwrap", action="store_true",
                    help="склеивать одиночные переносы строк внутри вопроса")
    ap.add_argument("--ocr", action="store_true",
                    help="режим распознанного текста (tesseract/marker): правила исправлений, "
                         "номера по порядку, склейка переносов. Включается сам для PDF; "
                         "для готового raw.md задайте вручную")
    ap.add_argument("--fixes", type=Path, help="файл с правилами исправления OCR: «регулярка => замена»")
    ap.add_argument("--tess-lang", default="rus+eng",
                    help="языки Tesseract, напр. rus+eng+ell (ell = греческий, помогает с Δ и φ)")
    args = ap.parse_args()
    global TESS_LANG
    TESS_LANG = args.tess_lang

    args.out.mkdir(parents=True, exist_ok=True)
    ocr = args.ocr

    if args.input.suffix.lower() == ".pdf":
        engine = pick_engine(args.engine)
        print(f"PDF -> MD через {engine}", file=sys.stderr)
        md = ENGINES[engine](args.input, args.out)
        (args.out / "raw.md").write_text(md, encoding="utf-8")
        ocr = ocr or engine != "mathpix"
    else:
        md = args.input.read_text(encoding="utf-8")

    # правила исправления OCR: в режиме OCR всегда, иначе – только если задан --fixes
    fixes_path = args.fixes
    if fixes_path is None:
        for cand in (Path.cwd() / "fixes.txt", Path(__file__).resolve().parent / "fixes.txt"):
            if cand.exists():
                fixes_path = cand
                break
    if ocr or args.fixes:
        if fixes_path:
            print(f"Правила исправлений: {fixes_path}", file=sys.stderr)
        md = apply_fixes(md, load_fixes(fixes_path))
        (args.out / "raw_fixed.md").write_text(md, encoding="utf-8")

    questions = parse_markdown(md, unwrap=args.unwrap, ocr=ocr)
    if not questions:
        sys.exit("Вопросы не найдены – посмотрите raw.md и проверьте формат")
    check_sequence(questions)

    write_markdown(questions, args.out / "questions.md")
    write_excel(questions, args.out / "questions.xlsx")
    print(f"Найдено вопросов: {len(questions)} -> {args.out}/questions.md, questions.xlsx")


if __name__ == "__main__":
    main()