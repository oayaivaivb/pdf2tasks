#!/usr/bin/env python3
"""CLI-утилита для парсинга сборников задач из PDF/MD в Markdown и Excel."""

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

# Поиск заголовков частей, номеров вопросов и разметки[cite: 1]
RE_PART = re.compile(
    r"^[ \t]*(?:\\(?:title|section\*?)\{\s*|[#*_]+[ \t]*)?(?:ЧАСТЬ|Часть)\s+(?P<part>\d+)[ \t]*[*_]*[ \t]*\}?",
    re.MULTILINE,
)
RE_QUESTION = re.compile(
    r"^[ \t>#*_\[|\-]*(?P<num>[BВΒCС]\s?[\dSOIЗ]{1,3})(?![A-Za-zА-Яа-я])[ \t]*[.)|\]*_:]*[ \t]*(?=\S)",
    re.MULTILINE,
)
RE_QUESTION_FRAME = re.compile(r"^[ \t\[]*(?P<num>[^\s|\]]{1,3})[ \t]*\|[ \t]*(?=\S)", re.MULTILINE)
RE_IMAGE = re.compile(r"!\[[^\]]*\]\((?P<url>[^)\s]+)\)")
RE_SPAN = re.compile(r"<span\b[^>]*></span>")
RE_MATH = re.compile(r"<math\b[^>]*>(.*?)</math>", re.DOTALL)
RE_MANY_NL = re.compile(r"\n{4,}")
RE_SOFT_WRAP = re.compile(r"(?<!\n)\n(?!\n)")
RE_TAIL_JUNK = re.compile(r"[\s#*_>|\-]+$")

# Типовые замены OCR-артефактов в формулах и единицах[cite: 1]
DEFAULT_FIXES = [
    (r"(?<![\w])[AАΛ]\s?t(?=\s*[=,])", "Δt"),
    (r"(?<![\w/])[BВ](?=\s*=\s*\d)", "R"),
    (r"(?<![\w])[аА](?=\s*=\s*\d)", "a"),
    (r"(?<![\w])Т(?=\s*=)", "T"),
    (r"(?<=м/с)\?", "²"),
    (r"\bΜΜ\b", "мм"),
    (r"\bPuc\.", "Рис."),
    (r"(?<=\d)(?=(?:мин|мм|км|см|м/с|рад/с)(?![а-яё]))", " "),
    (r"(?m)^\|[ \t]*", ""),
    (r"(?m)^[ \t]*[.|]+[ \t]*$", ""),
]

LETTER_SERIES = {"B": "B", "В": "B", "Β": "B", "C": "C", "С": "C"}
PART_SERIES = {"1": "A", "2": "B", "3": "C"}
LOOKALIKE = {"S": "85", "O": "0", "I": "1", "З": "3"}
RE_NUMBER = re.compile(r"^([BВΒCС])\s?([\dSOIЗ]{1,3})$")


@dataclass
class Question:
    part: str
    number: str
    text: str
    figure: str = ""


def load_fixes(path: Path | None) -> list[tuple[str, str]]:
    """Читает правила замены из файла (формат: regex => repl)."""
    rules = list(DEFAULT_FIXES)
    if not path or not path.exists():
        return rules
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=>" not in line:
            continue
        pat, rep = line.split("=>", 1)
        rules.append((pat.strip(), rep.strip()))
    return rules


def apply_fixes(text: str, rules: list[tuple[str, str]]) -> str:
    """Применяет список regex-замен к тексту."""
    for pat, rep in rules:
        text = re.sub(pat, rep, text)
    return text


def resolve_number(raw: str, last: dict, part: str, latin: bool) -> str:
    """Нормализует номер задачи, чинит буквы-двойники (S->8/5) или считает по порядку."""
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
    else:
        series = PART_SERIES.get(part, "B")
        letter = series
        pick = str(last.get(series, 0) + 1)

    if pick.isdigit():
        last[series] = int(pick)
    return (series if latin else letter) + pick


def clean_ocr_body(text: str) -> str:
    """Удаляет переносы слов и заменяет одиночные переводы строк на пробелы."""
    text = re.sub(r"(?<=[а-яёА-ЯЁ])-[ \t]*\n\s*(?=[а-яё])", "", text)
    text = re.sub(r"\s*\n\s*", " ", text)
    return re.sub(r"[ \t]{2,}", " ", text).strip()


def parse_markdown(md: str, unwrap: bool = False, ocr: bool = False) -> list[Question]:
    """Разбивает текст на разделы и блоки задач."""
    md = RE_SPAN.sub("", md)
    md = RE_MATH.sub(lambda m: f"${m.group(1).strip()}$", md)

    events = []
    for m in RE_PART.finditer(md):
        events.append((m.start(), m.end(), "part", m.group("part")))
    for m in RE_QUESTION.finditer(md):
        events.append((m.start(), m.end(), "q", m.group("num")))
    if ocr:
        for m in RE_QUESTION_FRAME.finditer(md):
            events.append((m.start(), m.end(), "q", m.group("num")))
    events.sort(key=lambda e: e[0])

    cleaned, seen = [], set()
    for ev in events:
        if ev[0] not in seen:
            seen.add(ev[0])
            cleaned.append(ev)

    questions: list[Question] = []
    current_part = ""
    last: dict = {}

    for i, (start, end, kind, value) in enumerate(cleaned):
        if kind == "part":
            current_part = value
            continue

        number = resolve_number(value, last, current_part, latin=ocr)
        body_end = cleaned[i + 1][0] if i + 1 < len(cleaned) else len(md)
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
    """Выводит предупреждения при пропусках или дублях номеров."""
    by_series: dict[str, list[int]] = {}
    for q in questions:
        key = LETTER_SERIES.get(q.number[0], q.number[0])
        by_series.setdefault(key, []).append(int(q.number[1:]))

    for series, nums in by_series.items():
        missing = sorted(set(range(1, max(nums) + 1)) - set(nums))
        if missing:
            print(f"Warning: пропущены {series}: {missing}", file=sys.stderr)
        if len(nums) != len(set(nums)):
            print(f"Warning: дубликаты в {series}", file=sys.stderr)


def write_markdown(questions: list[Question], path: Path) -> None:
    """Сохраняет структурированный список в MD."""
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
    """Сохраняет задачи в Excel с базовым автоформатированием."""
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
    for col, width in zip("ABCD", (8, 16, 100, 45)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(path)


def pdf_to_md_marker(pdf: Path, workdir: Path) -> str:
    """OCR через marker-pdf."""
    exe = shutil.which("marker_single")
    if not exe:
        raise RuntimeError("marker_single не найден")
    out_dir = workdir / "marker_out"
    subprocess.run([exe, str(pdf), "--output_dir", str(out_dir), "--force_ocr"], check=True)
    mds = sorted(out_dir.rglob("*.md"))
    if not mds:
        raise RuntimeError(f"marker не создал .md в {out_dir}")
    return mds[0].read_text(encoding="utf-8")


def pdf_to_md_mathpix(pdf: Path, workdir: Path) -> str:
    """OCR через Mathpix API."""
    import json, time, requests

    headers = {"app_id": os.environ["MATHPIX_APP_ID"], "app_key": os.environ["MATHPIX_APP_KEY"]}
    opts = {"conversion_formats": {"md": True}, "math_inline_delimiters": ["$", "$"]}
    with open(pdf, "rb") as f:
        r = requests.post(
            "https://api.mathpix.com/v3/pdf",
            headers=headers,
            data={"options_json": json.dumps(opts)},
            files={"file": f},
        )
    r.raise_for_status()
    pdf_id = r.json()["pdf_id"]

    while True:
        status_data = requests.get(f"https://api.mathpix.com/v3/pdf/{pdf_id}", headers=headers).json()
        if status_data.get("status") == "completed":
            break
        if status_data.get("status") == "error":
            raise RuntimeError(status_data)
        time.sleep(2)

    return requests.get(f"https://api.mathpix.com/v3/pdf/{pdf_id}.md", headers=headers).text


TESS_LANG = "rus+eng"


def pdf_to_md_tesseract(pdf: Path, workdir: Path) -> str:
    """Растеризация страниц (pdftoppm) и OCR через Tesseract."""
    pages = []
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", "300", "-png", str(pdf), f"{tmp}/p"], check=True)
        for img in sorted(Path(tmp).glob("p*.png")):
            out = subprocess.run(
                ["tesseract", str(img), "-", "-l", TESS_LANG, "--psm", "4"],
                check=True,
                capture_output=True,
                text=True,
            ).stdout
            pages.append(out)
    return "\n".join(pages)


def pick_engine(requested: str) -> str:
    """Выбор движка OCR по наличию бинарников или ключей."""
    if requested != "auto":
        return requested
    if shutil.which("marker_single"):
        return "marker"
    if os.environ.get("MATHPIX_APP_ID") and os.environ.get("MATHPIX_APP_KEY"):
        return "mathpix"
    return "tesseract"


ENGINES = {"marker": pdf_to_md_marker, "mathpix": pdf_to_md_mathpix, "tesseract": pdf_to_md_tesseract}


def main() -> None:
    ap = argparse.ArgumentParser(description="Парсер сборников задач в MD и Excel")
    ap.add_argument("input", type=Path, help="Входной PDF или Markdown")
    ap.add_argument("-o", "--out", type=Path, default=Path("."), help="Папка для результатов")
    ap.add_argument("--engine", choices=["auto", *ENGINES], default="auto", help="Движок OCR")
    ap.add_argument("--unwrap", action="store_true", help="Схлопывать одиночные переносы внутри условий")
    ap.add_argument("--ocr", action="store_true", help="Режим чистки OCR (переносы, рамочные номера)")
    ap.add_argument("--fixes", type=Path, help="Файл правок (по умолчанию fixes.txt)")
    ap.add_argument("--tess-lang", default="rus+eng", help="Языки Tesseract (напр. rus+eng+ell)")
    args = ap.parse_args()

    global TESS_LANG
    TESS_LANG = args.tess_lang

    args.out.mkdir(parents=True, exist_ok=True)
    ocr = args.ocr

    if args.input.suffix.lower() == ".pdf":
        engine = pick_engine(args.engine)
        md = ENGINES[engine](args.input, args.out)
        (args.out / "raw.md").write_text(md, encoding="utf-8")
        ocr = ocr or engine != "mathpix"
    else:
        md = args.input.read_text(encoding="utf-8")

    fixes_path = args.fixes
    if fixes_path is None:
        for cand in (Path.cwd() / "fixes.txt", Path(__file__).resolve().parent / "fixes.txt"):
            if cand.exists():
                fixes_path = cand
                break

    if ocr or args.fixes:
        md = apply_fixes(md, load_fixes(fixes_path))
        (args.out / "raw_fixed.md").write_text(md, encoding="utf-8")

    questions = parse_markdown(md, unwrap=args.unwrap, ocr=ocr)
    if not questions:
        sys.exit("Задачи не найдены")

    check_sequence(questions)
    write_markdown(questions, args.out / "questions.md")
    write_excel(questions, args.out / "questions.xlsx")
    print(f"Готово: {len(questions)} задач -> {args.out}")


if __name__ == "__main__":
    main()