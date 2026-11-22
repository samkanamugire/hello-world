#!/usr/bin/env python3
"""KFHE screed WhatsApp extractor (read-only).

Reads WhatsApp "Export chat" files (.zip or .txt, Android or iPhone format),
keeps messages from START_DATE to today, sorts them into the sheets
Attendance, Progress, Materials, Plant, Money, Instructions, Other and writes:

  KFHE_WhatsApp_extract_<today>.xlsx
  KFHE_WhatsApp_extract_<today>.md
  images/   (photos sent in the period, for transcribing name lists / payroll)

Nothing is guessed: dates or numbers that cannot be read are written UNCLEAR,
and the original message text is always kept in the Detail column.

Usage:
  python kfhe_extract.py OUTPUT_FOLDER EXPORT1.zip [EXPORT2.zip ...]
  python kfhe_extract.py OUTPUT_FOLDER exports_folder/
"""
import datetime as dt
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

START_DATE = dt.date(2026, 10, 16)
SHEETS = ["Attendance", "Progress", "Materials", "Plant", "Money", "Instructions", "Other"]
COLUMNS = ["Date", "Chat", "Sender", "Item", "Qty", "Unit", "Detail (verbatim text)"]

# Keyword rules, checked in this order; a message goes to the first sheet that matches.
RULES = [
    ("Money", ["payroll", "payment", "paid", "pay ", "salary", "advance", "avance", "rwf", "frw",
               "bpr", "momo", "transfer", "restaurant", "deduction", "rate", "invoice", "lpo amount",
               "amafaranga", "kwishyura"]),
    ("Instructions", ["mix ratio", "ratio", "fibre dose", "fiber dose", "dosage", "layer", "thickness",
                      "method statement", "lpo", "instruction", "approved", "rejected", "comment",
                      "mace", "specification", "spec "]),
    ("Plant", ["mixer", "dumper", "breakdown", "broken", "fuel", "diesel", "petrol", "generator",
               "vibrator", "power float", "plant"]),
    ("Materials", ["cement", "sand", "fibre", "fiber", "mesh", "brc", "bags", "delivery", "delivered",
                   "stock", "supplier", "aggregate", "water proof", "dpm", "curing compound", "isima",
                   "umucanga"]),
    ("Progress", ["grid", "screed", "inspection", "wir", "cleaning", "curing", "placed", "poured",
                  "m2", "m²", "m3", "m³", "area", "level", "floor", "completed", "progress"]),
    ("Attendance", ["attendance", "present", "absent", "workers", "abakozi", "names", "timekeep",
                    "time sheet", "timesheet", "list"]),
]

# Quantity + unit patterns (first match wins)
QTY_RE = re.compile(
    r"(?P<qty>\d{1,3}(?:[,\s]\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*"
    r"(?P<unit>rwf|frw|m2|m²|sqm|m3|m³|kg|kgs|bags?|sacks?|tons?|t\b|trips?|litres?|liters?|ltrs?|l\b|pcs|sheets?|rolls?|people|workers)",
    re.IGNORECASE,
)

ANDROID_RE = re.compile(r"^(\d{1,2}[/.]\d{1,2}[/.]\d{2,4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s?[APap]\.?[Mm]\.?)?)\s+-\s+(.*)$")
IOS_RE = re.compile(r"^\u200e?\[(\d{1,2}[/.]\d{1,2}[/.]\d{2,4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s?[APap]\.?[Mm]\.?)?)\]\s+(.*)$")
MEDIA_RE = re.compile(r"(?:<attached:\s*([^>]+)>|([\w\-. ]+\.(?:jpe?g|png|webp|heic|pdf|xlsx?|docx?|opus|mp4))\s*\(file attached\))", re.IGNORECASE)
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".heic"}


def parse_date(s, dayfirst=True):
    parts = re.split(r"[/.]", s)
    try:
        a, b, y = (int(p) for p in parts)
    except ValueError:
        return None
    if y < 100:
        y += 2000
    d, m = (a, b) if dayfirst else (b, a)
    try:
        return dt.date(y, m, d)
    except ValueError:
        return None


def guess_dayfirst(date_strings):
    """Rwanda/UK phones are day-first; switch to month-first only if the data proves it."""
    for s in date_strings:
        a, b = (int(p) for p in re.split(r"[/.]", s)[:2])
        if a > 12:
            return True
        if b > 12:
            return False
    return True


def read_messages(txt):
    lines = txt.splitlines()
    raw = []
    for line in lines:
        m = IOS_RE.match(line) or ANDROID_RE.match(line)
        if m:
            raw.append([m.group(1), m.group(2), m.group(3)])
        elif raw:
            raw[-1][2] += "\n" + line
    dayfirst = guess_dayfirst([r[0] for r in raw])
    msgs = []
    for d, t, rest in raw:
        if ": " in rest:
            sender, text = rest.split(": ", 1)
        else:
            sender, text = "(system)", rest
        msgs.append({"date": parse_date(d, dayfirst), "date_raw": d, "time": t,
                     "sender": sender.strip("\u200e ").strip(), "text": text.replace("\u200e", "").strip()})
    return msgs


def classify(text):
    low = " " + text.lower() + " "
    for sheet, words in RULES:
        if any(w in low for w in words):
            return sheet
    return "Other"


def qty_unit(text):
    m = QTY_RE.search(text)
    if not m:
        return "", ""
    q = m.group("qty").replace(",", "").replace(" ", "")
    unit = m.group("unit").lower()
    unit = {"m²": "m2", "sqm": "m2", "m³": "m3", "frw": "RWF", "rwf": "RWF", "kgs": "kg",
            "l": "litres", "ltr": "litres", "ltrs": "litres", "liter": "litres", "liters": "litres",
            "litre": "litres", "bag": "bags", "sack": "bags", "sacks": "bags"}.get(unit, unit)
    try:
        q = float(q)
        q = int(q) if q.is_integer() else q
    except ValueError:
        q = "UNCLEAR"
    return q, unit


def item_label(text, sheet):
    first = text.strip().splitlines()[0] if text.strip() else ""
    return (first[:80] + "…") if len(first) > 80 else first


def load_export(path, workdir):
    """Return (chat_name, txt, media_dir) for a .zip or .txt export."""
    path = Path(path)
    if path.suffix.lower() == ".zip":
        dest = Path(workdir) / path.stem
        dest.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                name = Path(info.filename).name
                if not name or info.is_dir():
                    continue
                with z.open(info) as src, open(dest / name, "wb") as out:  # flatten: no path traversal
                    shutil.copyfileobj(src, out)
        txts = sorted(dest.glob("*.txt"))
        if not txts:
            raise SystemExit(f"No chat .txt inside {path}")
        txt_path = txts[0]
        media_dir = dest
    else:
        txt_path = path
        media_dir = path.parent
    chat = path.stem
    chat = re.sub(r"^WhatsApp Chat (with|-)\s*", "", chat, flags=re.IGNORECASE).strip() or chat
    return chat, txt_path.read_text(encoding="utf-8", errors="replace"), media_dir


def main(out_dir, inputs):
    today = dt.date.today()
    out = Path(out_dir)
    img_dir = out / "images"
    img_dir.mkdir(parents=True, exist_ok=True)
    work = out / ".unzipped"
    files = []
    for p in inputs:
        p = Path(p)
        files += sorted(list(p.glob("*.zip")) + list(p.glob("*.txt"))) if p.is_dir() else [p]

    rows = {s: [] for s in SHEETS}
    stats = {"messages": 0, "images": 0, "unclear_dates": 0, "chats": []}
    for f in files:
        chat, txt, media_dir = load_export(f, work)
        stats["chats"].append(chat)
        for m in read_messages(txt):
            d = m["date"]
            if d is None:
                stats["unclear_dates"] += 1
                date_out = f"UNCLEAR ({m['date_raw']})"
            elif d < START_DATE or d > today:
                continue
            else:
                date_out = d.isoformat()
            text = m["text"]
            if not text or "messages and calls are end-to-end encrypted" in text.lower():
                continue
            stats["messages"] += 1
            media = MEDIA_RE.search(text)
            if media:
                fname = (media.group(1) or media.group(2)).strip()
                src = Path(media_dir) / fname
                if src.suffix.lower() in IMG_EXT:
                    if src.exists():
                        new = f"{date_out[:10]}_{re.sub(r'[^A-Za-z0-9]+', '_', chat)}_{src.name}"
                        shutil.copy2(src, img_dir / new)
                        stats["images"] += 1
                        detail = f"{text}\n[saved as images/{new} — TO TRANSCRIBE]"
                    else:
                        detail = f"{text}\n[image not in export — UNCLEAR]"
                    caption = MEDIA_RE.sub("", text).strip()
                    sheet = classify(caption) if caption else "Attendance"  # name lists / payroll photos
                    rows[sheet].append([date_out, chat, m["sender"], "Photo: " + (caption or "image"), "", "", detail])
                    continue
            sheet = classify(text)
            q, u = qty_unit(text)
            rows[sheet].append([date_out, chat, m["sender"], item_label(text, sheet), q, u, text])

    for s in SHEETS:
        rows[s].sort(key=lambda r: (r[0], r[1]))
    stem = f"KFHE_WhatsApp_extract_{today.isoformat()}"
    write_xlsx(out / f"{stem}.xlsx", rows)
    write_md(out / f"{stem}.md", rows, stats, today)
    shutil.rmtree(work, ignore_errors=True)
    print(f"Saved to: {out.resolve()}")
    print(f"Chats: {', '.join(stats['chats']) or 'none'}")
    print(f"Messages in period: {stats['messages']}  |  images saved: {stats['images']}  |  unclear dates: {stats['unclear_dates']}")
    for s in SHEETS:
        print(f"  {s}: {len(rows[s])} rows")


def write_xlsx(path, rows):
    wb = Workbook()
    wb.remove(wb.active)
    head_fill = PatternFill("solid", fgColor="1F4E78")
    widths = [14, 22, 18, 45, 10, 8, 90]
    for s in SHEETS:
        ws = wb.create_sheet(s)
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
        for r in rows[s]:
            ws.append(r)
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        for row in ws.iter_rows(min_row=2):
            row[6].alignment = Alignment(wrap_text=True, vertical="top")
            for c in row[:6]:
                c.alignment = Alignment(vertical="top")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def md_cell(v):
    return str(v).replace("|", "\\|").replace("\n", "<br>")


def write_md(path, rows, stats, today):
    L = [f"# KFHE screed – WhatsApp extract ({START_DATE.isoformat()} to {today.isoformat()})", "",
         "## Summary", "",
         f"- Chats read: {', '.join(stats['chats']) or 'none'}",
         f"- Messages in period: {stats['messages']}; photos saved to `images/`: {stats['images']}",
         f"- Messages with unreadable dates (marked UNCLEAR): {stats['unclear_dates']}"]
    for s in SHEETS:
        L.append(f"- {s}: {len(rows[s])} rows")
    L += ["- Rows are sorted by keyword; check each sheet. Photos marked TO TRANSCRIBE still need their names/amounts typed in.", ""]
    for s in SHEETS:
        L += [f"## {s}", "", "| " + " | ".join(COLUMNS) + " |", "|" + "---|" * len(COLUMNS)]
        L += ["| " + " | ".join(md_cell(v) for v in r) + " |" for r in rows[s]] or ["| (none) |" + " |" * (len(COLUMNS) - 1)]
        L.append("")
    path.write_text("\n".join(L), encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2:])
