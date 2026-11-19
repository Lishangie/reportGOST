#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборщик учебных отчётов по нормоконтролю (кафедра ПМ, РТУ МИРЭА) — версия 4.2.

    python build.py                    -> собирает report.md рядом со скриптом
    python build.py путь/report.md     -> собирает указанный файл
    python build.py --init             -> заготовка новой работы в текущей папке
    python build.py --pdf              -> дополнительно сохранить PDF
    python build.py --no-layout        -> без проверки вёрстки (быстро)

Синтаксис исходника — в ФОРМАТ.md. Оформление (шрифты, интервалы, поля,
стили) зашито в template.docx и приводится к требованиям при каждой сборке.

Что делается автоматически:
  * нумерация разделов, рисунков, таблиц, листингов, формул, приложений,
    источников (по первому упоминанию);
  * перекрёстные ссылки — настоящие поля Word (REF) без гиперссылок;
  * оглавление с номерами страниц, нумерация страниц (титул, задание,
    содержание считаются без номера, аннотация не считается);
  * формулы (LaTeX -> редактор формул Word), в том числе внутри строки: $x$;
  * «Продолжение Таблицы/Листинга N» при переносе на следующую страницу;
  * полупустые страницы внутри раздела: рисунок слегка уменьшается или
    следующий абзац переносится перед ним (в том числе длинный — он заполняет
    место и перетекает на следующую страницу);
  * вёрстка измеряется по черновикам Word с закладками-метками (3.1): закладки
    не занимают места, поэтому чистовик верстается точно как черновик;
    без Word — метки текстом через LibreOffice;
  * у формулы держатся вместе только первые строки «где» (3.1);
  * обрезка белых полей рисунков и обводка рисунков на белом фоне;
  * типографика: тире, кавычки «», неразрывные пробелы, знак минуса;
  * пунктуация списков, точки в подписях и заголовках, заглавные буквы
    в ссылках «Рисунок», «Таблица»;
  * страница-заглушка под титульный лист (форму вставляешь сам).
"""

import os
import re

import sys
import copy
import shutil
import argparse
import datetime
import subprocess
import tempfile
import zipfile
import json
import atexit
import threading
import queue

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from docx import Document
    from docx.shared import Cm, Pt, Emu, RGBColor
    from docx.enum.section import WD_SECTION, WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import OxmlElement, parse_xml
    from lxml import etree
except ImportError as e:  # pragma: no cover
    print("ОШИБКА: не установлены библиотеки:", e)
    print("Выполните:  python -m pip install python-docx pillow lxml pymupdf")
    sys.exit(1)

VERSION = "4.2"
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.docx")
YEAR = datetime.date.today().year

CM_PT = 28.3465
LINE_PT = 24.15                 # строка основного текста: 14 пт × 1,5
GAP_LINES = 5                   # допустимое пустое место внутри раздела
GAP_TOL = 0.3                   # допуск на измерение: 5,2 строки на глаз — это 5 строк
PORTRAIT_BOX = (17.0, 25.7)     # область набора, см (A4, поля 3/1/2/2)
LANDSCAPE_BOX = (25.7, 17.0)    # альбомная: поля 2/2 сбоку, 3 сверху, 1 снизу
MAX_PASSES = 40
KEEP_TABLE_ROWS = 10            # таблица до стольких строк по возможности не разрывается
KEEP_LISTING_LINES = 16         # листинг до стольких строк — тоже
TOC_TAB = 9632                  # позиция табуляции оглавления: 16,99 см (у правого поля)
PAGE_NUM_PT = 11                # номер страницы: Times New Roman, 10–11 пт

# ------------------------------------------------------------------ стили
S_TEXT = "Обычный текст ПМ"
S_H1 = "Заголовок первого уровня ПМ"
S_H2 = "Заголовок второго уровня ПМ"
S_H3 = "Заголовок третьего уровня ПМ"
S_FIG = "Рисунки ПМ"
S_FIGCAP = "Подпись рисунка ПМ"
S_TABCAP = "Подпись таблицы ПМ"
S_TABTEXT = "Текст таблицы внутри ПМ"
S_AFTERTAB = "Текст после таблицы ПМ"
S_LISTING = "Текст листинга ПМ"
S_BULLET = "Маркированный список ПМ"
S_NUMBER = "Нумерованный список ПМ"

UNNUMBERED = {"АННОТАЦИЯ", "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ",
              "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "ПРИЛОЖЕНИЯ", "ПРИЛОЖЕНИЕ"}
FRONT = {"АННОТАЦИЯ", "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"}
# буквы приложений по ГОСТ: без Ё, З, Й, О, Ч, Ъ, Ы, Ь
RU_LETTERS = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"
# буквы нумерованного списка: без е, з, й, о, ч, ъ, ы, ь
LIST_LETTERS = "абвгджиклмнпрстуфхцшщэюя"

DEFAULTS = {
    "титул": "да",                     # да (одна страница-заглушка) | нет | N — заглушка на N страниц
    "нумерация_страниц_с": "",         # заголовок; пусто — первый после содержания
    "нумерация_объектов": "авто",      # авто | по_разделам | сквозная
    "мин_источников": "0",
    "макс_источников": "0",
    "свежесть_источников": "0",        # не старше N лет; 0 — не проверять
    "разделы_с_новой_страницы": "да",
    "имя_файла": "",
    "pdf": "нет",
    "объём": "",                       # например 15-25 (страниц)
    "ширина_рисунков": "",             # ширина по умолчанию, см
    "рамка_рисунков": "авто",          # авто | да | нет
    "обрезка_рисунков": "да",
    "вёрстка": "авто",                 # авто | word | libreoffice | нет
    "обязательно": "",                 # заголовки через «;»
    "шапка_в_продолжении": "да",       # повторять шапку таблицы после «Продолжение Таблицы»
}
ALIASES = {"объем": "объём", "верстка": "вёрстка"}


def norm(t):
    return re.sub(r"\s+", " ", (t or "").strip().upper()).replace("Ё", "Е")


def is_toc_heading(up):
    return up in ("СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ")


def is_sources_heading(up):
    return bool(re.match(r"^(СПИСОК\b.*(ИСТОЧНИК|ЛИТЕРАТУР)|БИБЛИОГРАФ)", up))


def is_appendix_heading(up):
    return up in ("ПРИЛОЖЕНИЯ", "ПРИЛОЖЕНИЕ")


def is_structural(up):
    return (up in UNNUMBERED or is_sources_heading(up) or
            bool(re.match(r"^(ОПРЕДЕЛЕНИЯ|ОБОЗНАЧЕНИЯ|ТЕРМИНЫ|ПЕРЕЧЕНЬ СОКРАЩЕНИЙ)", up)))


def read_text(path):
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


# ================================================================ состояние
class State:
    def __init__(self):
        self.settings = {}
        self.warnings = []          # (строка, текст)
        self.fixes = []             # автоисправления
        self.layout_notes = []      # замечания по вёрстке (страницы)
        self.search_dirs = []
        self.seen = set()

    def setting(self, key):
        return str(self.settings.get(key, DEFAULTS[key])).strip()

    def flag(self, key):
        return self.setting(key).lower() in ("да", "yes", "1", "true", "авто")

    def warn(self, msg, line=None):
        k = (line, msg)
        if k not in self.seen:
            self.seen.add(k)
            self.warnings.append(k)

    def fix(self, msg, line=None):
        k = ("fix", line, msg)
        if k not in self.seen:
            self.seen.add(k)
            self.fixes.append((line, msg))

    def find_file(self, rel):
        if not rel:
            return None
        if os.path.isabs(rel) and os.path.exists(rel):
            return rel
        for d in self.search_dirs:
            for cand in (os.path.join(d, rel), os.path.join(d, os.path.basename(rel))):
                if os.path.isfile(cand):
                    return cand
        return None


class Block:
    def __init__(self, kind, line=0, **kw):
        self.kind = kind
        self.line = line
        self.__dict__.update(kw)

    def __repr__(self):
        return f"<{self.kind} {getattr(self, 'text', getattr(self, 'id', ''))!r}>"


# ================================================================== парсинг
KV_RE = re.compile(r"^[а-яёa-z_ ]+=", re.I)


def split_cells(s):
    """Делит строку таблицы по «|», кроме экранированных «\\|»."""
    s = s.replace("\\|", "\x00")
    return [c.strip().replace("\x00", "|") for c in s.split("|")]


def parse_kv_block(lines, i, target, S, known=None, label="настройка"):
    """Читает строки «ключ = значение» начиная с i, возвращает новый i."""
    while i < len(lines):
        ln = lines[i].strip()
        if ln and not ln.startswith("//") and not KV_RE.match(ln):
            break
        ln = re.split(r"\s//", " " + ln)[0].strip() if "//" in ln else ln
        if "=" in ln:
            k, v = ln.split("=", 1)
            k = k.strip().lower().replace(" ", "_")
            k = ALIASES.get(k, k)
            if known is not None and k not in known:
                S.warn(f"неизвестная {label} «{k}» — пропущена", i + 1)
            else:
                target[k] = v.strip()
        i += 1
    return i


def parse(src, S):
    lines = src.split("\n")
    blocks, sources = [], {}
    buf, buf_line = [], [0]
    i = 0

    def flush():
        if buf:
            blocks.append(Block("p", buf_line[0], text=" ".join(buf).strip()))
            buf.clear()

    while i < len(lines):
        raw = lines[i].rstrip()
        s = raw.strip()
        ln = i + 1

        if s.startswith("[настройки]"):
            flush()
            i = parse_kv_block(lines, i + 1, S.settings, S, DEFAULTS)
            continue
        if s.startswith("[источники]"):
            flush()
            i += 1
            while i < len(lines):
                t = lines[i].strip()
                if t and not re.match(r"^[\w\-.]+\s*=", t):
                    break
                if "=" in t:
                    k, v = t.split("=", 1)
                    sources[k.strip()] = (v.strip(), i + 1)
                i += 1
            continue

        if not s:
            flush()
            i += 1
            continue

        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            flush()
            level = len(m.group(1))
            if level > 3:
                S.warn("заголовки глубже третьего уровня не допускаются — понижен до ###", ln)
                level = 3
            text = m.group(2).strip()
            forced = bool(re.search(r"\{-\}\s*$", text))
            text = re.sub(r"\s*\{-\}\s*$", "", text)
            text, cut = strip_end_period(text)
            if cut:
                S.fix(f"убрана точка в конце заголовка «{text}»", ln)
            blocks.append(Block("h", ln, level=level, text=text, forced_unnumbered=forced))
            i += 1
            continue

        m = re.match(r"^\[рис:\s*(.*?)\]$", s)
        if m:
            flush()
            parts = split_cells(m.group(1)) + [""] * 4
            try:
                width = float(parts[3].replace(",", ".")) if parts[3] else 0
            except ValueError:
                width = 0
                S.warn(f"рисунок «{parts[0]}»: ширина «{parts[3]}» не число", ln)
            blocks.append(Block("fig", ln, id=parts[0], path=parts[1],
                                caption=parts[2], width=width))
            i += 1
            continue

        m = re.match(r"^\[табл:\s*(.*?)\]$", s)
        if m:
            flush()
            parts = split_cells(m.group(1)) + [""]
            rows = []
            i += 1
            while i < len(lines) and lines[i].strip().startswith("|"):
                r = lines[i].strip()
                r = r[1:] if r.startswith("|") else r
                r = r[:-1] if r.endswith("|") and not r.endswith("\\|") else r
                if not re.match(r"^[\s\-:|]+$", r):
                    rows.append(split_cells(r))
                i += 1
            blocks.append(Block("table", ln, id=parts[0], caption=parts[1], rows=rows))
            continue

        m = re.match(r"^\[листинг:\s*(.*?)\]$", s)
        if m:
            flush()
            parts = split_cells(m.group(1)) + [""] * 4
            code = None
            i += 1
            # код можно дать прямо в report.md блоком ``` … ```
            j = i
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and lines[j].strip().startswith("```"):
                k = j + 1
                body = []
                while k < len(lines) and not lines[k].strip().startswith("```"):
                    body.append(lines[k].rstrip("\r"))
                    k += 1
                code = "\n".join(body)
                i = k + 1
            if code is None:
                found = S.find_file(parts[1])
                if found:
                    code = read_text(found)
                else:
                    if parts[1]:
                        S.warn(f"листинг «{parts[0]}»: файл «{parts[1]}» не найден", ln)
                    else:
                        S.warn(f"листинг «{parts[0]}»: нет ни файла, ни блока ``` с кодом", ln)
                    code = "# вставьте сюда код"
            numbered = parts[3].strip().lower() in ("строки", "нумерация", "да")
            blocks.append(Block("listing", ln, id=parts[0], caption=parts[2],
                                code=code, line_numbers=numbered))
            continue

        m = re.match(r"^\[формула:\s*(.*?)\]$", s)
        if m:
            flush()
            parts = split_cells(m.group(1)) + [""] * 3
            blocks.append(Block("formula", ln, id=parts[0], tex=parts[1], where=parts[2]))
            i += 1
            continue

        m = re.match(r"^\[приложение:\s*(.*?)\]$", s)
        if m:
            flush()
            parts = split_cells(m.group(1)) + [""] * 2
            blocks.append(Block("appendix", ln, letter=parts[0].upper(), title=parts[1]))
            i += 1
            continue

        if s in ("[альбом]", "[/альбом]"):
            flush()
            blocks.append(Block("land_on" if s == "[альбом]" else "land_off", ln))
            i += 1
            continue

        if s == "[разрыв]":
            flush()
            blocks.append(Block("pagebreak", ln))
            i += 1
            continue

        indent = len(raw) - len(raw.lstrip(" \t").replace("\t", "  "))
        lvl_ind = min(2, indent // 2)
        m = re.match(r"^[-*•]\s+(.*)$", s)
        if m:
            flush()
            blocks.append(Block("li", ln, text=m.group(1), ordered=None, level=lvl_ind))
            i += 1
            continue
        m = re.match(r"^(\d+(?:\.\d+){0,2})[.)]\s+(.*)$", s)
        if m:
            flush()
            dots = m.group(1).count(".")
            blocks.append(Block("li", ln, text=m.group(2), ordered="num",
                                level=max(dots, lvl_ind)))
            i += 1
            continue
        m = re.match(r"^([а-яё])\)\s+(.*)$", s)
        if m:
            flush()
            blocks.append(Block("li", ln, text=m.group(2), ordered="let", level=lvl_ind))
            i += 1
            continue

        # продолжение пункта списка без пустой строки
        if not buf and blocks and blocks[-1].kind == "li" and i > 0 and lines[i - 1].strip():
            blocks[-1].text += " " + s
            i += 1
            continue

        if not buf:
            buf_line[0] = ln
        buf.append(s)
        i += 1

    flush()
    # заголовок сразу перед [альбом] переносим на альбомную страницу
    k = 0
    while k < len(blocks) - 1:
        if blocks[k].kind == "h" and blocks[k + 1].kind == "land_on":
            blocks[k], blocks[k + 1] = blocks[k + 1], blocks[k]
            k += 2
        else:
            k += 1
    return blocks, sources


def apply_moves(blocks, moves):
    """Переносит несколько абзацев, идущих после рисунка или альбома, перед ним —
    чтобы не оставалось полупустой страницы (решение вёрстки)."""
    if not moves:
        return blocks
    for key, k in moves.items():
        idx = None
        for n, b in enumerate(blocks):
            if obj_key(b, blocks, n) == key:
                idx = n
                break
        if idx is None:
            continue
        end = idx + 1
        if blocks[idx].kind == "land_on":
            while end < len(blocks) and blocks[end].kind != "land_off":
                end += 1
            end += 1
        taken = []
        j = end
        while j < len(blocks) and len(taken) < k and blocks[j].kind == "p":
            taken.append(j)
            j += 1
        if not taken:
            continue
        moved = [blocks[t] for t in taken]
        rest = [b for t, b in enumerate(blocks) if t not in taken]
        pos = rest.index(blocks[idx])
        blocks = rest[:pos] + moved + rest[pos:]
    return blocks


def obj_key(b, blocks=None, n=None):
    if b.kind == "fig":
        return "fig:" + (b.id or str(b.line))
    if b.kind == "land_on":
        return "land:" + str(b.line)
    if b.kind in ("table", "listing"):
        return b.kind + ":" + (b.id or str(b.line))
    return None


# ============================================================ типографика
UNITS = (r"(?:мм|см|дм|км|м|кг|мг|г|т|мс|мкс|с|мин|ч|сут|%|‰|руб\.|₽|шт\.|пт|бит|байт|"
         r"Кб|Мб|Гб|Тб|КБ|МБ|ГБ|ТБ|тыс\.|млн|млрд|°C|°|Гц|кГц|МГц|ГГц|px|dpi|В|А|Вт|кВт)")


def typography(t):
    """Тире, кавычки, минус, неразрывные пробелы — для обычного текста
    (код, формулы и адреса сюда не попадают)."""
    if not t:
        return t
    t = t.replace("\u2212", "–")
    t = t.replace("...", "…")
    t = re.sub(r"\s--\s", " — ", t)
    t = t.replace("--", "—")
    t = re.sub(r"(?<=\S) - (?=\S)", " — ", t)
    t = re.sub(r"^- (?=\S)", "— ", t)
    # минус перед числом: «U(-5; 5)» -> «U(–5; 5)»
    t = re.sub(r"(?:(?<=^)|(?<=[\s(\[;=:<>]))-(?=\d)", "–", t)
    t = re.sub(r'"([^"]*)"', "«\\1»", t)
    t = re.sub(r"(?<=[^\s(«]) +(?=[,;:!?»)](?:\s|$|[,.;:!?»)]))", "", t)   # «слово ,» -> «слово,»
    t = re.sub(r"(?<=[а-яёА-ЯЁa-zA-Z0-9»)]) +\.(?=\s|$)", ".", t)            # «слово .» -> «слово.»
    t = re.sub(r"(?<=[(«]) +(?=\S)", "", t)                                 # «( текст» -> «(текст»
    t = re.sub(r"(\d)\s+(" + UNITS + r")(?![а-яёА-ЯЁa-zA-Z])", "\\1\u00a0\\2", t)
    t = re.sub(r"(?<=\d) (?=\d{3}(?!\d))", "\u00a0", t)
    t = re.sub(r"(?<![\w\-])([а-яёА-ЯЁ]{1,2}) (?=\S)", "\\1\u00a0", t)
    t = re.sub(r"(?<=\S) —", "\u00a0—", t)
    t = t.replace("№ ", "№\u00a0")
    t = re.sub(r"(?<=\b[А-ЯЁ]\.) (?=[А-ЯЁ])", "\u00a0", t)
    t = re.sub(r"(?<=[А-ЯЁ][а-яё]) (?=[А-ЯЁ]\.\s?[А-ЯЁ]\.)", "\u00a0", t)
    t = re.sub(r"\bт\. (?=[едп]\.)", "т.\u00a0", t)
    t = re.sub(r"[ \t]{2,}", " ", t)
    return t


ESCAPES = {"\\*": "\ue010", "\\$": "\ue011", "\\{": "\ue012", "\\}": "\ue013",
           "\\`": "\ue014", "\\_": "_"}
UNESC = {"\ue010": "*", "\ue011": "$", "\ue012": "{", "\ue013": "}", "\ue014": "`"}

TOKEN_RE = re.compile(
    r"(`[^`]+`)"                                         # код
    r"|(\$\$[^$]+\$\$|\$[^$\s][^$]*?\$)"                 # формула
    r"|(\{(?:рис|табл|лист|форм|прил|ист):[^{}]+\})"     # ссылка
    r"|(https?://[^\s«»<>\"]+)"                          # адрес
)
REF_KIND = {"рис": "fig", "табл": "table", "лист": "listing", "форм": "formula",
            "прил": "app"}


def unesc(t):
    for k, v in UNESC.items():
        t = t.replace(k, v)
    return t


def escape(t):
    for k, v in ESCAPES.items():
        t = t.replace(k, v)
    return t


def tokenize(text):
    """-> [(тип, значение)], тип: t, code, math, ref, url."""
    text = escape(text or "")
    out, pos = [], 0
    for m in TOKEN_RE.finditer(text):
        if m.start() > pos:
            out.append(("t", text[pos:m.start()]))
        g = m.group(0)
        if m.group(1):
            out.append(("code", unesc(g[1:-1])))
        elif m.group(2):
            tex = g.strip("$")
            out.append(("math", unesc(tex).strip()))
        elif m.group(3):
            out.append(("ref", g[1:-1]))
        else:
            url = g
            tail = ""
            while url and url[-1] in ".,;:)]»":
                tail = url[-1] + tail
                url = url[:-1]
            out.append(("url", url))
            if tail:
                out.append(("t", tail))
        pos = m.end()
    if pos < len(text):
        out.append(("t", text[pos:]))
    return out


def plain_text(text):
    """Только обычный текст без кода, формул, ссылок и адресов (для проверок)."""
    return "".join(v if k == "t" else " ¤ " for k, v in tokenize(text))


def collect_math(text, acc):
    for k, v in tokenize(text):
        if k == "math":
            acc.add((v, False))


OBJ_WORDS = {"рис": "Рисун", "табл": "Таблиц", "лист": "Листинг", "прил": "Приложени"}


def prepare_text(text, S, line):
    """Автоисправления в тексте до разбора: заглавная буква в «Рисунок 1.1»,
    скобки вокруг номера формулы, лишние скобки вокруг ссылок на источник."""
    def cap(m):
        w = m.group(1)
        if w[0].islower():
            S.fix(f"«{w}» перед ссылкой — с прописной буквы", line)
            return w[0].upper() + w[1:] + m.group(2)
        return m.group(0)
    text = re.sub(r"(?<![\w])((?:рисун|таблиц|листинг|приложени)[а-яё]*)(\s+)(?=\{(?:рис|табл|лист|прил):)",
                  cap, text, flags=re.I)
    text = re.sub(r"(?<!\()\{форм:([^}]+)\}(?!\))", r"({форм:\1})", text)
    text = re.sub(r"\[\s*\{ист:([^}]+)\}\s*\]", r"{ист:\1}", text)
    return text


# ============================================================== проверки
def lint_text(text, S, line, where=""):
    pt = plain_text(text)
    if re.search(r"(?<![\w])(рис|табл)\.\s*\d", pt, re.I) or re.search(r"(?<![\w])(рис|табл)\.\s*¤", pt, re.I):
        S.warn("сокращения «рис.»/«табл.» не допускаются — пишите «Рисунок», «Таблица»", line)
    if re.search(r"(рисун\w*|таблиц\w*|формул\w*|листинг\w*|приложени\w*)\s+(ниже|выше)\b", pt, re.I) or \
            re.search(r"\b(ниже|выше)\s+(на\s+рисунке|в\s+таблице|приведен\w*|представлен\w*|показан\w*)", pt, re.I) or \
            re.search(r"на следующей странице", pt, re.I):
        S.warn("ссылки вида «ниже/выше/на следующей странице» не допускаются — укажите номер объекта", line)
    for m in re.finditer(r"(?<![\d.,\w])(\d+\.\d+)(?![\d.\w])", pt):
        num = m.group(1)
        if num.startswith("0.") or len(num.split(".")[1]) >= 3:
            S.warn(f"десятичная дробь «{num}» — в русском тексте пишется через запятую", line)
            break
    m = re.search(r"\b\d+-(ти|х|ми|го|му|ым|ой|й|я|ом)\b", pt)
    if m:
        S.warn(f"«{m.group(0)}» — числительные пишутся без падежных окончаний", line)
    for m in re.finditer(r"(?:^|\s)([=<>+≤≥≠])(?:\s|$)", pt):
        ctx = pt[max(0, m.start() - 3): m.end() + 3]
        if not re.search(r"\d", ctx):
            S.warn(f"знак «{m.group(1)}» в тексте без чисел — пишите словами («меньше», «равно») "
                   f"или оформите формулой $…$", line)
            break
    raw = text or ""
    for m in re.finditer(r"\{(рис|табл|лист|прил):([^}]+)\}", raw):
        before = raw[max(0, m.start() - 60):m.start()]
        if not re.search(OBJ_WORDS[m.group(1)], before, re.I):
            S.warn(f"перед ссылкой {{{m.group(1)}:{m.group(2)}}} нет слова "
                   f"«{OBJ_WORDS[m.group(1)]}…» — ссылка пишется полностью: «на Рисунке 1.2»", line)
    if re.search(r"[.!?]\s*\{ист:", raw):
        S.warn("ссылка на источник ставится до точки: «… текст [1].»", line)


def normalize_lists(blocks, S):
    """Пунктуация списков по требованиям: маркированный — со строчной, «;»,
    последний пункт «.»; нумерованный — с прописной и с точкой."""
    n = len(blocks)
    k = 0
    while k < n:
        if blocks[k].kind != "li":
            k += 1
            continue
        j = k
        while j < n and blocks[j].kind == "li":
            j += 1
        group = blocks[k:j]
        for gi, b in enumerate(group):
            t = b.text.rstrip()
            nxt = group[gi + 1] if gi + 1 < len(group) else None
            if b.ordered == "num":
                if re.match(r"[а-яё]", t[:1]):
                    t = t[0].upper() + t[1:]
                    S.fix("пункт нумерованного списка начат с прописной буквы", b.line)
                if not t.endswith((":", ".")):
                    t = re.sub(r"[;,]+$", "", t) + "."
                    S.fix("пункт нумерованного списка закончен точкой", b.line)
            else:
                if re.match(r"[А-ЯЁ][а-яё]", t):
                    t = t[0].lower() + t[1:]
                    S.fix("пункт маркированного списка начат со строчной буквы "
                          "(проверьте, не имя ли собственное)", b.line)
                if not t.endswith(":"):
                    last = nxt is None or (nxt.ordered == "num" and b.ordered != "num")
                    want = "." if last else ";"
                    cur = re.sub(r"[;,.]+$", "", t)
                    if t != cur + want:
                        S.fix(f"пункт списка закончен «{want}»", b.line)
                    t = cur + want
                body = plain_text(t[:-1])
                if re.search(r"[а-яёa-z0-9)»][.!?]\s+[А-ЯЁ][а-яё]", body):
                    S.warn("в пункте маркированного списка несколько предложений — "
                           "так нельзя, разбейте на пункты", b.line)
            b.text = t
        k = j


# =============================================================== нумерация
OBJ_KINDS = ("fig", "table", "listing", "formula")


def mark_headings(blocks):
    appendix = False
    parent_numbered = False
    sec = sub = subsub = 0
    for b in blocks:
        if b.kind != "h":
            continue
        up = norm(b.text)
        if is_appendix_heading(up):
            appendix = True
        if b.level == 1:
            b.numbered = not (b.forced_unnumbered or is_structural(up) or appendix)
            parent_numbered = b.numbered
            if b.numbered:
                sec += 1
                sub = subsub = 0
                b.num = str(sec)
        else:
            b.numbered = parent_numbered and not b.forced_unnumbered and not appendix
            if b.numbered and b.level == 2:
                sub += 1
                subsub = 0
                b.num = f"{sec}.{sub}"
            elif b.numbered and b.level == 3:
                subsub += 1
                b.num = f"{sec}.{sub}.{subsub}"
        if not b.numbered:
            b.num = ""


def choose_mode(blocks, S):
    mode = S.setting("нумерация_объектов").lower()
    if mode in ("по_разделам", "сквозная"):
        return mode
    in_numbered = appendix = has_numbered = False
    for b in blocks:
        if b.kind == "h" and b.level == 1:
            if is_appendix_heading(norm(b.text)):
                appendix = True
            in_numbered = b.numbered
            has_numbered = has_numbered or b.numbered
        elif b.kind == "appendix":
            appendix = True
        elif b.kind in OBJ_KINDS and not appendix and not in_numbered:
            if b.kind == "formula" and not getattr(b, "numbered_formula", True):
                continue
            return "сквозная"
    return "по_разделам" if has_numbered else "сквозная"


def all_texts(b):
    if b.kind in ("p", "li", "h"):
        yield b.text
    elif b.kind == "fig":
        yield b.caption
    elif b.kind in ("table",):
        yield b.caption
        for r in b.rows:
            for c in r:
                yield c
    elif b.kind == "listing":
        yield b.caption
    elif b.kind == "formula":
        yield b.where
    elif b.kind == "appendix":
        yield b.title


def enumerate_all(blocks, sources, S):
    mark_headings(blocks)
    referenced = set()
    for b in blocks:
        for t in all_texts(b):
            for m in re.finditer(r"\{(рис|табл|лист|форм|прил):([^}]+)\}", t or ""):
                referenced.add((REF_KIND[m.group(1)], m.group(2).strip()))
    for b in blocks:
        if b.kind == "formula":
            b.numbered_formula = bool(b.id) and ("formula", b.id) in referenced
            if b.id and not b.numbered_formula:
                S.warn(f"на формулу «{b.id}» нет ссылки — номер не ставится "
                       f"(нумеруются только формулы, на которые ссылаются)", b.line)
    mode = choose_mode(blocks, S)

    sec = 0
    counters = {k: 0 for k in OBJ_KINDS}
    labels = {}
    appendix_mode = False
    cur_letter = None
    src_order, src_no = [], {}
    app_letters = []

    def scan_sources(text, line):
        for key in re.findall(r"\{ист:([^}]+)\}", text or ""):
            for k in [x.strip() for x in key.split(",")]:
                if not k or re.match(r"^(с|c|p|pp|стр)\.", k, re.I):
                    continue
                if k not in src_no:
                    src_no[k] = len(src_order) + 1
                    src_order.append(k)

    for b in blocks:
        if b.kind == "h":
            if is_appendix_heading(norm(b.text)):
                appendix_mode = True
                if norm(b.text) == "ПРИЛОЖЕНИЕ":
                    cur_letter = "А"
                    counters = {k: 0 for k in counters}
            if b.level == 1 and b.numbered:
                sec += 1
                if mode == "по_разделам":
                    counters = {k: 0 for k in counters}
            scan_sources(b.text, b.line)
        elif b.kind == "appendix":
            appendix_mode = True
            cur_letter = b.letter
            app_letters.append((b.letter, b.line))
            labels[("app", b.letter)] = b.letter
            counters = {k: 0 for k in counters}
            scan_sources(b.title, b.line)
        elif b.kind in counters:
            if b.kind == "formula" and not b.numbered_formula:
                b.number = ""
                scan_sources(b.where, b.line)
                continue
            counters[b.kind] += 1
            n = counters[b.kind]
            if appendix_mode and cur_letter:
                b.number = f"{cur_letter}.{n}"
            elif mode == "по_разделам":
                if sec == 0:
                    S.warn(f"объект «{b.id}» стоит вне нумерованного раздела — получит "
                           f"номер 0.{n}; поставьте нумерация_объектов = сквозная", b.line)
                b.number = f"{sec}.{n}"
            else:
                b.number = str(n)
            labels[(b.kind, b.id)] = b.number
            for t in all_texts(b):
                scan_sources(t, b.line)
            continue
        if b.kind in ("p", "li"):
            scan_sources(b.text, b.line)

    # приложения: буквы, порядок, ссылки
    expected = list(RU_LETTERS)
    for n, (letter, line) in enumerate(app_letters):
        if letter not in RU_LETTERS:
            S.warn(f"буква приложения «{letter}» недопустима (без Ё, З, Й, О, Ч, Ъ, Ы, Ь)", line)
        elif n < len(expected) and letter != expected[n]:
            S.warn(f"приложение «{letter}» должно обозначаться «{expected[n]}» — "
                   f"буквы идут по порядку", line)
    if app_letters:
        first_ref = {}
        for b in blocks:
            if b.kind in ("p", "li"):
                for m in re.finditer(r"\{прил:([^}]+)\}", b.text):
                    first_ref.setdefault(m.group(1).strip().upper(), len(first_ref))
        for letter, line in app_letters:
            if letter not in first_ref:
                S.warn(f"на Приложение {letter} нет ссылки в тексте ({{прил:{letter}}})", line)
        order = sorted([l for l, _ in app_letters if l in first_ref], key=lambda l: first_ref[l])
        if order != [l for l, _ in app_letters if l in first_ref]:
            S.warn("приложения должны идти в порядке ссылок на них в тексте")

    # источники
    for k, (desc, line) in sources.items():
        if k not in src_no:
            S.warn(f"источник «{k}» описан, но не упомянут в тексте — так нельзя", line)
    for k in src_order:
        if k not in sources:
            S.warn(f"на источник «{k}» есть ссылка, но нет описания в блоке [источники]")
    has_src_heading = any(b.kind == "h" and is_sources_heading(norm(b.text)) for b in blocks)
    if src_order and not has_src_heading:
        S.warn("в тексте есть ссылки на источники, но нет раздела «СПИСОК ИСТОЧНИКОВ»")
    if has_src_heading:
        need = int(re.sub(r"\D", "", S.setting("мин_источников")) or 0)
        top = int(re.sub(r"\D", "", S.setting("макс_источников")) or 0)
        if not src_order:
            S.warn("раздел со списком источников есть, но ссылок {ист:…} в тексте нет")
        elif need and len(src_order) < need:
            S.warn(f"источников {len(src_order)} — нужно не менее {need}")
        if top and len(src_order) > top:
            S.warn(f"источников {len(src_order)} — по требованиям не более {top}")
        fresh = int(re.sub(r"\D", "", S.setting("свежесть_источников")) or 0)
        for k in src_order:
            if k not in sources:
                continue
            desc, line = sources[k]
            years = [int(y) for y in re.findall(r"(?<!\d)(19\d\d|20\d\d)(?!\d)", desc)]
            if fresh and years and not re.match(r"^\s*ГОСТ", desc) and \
                    max(years) < YEAR - fresh:
                S.warn(f"источник «{k}» {max(years)} года — старше {fresh} лет", line)
            if re.search(r"https?://", desc) and not re.search(r"дата обращения", desc, re.I):
                S.warn(f"источник «{k}»: у электронного ресурса укажите «(дата обращения: ДД.ММ.ГГГГ)»", line)
    return labels, src_order, src_no, mode


# ===================================================== общие проверки блоков
def lint_blocks(blocks, labels, S):
    # ссылки до объекта, расстояние от первого упоминания
    first_mention = {}
    text_pos = 0
    positions = {}
    for idx, b in enumerate(blocks):
        if b.kind in ("p", "li"):
            text_pos += 1
            for m in re.finditer(r"\{(рис|табл|лист|форм|прил):([^}]+)\}", b.text):
                key = (REF_KIND[m.group(1)], m.group(2).strip())
                first_mention.setdefault(key, (text_pos, idx))
                if key[0] != "app" and key not in labels and key[0] != "formula":
                    S.warn(f"ссылка {{{m.group(1)}:{m.group(2)}}} — объект не найден", b.line)
                if key[0] == "formula" and not any(x.kind == "formula" and x.id == key[1] for x in blocks):
                    S.warn(f"ссылка {{форм:{m.group(2)}}} — формула не найдена", b.line)
        elif b.kind in ("fig", "table", "listing", "formula") and b.id:
            positions[(b.kind, b.id)] = (text_pos, idx)
    # диапазон «Рисунки {рис:a}–{рис:c}» упоминает и все рисунки между ними
    for idx, b in enumerate(blocks):
        if b.kind not in ("p", "li"):
            continue
        for m in re.finditer(r"\{(рис|табл):([^}]+)\}\s*[–—-]\s*\{\1:([^}]+)\}", b.text):
            kind = REF_KIND[m.group(1)]
            seq = [x.id for x in blocks if x.kind == kind]
            a, z = m.group(2).strip(), m.group(3).strip()
            if a in seq and z in seq:
                for oid in seq[seq.index(a):seq.index(z) + 1]:
                    first_mention.setdefault((kind, oid), first_mention.get((kind, a), (0, idx)))
    land = False
    for idx, b in enumerate(blocks):
        if b.kind == "land_on":
            land = True
        elif b.kind == "land_off":
            land = False
        if b.kind in ("fig", "table", "listing"):
            word = {"fig": "Рисунок", "table": "Таблицу", "listing": "Листинг"}[b.kind]
            tag = {"fig": "рис", "table": "табл", "listing": "лист"}[b.kind]
            key = (b.kind, b.id)
            fm = first_mention.get(key)
            if not fm or fm[1] > idx:
                S.warn(f"на {word} {getattr(b, 'number', '')} нет ссылки в тексте ДО него "
                       f"(добавьте «{{{tag}:{b.id}}}» в абзац перед ним)", b.line)
            elif not land and fm and positions.get(key):
                gap = positions[key][0] - fm[0]
                if gap > 2:
                    S.warn(f"{word} {b.number} стоит далеко от первой ссылки "
                           f"({gap} абзаца между ними) — размещайте сразу после абзаца "
                           f"с первым упоминанием", b.line)

    # структура вокруг объектов и списков
    for idx, b in enumerate(blocks):
        nxt = blocks[idx + 1] if idx + 1 < len(blocks) else None
        if b.kind == "li" and nxt is not None and nxt.kind in ("fig", "table", "listing") and \
                idx + 2 < len(blocks) and blocks[idx + 2].kind == "li":
            S.warn("внутри списка не может быть рисунков и таблиц", nxt.line)
        if b.kind == "li" and (nxt is None or nxt.kind in ("h", "appendix")):
            S.warn("раздел не может заканчиваться списком — добавьте после него поясняющее предложение",
                   b.line)

    # раздел из одного подраздела
    for lvl in (2, 3):
        count, parent = 0, None
        for b in blocks + [Block("h", 0, level=1, text="", numbered=False, num="")]:
            if b.kind != "h":
                continue
            if b.level < lvl:
                if parent is not None and count == 1 and parent.numbered:
                    S.warn(f"в разделе «{parent.text}» один подраздел уровня {lvl} — "
                           f"подразделов должно быть не меньше двух (или уберите деление)",
                           parent.line)
                parent, count = b, 0
            elif b.level == lvl:
                count += 1

    # обязательные заголовки
    req = [x.strip() for x in S.setting("обязательно").split(";") if x.strip()]
    heads = [norm(b.text) for b in blocks if b.kind == "h"]
    for r in req:
        if norm(r) not in heads:
            S.warn(f"по требованиям предмета нужен раздел «{r}» — его нет")

    # введение и заключение без ссылок на источники
    cur = None
    for b in blocks:
        if b.kind == "h" and b.level == 1:
            cur = norm(b.text)
        elif b.kind in ("p", "li") and cur in ("ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ") and "{ист:" in b.text:
            S.warn(f"во {'введении' if cur == 'ВВЕДЕНИЕ' else 'заключении'} ссылки на источники "
                   f"не ставятся", b.line)

    for b in blocks:
        for t in all_texts(b):
            if t:
                lint_text(t, S, b.line)


# ============================================================ формулы (OMML)
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


ABBR_END = re.compile(r"(?:\b(?:руб|тыс|млн|млрд|шт|ед|др|пр|см|г|гг|в|вв|с|стр|экз|т\.\s?[дпе]|и\.\s?о)\.|\b[а-яa-z]{1,3}\.)$")


def strip_end_period(text):
    """Точка в конце подписи/заголовка не ставится — но сокращение «руб.» остаётся."""
    t = text.rstrip()
    if t.endswith(".") and not t.endswith("..") and not ABBR_END.search(t):
        return t[:-1].rstrip(), True
    return t, False


def split_where(s):
    """Делит пояснения «где» по «;» вне формул $…$."""
    out, cur, inmath = [], "", False
    for ch in s or "":
        if ch == "$":
            inmath = not inmath
        if ch == ";" and not inmath:
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur.strip())
    return [x for x in out if x]


def prepare_where(where):
    """«x — смысл» -> «$x$ — смысл»: обозначения в пояснениях тоже формулами."""
    out = []
    for part in split_where(where):
        m = re.match(r"^\s*(.+?)\s+[—–-]\s+(.+)$", part)
        if m and "$" not in m.group(1) and len(m.group(1)) <= 40 and \
                not re.search(r"[а-яёА-ЯЁ]{2,}", m.group(1)):
            part = "$" + m.group(1).strip() + "$ — " + m.group(2)
        out.append(part)
    return "; ".join(out)


def upper_plain(text):
    """Прописные буквы только в обычном тексте (формулы и ссылки не трогаем)."""
    return "".join(v.upper() if k == "t" else _retoken(k, v) for k, v in tokenize(text))


def _retoken(kind, val):
    if kind == "code":
        return "`" + val + "`"
    if kind == "math":
        return "$" + val + "$"
    if kind == "ref":
        return "{" + val + "}"
    return val


def find_pandoc():
    p = shutil.which("pandoc")
    if p:
        return p
    try:                                    # pip install pypandoc_binary — pandoc внутри пакета
        import pypandoc
        p = pypandoc.get_pandoc_path()
        if p and os.path.exists(p):
            return p
    except Exception:
        pass
    for c in (os.path.join(os.environ.get("LOCALAPPDATA", ""), "Pandoc", "pandoc.exe"),
              r"C:\Program Files\Pandoc\pandoc.exe",
              r"C:\Program Files (x86)\Pandoc\pandoc.exe",
              "/usr/local/bin/pandoc", "/opt/homebrew/bin/pandoc"):
        if c and os.path.exists(c):
            return c
    return None


GREEK = {
    "\\alpha": "α", "\\beta": "β", "\\gamma": "γ", "\\delta": "δ",
    "\\epsilon": "ε", "\\varepsilon": "ε", "\\zeta": "ζ", "\\eta": "η",
    "\\theta": "θ", "\\lambda": "λ", "\\mu": "μ", "\\nu": "ν",
    "\\xi": "ξ", "\\pi": "π", "\\rho": "ρ", "\\sigma": "σ",
    "\\tau": "τ", "\\phi": "φ", "\\chi": "χ", "\\psi": "ψ",
    "\\omega": "ω", "\\Delta": "Δ", "\\Sigma": "Σ", "\\Omega": "Ω",
    "\\times": "×", "\\cdot": "·", "\\leq": "≤", "\\le": "≤",
    "\\geq": "≥", "\\ge": "≥", "\\neq": "≠", "\\infty": "∞", "\\sum": "Σ",
    "\\approx": "≈", "\\pm": "±",
}


def delatex(t):
    for k, v in sorted(GREEK.items(), key=lambda x: -len(x[0])):
        t = t.replace(k, v)
    return re.sub(r"\\[a-zA-Z]+", "", t).replace("{", "").replace("}", "")


def merge_decimal_commas(root):
    """«0», «,», «5681» отдельными элементами формулы -> одно число «0,5681»,
    чтобы после десятичной запятой не появлялся пробел."""
    M = "{%s}" % MATH_NS
    for parent in list(root.iter()):
        kids = list(parent)
        i = 0
        while i + 2 < len(kids):
            a, c, b = kids[i], kids[i + 1], kids[i + 2]
            if all(k.tag == M + "r" for k in (a, c, b)):
                ta, tc, tb = (k.find(M + "t") for k in (a, c, b))
                if ta is not None and tc is not None and tb is not None and \
                        re.fullmatch(r"\d+(,\d+)*", ta.text or "") and tc.text == "," and \
                        re.match(r"\d", tb.text or ""):
                    m = re.match(r"(\d+)(.*)$", tb.text)
                    ta.text = ta.text + "," + m.group(1)
                    parent.remove(c)
                    if m.group(2):
                        tb.text = m.group(2)
                        kids = list(parent)
                    else:
                        parent.remove(b)
                        kids = list(parent)
                    continue
            i += 1


class MathCache:
    def __init__(self, S):
        self.S = S
        self.cache = {}
        self.pandoc = find_pandoc()
        self.failed_note = False
        self.failed = set()

    def prepare(self, items):
        todo = [it for it in sorted(items) if it not in self.cache and it[0].strip()]
        if not todo:
            return
        if not self.pandoc:
            for it in todo:
                self.cache[it] = None
            return
        res = self._convert(todo)
        if res is None or len(res) != len(todo):
            res = []
            for it in todo:
                r = self._convert([it])
                res.append(r[0] if r else None)
        for it, el in zip(todo, res):
            self.cache[it] = el
            if el is None:
                self.failed.add(it[0])

    def _convert(self, items):
        try:
            with tempfile.TemporaryDirectory() as td:
                md = os.path.join(td, "f.md")
                dx = os.path.join(td, "f.docx")
                parts = []
                for tex, disp in items:
                    t = tex.strip().replace("\n", " ")
                    parts.append(("$$%s$$" if disp else "$%s$") % t)
                with open(md, "w", encoding="utf-8") as f:
                    f.write("\n\n".join(parts) + "\n")
                subprocess.run([self.pandoc, md, "-f", "markdown", "-o", dx],
                               check=True, capture_output=True, timeout=120)
                with zipfile.ZipFile(dx) as z:
                    xml = z.read("word/document.xml")
            root = etree.fromstring(xml)
            # pandoc пишет begChr, endChr, sepChr — по схеме нужен иной порядок
            order = ["begChr", "sepChr", "endChr", "grow", "shp", "ctrlPr"]
            for dpr in root.iter("{%s}dPr" % MATH_NS):
                kids = sorted(list(dpr), key=lambda e: order.index(etree.QName(e).localname)
                              if etree.QName(e).localname in order else 99)
                for k in kids:
                    dpr.remove(k)
                for k in kids:
                    dpr.append(k)
            merge_decimal_commas(root)
            body = root.find(qn("w:body"))
            out = []
            for p in body.findall(qn("w:p")):
                el = p.find("{%s}oMathPara" % MATH_NS)
                if el is None:
                    el = p.find("{%s}oMath" % MATH_NS)
                out.append(el)
            return out
        except Exception:
            return None

    def emit(self, S):
        if not self.pandoc and self.cache:
            S.warn("не найден pandoc — формулы вставлены текстом. Выполните "
                   "«python -m pip install pypandoc_binary» (или установите pandoc с "
                   "https://pandoc.org/installing.html), иначе формулы будут не в редакторе формул")
        for t in sorted(self.failed):
            S.warn(f"формулу «{t}» не удалось разобрать — проверьте LaTeX")

    def get(self, tex, display):
        return self.cache.get((tex, display))


# ======================================================== низкоуровневый XML
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
             "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
             "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
             "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
             "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
              "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
              "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
              "printerSettings", "sectPrChange"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
               "shd", "tblLayout", "tblCellMar", "tblLook"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
             "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
             "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
             "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em",
             "lang", "eastAsianLayout", "specVanish", "oMath"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
              "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]
TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
              "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"]


def put(parent, tag, attrs=None, order=PPR_ORDER, children=None):
    """Ставит дочерний элемент w:tag в правильное по схеме место (с заменой)."""
    name = tag.split(":")[-1]
    for old in parent.findall(qn("w:" + name)):
        parent.remove(old)
    el = OxmlElement("w:" + name)
    for k, v in (attrs or {}).items():
        el.set(qn("w:" + k), str(v))
    for ch in children or []:
        el.append(ch)
    idx = order.index(name) if name in order else len(order)
    for i, child in enumerate(parent):
        cname = child.tag.split("}")[-1]
        if cname in order and order.index(cname) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def drop(parent, tag):
    for old in parent.findall(qn("w:" + tag)):
        parent.remove(old)


def ppr(p):
    return p._p.get_or_add_pPr()


def set_numpr(p, num_id, ilvl=0):
    el = put(ppr(p), "numPr")
    a = OxmlElement("w:ilvl"); a.set(qn("w:val"), str(ilvl))
    b = OxmlElement("w:numId"); b.set(qn("w:val"), str(num_id))
    el.append(a); el.append(b)


def kill_numbering(p):
    set_numpr(p, 0, 0)


def add_field(p, instr, result="", bold=None, italic=None, size=None):
    def run():
        r = p.add_run()
        if bold is not None:
            r.bold = bold
        if italic is not None:
            r.italic = italic
        if size:
            r.font.size = size
        return r
    r = run(); f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "begin"); r._r.append(f)
    r = run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = instr; r._r.append(it)
    r = run(); f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "separate"); r._r.append(f)
    run().text = result
    r = run(); f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), "end"); r._r.append(f)


BM_ID = [100]


def toc_bookmark(idx):
    """Имя закладки заголовка для оглавления (как у Word: _Toc + 9 цифр)."""
    return "_Toc%09d" % (100000000 + idx)


def bookmark_start(p, name):
    BM_ID[0] += 1
    el = OxmlElement("w:bookmarkStart")
    el.set(qn("w:id"), str(BM_ID[0]))
    el.set(qn("w:name"), name)
    p._p.append(el)
    return BM_ID[0]


def bookmark_end(p, bid):
    el = OxmlElement("w:bookmarkEnd")
    el.set(qn("w:id"), str(bid))
    p._p.append(el)


def tbl_put(table, tag, attrs=None, children=None):
    return put(table._tbl.tblPr, tag, attrs, TBLPR_ORDER, children)


def borders_el(spec):
    """spec: {'top': 'single', ...} -> w:tblBorders"""
    el = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        v = spec.get(edge, "none")
        e.set(qn("w:val"), v)
        e.set(qn("w:sz"), "4" if v != "none" else "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        el.append(e)
    return el


def table_borders(table, spec):
    tblPr = table._tbl.tblPr
    drop(tblPr, "tblBorders")
    el = borders_el(spec)
    idx = TBLPR_ORDER.index("tblBorders")
    for child in tblPr:
        cname = child.tag.split("}")[-1]
        if cname in TBLPR_ORDER and TBLPR_ORDER.index(cname) > idx:
            child.addprevious(el)
            return
    tblPr.append(el)


def full_width(table):
    tbl_put(table, "tblW", {"w": 5000, "type": "pct"})
    tbl_put(table, "tblLayout", {"type": "autofit"})


def row_flags(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    put(trPr, "cantSplit", order=TRPR_ORDER)
    if header:
        put(trPr, "tblHeader", order=TRPR_ORDER)


# ========================================================== стили шаблона
def fix_styles(doc):
    """Приводит стили шаблона к требованиям нормоконтроля ПМ (таблицы 4.1–5.3)."""
    st = doc.styles

    def sp(name):
        try:
            return st[name].element.get_or_add_pPr()
        except KeyError:
            return None

    normal = st["Normal"].element.get_or_add_pPr()
    put(normal, "widowControl")
    for name in (S_TEXT, S_AFTERTAB, S_TABCAP, S_FIGCAP, S_H2, S_H3):
        p = sp(name)
        if p is not None:
            put(p, "widowControl")
    p = sp(S_TABCAP)
    if p is not None:
        put(p, "keepNext")
        put(p, "jc", {"val": "left"})
        put(p, "suppressAutoHyphens")
    p = sp(S_FIG)
    if p is not None:
        put(p, "keepNext")
    p = sp(S_FIGCAP)
    if p is not None:
        put(p, "suppressAutoHyphens")
    for name in (S_LISTING, S_TABTEXT):
        p = sp(name)
        if p is not None:
            put(p, "jc", {"val": "left"})
    for name in (S_BULLET, S_NUMBER):
        p = sp(name)
        if p is not None:
            put(p, "ind", {"left": 1276, "hanging": 567})
    for name in (S_H1, S_H2, S_H3):
        p = sp(name)
        if p is not None:
            put(p, "suppressAutoHyphens")
    # шрифт без «запасных» гарнитур
    for name in (S_FIGCAP, S_FIG, S_AFTERTAB):
        try:
            rpr = st[name].element.get_or_add_rPr()
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is not None:
                for a in ("w:eastAsia", "w:cs"):
                    fonts.set(qn(a), "Times New Roman")
        except KeyError:
            pass
    # стиль «Сетка таблицы» не должен навязывать шрифт темы (Calibri/Aptos)
    try:
        tg = st["Table Grid"].element
        rpr = tg.get_or_add_rPr()
        for ch in list(rpr):
            if ch.tag in (qn("w:rFonts"), qn("w:sz"), qn("w:szCs")):
                rpr.remove(ch)
        f = OxmlElement("w:rFonts")
        for a in ("ascii", "hAnsi", "cs", "eastAsia"):
            f.set(qn("w:" + a), "Times New Roman")
        rpr.insert(0, f)
    except KeyError:
        pass
    # Оглавление 1–3: 14 пт, по ширине, 1,5 строки, без отступов, позиция табуляции
    # у правого поля с отточием; первый уровень — все прописные
    for lvl in (1, 2, 3):
        try:
            el = st["toc %d" % lvl].element
        except KeyError:
            continue
        p = el.get_or_add_pPr()
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right"); tab.set(qn("w:leader"), "dot"); tab.set(qn("w:pos"), str(TOC_TAB))
        put(p, "tabs", children=[tab])
        put(p, "spacing", {"before": 0, "after": 0, "line": 360, "lineRule": "auto"})
        put(p, "ind", {"left": 0, "right": 0, "firstLine": 0})
        put(p, "jc", {"val": "both"})
        put(p, "suppressAutoHyphens")
        rpr = el.get_or_add_rPr()
        for ch in list(rpr):
            rpr.remove(ch)
        f = OxmlElement("w:rFonts")
        for a in ("ascii", "hAnsi", "cs", "eastAsia"):
            f.set(qn("w:" + a), "Times New Roman")
        rpr.append(f)
        if lvl == 1:
            rpr.append(OxmlElement("w:caps"))
        for tag in ("sz", "szCs"):
            e = OxmlElement("w:" + tag); e.set(qn("w:val"), "28"); rpr.append(e)
    # нижний колонтитул: одинарный интервал, без отступа, по центру
    try:
        p = st["footer"].element.get_or_add_pPr()
        put(p, "spacing", {"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
        put(p, "ind", {"left": 0, "right": 0, "firstLine": 0})
        put(p, "jc", {"val": "center"})
    except KeyError:
        pass
    # без переносов слов во всём документе
    settings = doc.settings.element
    drop(settings, "autoHyphenation")


class Numbering:
    """Свои определения списков: один маркер «•» на все уровни, нумерация
    «1.», «1.1.», буквенные пункты «а)» без запрещённых букв, каждый список
    нумеруется заново."""

    def __init__(self, doc):
        self.el = doc.part.numbering_part.element
        self.abs_ids = [int(a.get(qn("w:abstractNumId"))) for a in self.el.findall(qn("w:abstractNum"))]
        self.num_ids = [int(n.get(qn("w:numId"))) for n in self.el.findall(qn("w:num"))]
        self.bullet_abs = self._abstract("bullet")
        self.number_abs = self._abstract("decimal")
        self.letter_abs = self._abstract("letter")
        self.bullet_num = self.new_num(self.bullet_abs)

    @staticmethod
    def _lvl(ilvl, fmt, text, start=1):
        rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:hint="default"/></w:rPr>'
        return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:tabs><w:tab w:val="num" w:pos="1276"/></w:tabs>'
                f'<w:ind w:left="1276" w:hanging="567"/></w:pPr>{rpr}</w:lvl>')

    def _abstract(self, kind):
        aid = max(self.abs_ids + [0]) + 1
        self.abs_ids.append(aid)
        lv = []
        for i in range(9):
            if kind == "bullet":
                lv.append(self._lvl(i, "bullet", "•"))
            elif kind == "decimal":
                lv.append(self._lvl(i, "decimal", ".".join(f"%{k + 1}" for k in range(i + 1)) + "."))
            else:
                lv.append(self._lvl(i, "decimal", "а)"))
        xml = (f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}">'
               f'<w:multiLevelType w:val="multilevel"/>' + "".join(lv) + '</w:abstractNum>')
        el = parse_xml(xml)
        first_num = self.el.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(el)
        else:
            self.el.append(el)
        return aid

    def new_num(self, abs_id, letter=None):
        nid = max(self.num_ids + [0]) + 1
        self.num_ids.append(nid)
        over = ('<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')
        if letter:
            over = ('<w:lvlOverride w:ilvl="0">' + self._lvl(0, "decimal", letter + ")") +
                    '</w:lvlOverride>')
        el = parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{abs_id}"/>'
                       f'{over}</w:num>')
        self.el.append(el)
        return nid


# ============================================================== рисунки
_PREP = {}


def prep_image(src, S, crop=True, frame_mode="авто"):
    """Обрезает белые поля, решает, нужна ли обводка. -> (путь, px_w, px_h, frame)
    Результат запоминается: вёрстка собирает документ много раз подряд."""
    try:
        memo = (os.path.abspath(src), os.path.getmtime(src), bool(crop), str(frame_mode))
    except OSError:
        memo = None
    if memo in _PREP:
        return _PREP[memo]
    res = _prep_image(src, S, crop, frame_mode)
    if memo is not None:
        _PREP[memo] = res
    return res


def _prep_image(src, S, crop=True, frame_mode="авто"):
    from PIL import Image, ImageChops
    try:
        im = Image.open(src)
        im.load()
    except Exception:
        return src, None, None, False
    w, h = im.size
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im.convert("RGBA"), mask=im.convert("RGBA").split()[-1])
        rgb = bg
    else:
        rgb = im.convert("RGB")
    out_im = rgb
    changed = rgb.mode != im.mode and im.mode in ("RGBA", "LA", "P")
    if crop:
        diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255))).convert("L")
        bbox = diff.point(lambda v: 255 if v > 18 else 0).getbbox()
        if bbox:
            pad = max(3, int(0.006 * max(w, h)))
            x0, y0 = max(0, bbox[0] - pad), max(0, bbox[1] - pad)
            x1, y1 = min(w, bbox[2] + pad), min(h, bbox[3] + pad)
            if (x0 > 0.015 * w or y0 > 0.015 * h or w - x1 > 0.015 * w or h - y1 > 0.015 * h):
                out_im = rgb.crop((x0, y0, x1, y1))
                changed = True
    frame = False
    fm = (frame_mode or "авто").lower()
    if fm in ("да", "yes"):
        frame = True
    elif fm == "авто":
        g = out_im.convert("L")
        W, H = g.size
        px = g.load()
        ring = []
        step = max(1, (W + H) // 400)
        for x in range(0, W, step):
            ring += [px[x, 0], px[x, H - 1]]
        for y in range(0, H, step):
            ring += [px[0, y], px[W - 1, y]]
        white = sum(1 for v in ring if v > 235) / max(1, len(ring))
        frame = white > 0.6
    if changed:
        cache = os.path.join(tempfile.gettempdir(), "build_report_img")
        os.makedirs(cache, exist_ok=True)
        base = re.sub(r"[^\w.-]", "_", os.path.basename(src))
        import hashlib
        tag = hashlib.md5(f"{os.path.abspath(src)}|{os.path.getmtime(src)}|{crop}".encode("utf-8")).hexdigest()[:16]
        dst = os.path.join(cache, f"{tag}_{base}.png")
        if not os.path.exists(dst):
            out_im.save(dst, dpi=im.info.get("dpi", (96, 96)))
        return dst, out_im.size[0], out_im.size[1], frame
    return src, w, h, frame


def add_outline(run):
    for spPr in run._r.iter("{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr"):
        a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ln = etree.SubElement(spPr, "{%s}ln" % a, w="6350")
        fill = etree.SubElement(ln, "{%s}solidFill" % a)
        etree.SubElement(fill, "{%s}srgbClr" % a, val="000000")


def make_placeholder(caption, width_cm, path, aspect=0.62):
    from PIL import Image, ImageDraw, ImageFont
    w = int(width_cm / 2.54 * 150)
    h = int(w * aspect)
    img = Image.new("RGB", (w, h), (245, 245, 245))
    d = ImageDraw.Draw(img)
    for x in range(0, w, 24):
        d.line([(x, 0), (min(x + 12, w), 0)], fill=(150, 150, 150), width=3)
        d.line([(x, h - 2), (min(x + 12, w), h - 2)], fill=(150, 150, 150), width=3)
    for y in range(0, h, 24):
        d.line([(0, y), (0, min(y + 12, h))], fill=(150, 150, 150), width=3)
        d.line([(w - 2, y), (w - 2, min(y + 12, h))], fill=(150, 150, 150), width=3)
    font = None
    for cand in ("C:/Windows/Fonts/arial.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                 "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
                 "/Library/Fonts/Arial.ttf"):
        if os.path.exists(cand):
            font = ImageFont.truetype(cand, max(14, w // 34))
            break
    cap = caption if len(caption) < 60 else caption[:57] + "..."
    d.multiline_text((w // 2, h // 2), "МЕСТО ДЛЯ РИСУНКА\n" + cap, fill=(110, 110, 110),
                     anchor="mm", align="center", font=font, spacing=10)
    img.save(path)


NUM_CELL = re.compile(r"[–\-−+±]?\s?[<>≈~]?\s?\d[\d\u00a0 ]*(?:[.,]\d+)?(?:[eE][–\-+]?\d+)?\s?%?"
                      r"(?:\s?[±]\s?\d+(?:[.,]\d+)?)?")


# ================================================================ рендер
class Renderer:
    def __init__(self, S, blocks, sources, labels, src_order, src_no, mode, layout, math,
                 markers=False):
        self.S = S
        self.blocks = blocks
        self.sources = sources
        self.labels = labels
        self.src_order = src_order
        self.src_no = src_no
        self.mode = mode
        self.L = layout
        self.math = math
        self.markers = markers
        self.marks = {}
        self.mark_no = 0
        self.bm_names = {}
        self.kept = set()           # таблицы и листинги, собранные неразрывными
        self.stats = {"split": 0, "shrink": 0, "moved": 0}

    # ------------------------------------------------------------ служебное
    def mark(self, p, kind, key=None, extra=None, size=None):
        """Невидимая метка для измерения вёрстки (только в черновых проходах).

        size — кегль текста абзаца, если он задан у прогонов, а не стилем (ячейки
        таблиц: 12 пт при стиле крупнее). Иначе метка делает строку выше, чем
        в чистовике, и на длинной таблице черновик уезжает на 1 пт за строку."""
        if not self.markers:
            return
        self.mark_no += 1
        mid = self.mark_no
        self.marks[mid] = (kind, key, extra)
        if getattr(self.L, "marker_mode", "text") == "bm":
            # Word: метка — закладка «zmN». Она не занимает места, поэтому черновик
            # верстается точно как чистовик; Word выгружает её в PDF с координатами.
            BM_ID[0] += 1
            st = OxmlElement("w:bookmarkStart")
            st.set(qn("w:id"), str(BM_ID[0]))
            st.set(qn("w:name"), f"zm{mid}")
            en = OxmlElement("w:bookmarkEnd")
            en.set(qn("w:id"), str(BM_ID[0]))
            pPr = p._p.pPr
            if pPr is not None:
                pPr.addnext(en)
                pPr.addnext(st)
            else:
                p._p.insert(0, en)
                p._p.insert(0, st)
            return
        r = p.add_run(f"#@{mid}@#")
        # Метка не должна менять вёрстку, иначе чистовик разойдётся с измеренным
        # черновиком. Кегль 1 пт нельзя: в строке без видимых знаков (пустая
        # строка листинга) высоту задаёт метка, строка «сжимается», и черновик
        # вмещает больше строк, чем чистовик. Поэтому кегль — как у текста абзаца,
        # а ширина символов 10 % (не меньше: при 1 % одинаковые цифры «11» налезают
        # друг на друга и PyMuPDF склеивает их) — метка почти не занимает места.
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        put(r._r.get_or_add_rPr(), "w", {"val": 10}, order=RPR_ORDER)
        if size:
            put(r._r.get_or_add_rPr(), "sz", {"val": int(size * 2)}, order=RPR_ORDER)
            put(r._r.get_or_add_rPr(), "szCs", {"val": int(size * 2)}, order=RPR_ORDER)
        if kind == "fig":
            # В абзаце рисунка любой текст добавляет строке нижний выносной элемент
            # шрифта (2–3 пт под картинкой). На каждом рисунке черновик уезжал вниз
            # относительно чистовика, и разрывы таблиц («Продолжение…») рассчитывались
            # не там, где они нужны в чистовике.
            # Высоту строки здесь задаёт картинка, поэтому метке можно дать кегль 1 пт:
            # выносной элемент ~0,25 пт. Ширину символов — обычную, чтобы цифры
            # не налезали друг на друга при чтении PDF.
            put(r._r.get_or_add_rPr(), "w", {"val": 100}, order=RPR_ORDER)
            put(r._r.get_or_add_rPr(), "sz", {"val": 2}, order=RPR_ORDER)
            put(r._r.get_or_add_rPr(), "szCs", {"val": 2}, order=RPR_ORDER)
        # метка должна стоять первой в абзаце
        pPr = p._p.pPr
        r._r.getparent().remove(r._r)
        if pPr is not None:
            pPr.addnext(r._r)
        else:
            p._p.insert(0, r._r)

    def bm(self, kind, key):
        k = (kind, key)
        if k not in self.bm_names:
            self.bm_names[k] = f"_Ref{kind[:3]}{len(self.bm_names) + 1:04d}"
        return self.bm_names[k]

    def labelled(self, p, kind, key, text, bold=None, italic=None):
        """Номер объекта внутри закладки — на неё ссылаются поля REF."""
        bid = bookmark_start(p, self.bm(kind, key))
        r = p.add_run(text)
        if bold is not None:
            r.bold = bold
        if italic is not None:
            r.italic = italic
        bookmark_end(p, bid)

    def para(self, style, text=None, **kw):
        p = self.doc.add_paragraph(style=style)
        if text:
            self.rich(p, text, **kw)
        return p

    # -------------------------------------------------------- текст строки
    def rich(self, p, text, bold=False, italic=False, typo=True, size=None, mono=False):
        state = {"b": bold, "i": italic}
        toks = tokenize(text)
        for n, (kind, val) in enumerate(toks):
            if kind == "t":
                if re.search(r"\\[a-zA-Z]{2,}", val):
                    val = delatex(val)          # LaTeX-команды вне $…$ (старый синтаксис)
                v = typography(val) if typo else val
                nxt = toks[n + 1] if n + 1 < len(toks) else None
                if nxt and nxt[0] == "ref" and re.search(
                        r"(Рисун\w*|Таблиц\w*|Листинг\w*|Приложени\w*|формул\w*)\s$", v):
                    v = v[:-1] + "\u00a0"
                self._emit_text(p, unesc(v), state, size, mono)
            elif kind == "code":
                r = p.add_run(val)
                r.font.name = "Courier New"
                r.bold = state["b"] or None
                if size:
                    r.font.size = size
            elif kind == "url":
                r = p.add_run(val)
                r.bold = state["b"] or None
                r.italic = state["i"] or None
                if size:
                    r.font.size = size
            elif kind == "math":
                el = self.math.get(val, False)
                if el is None:
                    r = p.add_run(delatex(val))
                    r.italic = True
                else:
                    p._p.append(copy.deepcopy(el))
            elif kind == "ref":
                self._emit_ref(p, val, state, size)

    def _emit_text(self, p, text, state, size, mono=False):
        parts = re.split(r"(\*\*|\*)", text)
        for k, part in enumerate(parts):
            if part == "**":
                state["b"] = not state["b"]
                continue
            if part == "*":
                prev = parts[k - 1] if k > 0 else ""
                nxt = parts[k + 1] if k + 1 < len(parts) else ""
                opening = not state["i"]
                if (opening and nxt[:1] and not nxt[:1].isspace()) or \
                        (not opening and prev[-1:] and not prev[-1:].isspace()):
                    state["i"] = not state["i"]
                    continue
            if not part:
                continue
            r = p.add_run(part)
            if state["b"]:
                r.bold = True
            if state["i"]:
                r.italic = True
            if size:
                r.font.size = size
            if mono:
                r.font.name = "Courier New"

    def _emit_ref(self, p, val, state, size):
        tag, _, key = val.partition(":")
        if tag == "ист":
            items = [x.strip() for x in key.split(",")]
            nums, pages = [], []
            for it in items:
                if not it:
                    continue
                if re.match(r"^(с|c|p|pp|стр)\.", it, re.I):
                    pages.append(it)
                else:
                    nums.append(it)
            nums_sorted = sorted(nums, key=lambda k: self.src_no.get(k, 9999))
            r = p.add_run("[")
            r.bold = state["b"] or None
            # подряд идущие номера 1, 2, 3 -> 1-3
            groups, cur = [], []
            for k in nums_sorted:
                n = self.src_no.get(k)
                if cur and n is not None and self.src_no.get(cur[-1]) is not None and \
                        n == self.src_no[cur[-1]] + 1:
                    cur.append(k)
                else:
                    if cur:
                        groups.append(cur)
                    cur = [k]
            if cur:
                groups.append(cur)
            for gi, g in enumerate(groups):
                if gi:
                    p.add_run(", ")
                seq = [g[0], g[-1]] if len(g) >= 3 else g
                for si, k in enumerate(seq):
                    if si:
                        p.add_run("-" if len(g) >= 3 else ", ")
                    n = self.src_no.get(k)
                    if n is None:
                        p.add_run("?")
                    else:
                        add_field(p, f" REF {self.bm('src', k)} ", str(n), bold=state["b"] or None,
                                  size=size)
            if pages:
                p.add_run(", " + typography(", ".join(pages)).replace(" -", "-"))
            r = p.add_run("]")
            r.bold = state["b"] or None
            return
        kind = REF_KIND.get(tag)
        key = key.strip()
        if kind == "app":
            key = key.upper()
        lab = self.labels.get((kind, key))
        if lab is None:
            r = p.add_run("??")
            r.bold = True
            return
        add_field(p, f" REF {self.bm(kind, key)} ", lab, bold=state["b"] or None,
                  italic=state["i"] or None, size=size)

    # ----------------------------------------------------------- секции
    def footer_with_page(self, section):
        section.footer.is_linked_to_previous = False
        fp = section.footer.paragraphs[0]
        try:
            fp.style = self.doc.styles["footer"]
        except KeyError:
            fp.style = self.doc.styles["Normal"]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.first_line_indent = Cm(0)
        add_field(fp, " PAGE ", "1", size=Pt(PAGE_NUM_PT))

    def new_section(self, landscape=None, restart=None):
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        sp = sec._sectPr
        drop(sp, "pgNumType")
        drop(sp, "titlePg")
        drop(sp, "vAlign")
        if landscape is None:
            landscape = self.state["land"]
        if landscape:
            sec.orientation = WD_ORIENT.LANDSCAPE
            sec.page_width, sec.page_height = self.PORT["h"], self.PORT["w"]
            sec.top_margin, sec.bottom_margin = Cm(3), Cm(1)
            sec.left_margin, sec.right_margin = Cm(2), Cm(2)
            sec.footer_distance = Cm(0.45)      # номер страницы — внутри нижнего поля 1 см
            put(sp, "vAlign", {"val": "center"}, SECT_ORDER)
        else:
            sec.orientation = WD_ORIENT.PORTRAIT
            sec.page_width, sec.page_height = self.PORT["w"], self.PORT["h"]
            sec.top_margin, sec.bottom_margin = self.PORT["t"], self.PORT["b"]
            sec.left_margin, sec.right_margin = self.PORT["l"], self.PORT["r"]
            sec.footer_distance = self.PORT["f"]
        if restart:
            put(sp, "pgNumType", {"start": restart}, SECT_ORDER)
        self.state["land"] = landscape
        self.state["fresh"] = True
        self._tidy_section_break()
        return sec

    def _tidy_section_break(self):
        """Переносит sectPr в предыдущий абзац, чтобы пустая строка не
        создавала лишнюю страницу."""
        body = self.doc.element.body
        children = list(body)
        if len(children) < 3:
            return
        holder = children[-2]
        if holder.tag != qn("w:p"):
            return
        prev = holder.getprevious()
        sect = holder.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        if sect is None:
            return
        if prev is not None and prev.tag == qn("w:p") and \
                prev.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None and \
                not prev.findall(".//" + qn("w:br")):
            ppr_prev = prev.find(qn("w:pPr"))
            if ppr_prev is None:
                ppr_prev = OxmlElement("w:pPr")
                prev.insert(0, ppr_prev)
            ppr_prev.append(sect)
            body.remove(holder)
        else:
            pp = holder.find(qn("w:pPr"))
            put(pp, "spacing", {"before": 0, "after": 0, "line": 20, "lineRule": "exact"})
            rpr = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "2"); rpr.append(sz)
            pp.find(qn("w:sectPr")).addprevious(rpr)

    def text_box(self):
        return LANDSCAPE_BOX if self.state["land"] else PORTRAIT_BOX

    # ------------------------------------------------------------- сборка
    def render(self):
        S = self.S
        doc = Document(TEMPLATE)
        self.doc = doc
        body = doc.element.body
        for el in list(body):
            if el.tag == qn("w:p"):
                body.remove(el)
        fix_styles(doc)
        self.num = Numbering(doc)
        base = doc.sections[0]
        self.PORT = dict(w=base.page_width, h=base.page_height, t=base.top_margin,
                         b=base.bottom_margin, l=base.left_margin, r=base.right_margin,
                         f=base.footer_distance or Cm(1.25))
        self.state = {"land": False, "fresh": False, "land_figs": 0}
        blocks = self.blocks

        # начало видимой нумерации страниц
        want = S.setting("нумерация_страниц_с")
        if want:
            start_idx = next((i for i, b in enumerate(blocks)
                              if b.kind == "h" and norm(b.text) == norm(want)), None)
            if start_idx is None:
                S.warn(f"нумерация_страниц_с: заголовок «{want}» не найден")
        else:
            # первый блок основной части: всё, что не аннотация и не содержание
            start_idx, in_front = None, False
            for i, b in enumerate(blocks):
                if b.kind == "h" and b.level == 1:
                    in_front = norm(b.text) in FRONT
                if b.kind in ("land_on", "land_off", "pagebreak"):
                    continue
                if not in_front:
                    start_idx = i
                    break
        ann_idx = next((i for i, b in enumerate(blocks)
                        if b.kind == "h" and b.level == 1 and norm(b.text) in ("АННОТАЦИЯ", "РЕФЕРАТ")
                        and (start_idx is None or i < start_idx)), None)
        after_ann_idx = None
        if ann_idx is not None:
            after_ann_idx = next((i for i in range(ann_idx + 1, len(blocks))
                                  if blocks[i].kind == "h" and blocks[i].level == 1), None)
        self.start_idx, self.ann_idx = start_idx, ann_idx

        # титул: страница-заглушка, форму вуза вставляешь сам
        tv = S.setting("титул").strip().lower()
        title_pages = int(tv) if tv.isdigit() else (0 if tv in ("нет", "no", "false") else 1)
        title_pages = min(title_pages, 10)
        has_title = title_pages > 0
        if has_title:
            self._title_stub(title_pages)
        self.after_title = has_title
        body_started = False
        if start_idx is None or (not has_title and start_idx == 0):
            self.footer_with_page(doc.sections[0])
            body_started = True

        has_content_after = [False] * len(blocks)
        seen = False
        for k in range(len(blocks) - 1, -1, -1):
            has_content_after[k] = seen
            if blocks[k].kind not in ("land_on", "land_off", "pagebreak"):
                seen = True

        self.prev_kind = None
        self.last_visual = None
        self.appendix_mode = False
        first_heading_done = False
        # заголовки, которые входят в оглавление: все после «СОДЕРЖАНИЕ»
        toc_at = next((i for i, b in enumerate(blocks)
                       if b.kind == "h" and is_toc_heading(norm(b.text))), None)
        self.has_toc = toc_at is not None
        self.toc_entries = [i for i, b in enumerate(blocks)
                            if b.kind == "h" and toc_at is not None and i > toc_at]
        self.ref_seen = set()
        self.cur_heading = None

        for idx, b in enumerate(blocks):
            if b.kind == "land_on":
                if self.state["land"]:
                    S.warn("[альбом] внутри уже открытого [альбом] — пропущено", b.line)
                elif has_content_after[idx]:
                    self.new_section(True)
                    self.state["land_figs"] = 0
                    self.state["land_key"] = obj_key(b)
                continue
            if b.kind == "land_off":
                if not self.state["land"]:
                    S.warn("[/альбом] без открывающего [альбом] — пропущено", b.line)
                elif has_content_after[idx]:
                    self.new_section(False)
                else:
                    self.state["land"] = False
                continue

            fresh_now, self.state["fresh"] = self.state["fresh"], False
            if self.after_title and b.kind != "h":
                will_section = idx == start_idx and not body_started
                if not fresh_now and not will_section:
                    self.page_break()
                self.after_title = False

            if b.kind == "h":
                up = norm(b.text)
                if is_appendix_heading(up):
                    self.appendix_mode = True
                started_here = False
                if idx == ann_idx:
                    self.new_section(False)
                    fresh_now = True
                elif idx == after_ann_idx and idx != start_idx:
                    self.new_section(False, restart=self.L.ann_page)
                    fresh_now = True
                if idx == start_idx and not body_started:
                    body_started = True
                    started_here = True
                    restart = self.L.ann_page if (ann_idx is not None and after_ann_idx == idx) else None
                    self.footer_with_page(self.new_section(False, restart=restart))
                self.heading(b, idx, up, no_break=(started_here or fresh_now or
                                              (not first_heading_done and not has_title)))
                first_heading_done = True
                continue

            if idx == start_idx and not body_started:
                body_started = True
                self.footer_with_page(self.new_section(False))
                fresh_now = True

            if b.kind == "p":
                style = S_AFTERTAB if self.prev_kind in ("table", "listing") else S_TEXT
                p = self.para(style)
                self.mark(p, "p", idx)
                self.rich(p, b.text)
                self.prev_kind, self.last_visual = "p", None
            elif b.kind == "li":
                self.list_item(b, idx)
            elif b.kind == "pagebreak":
                p = self.doc.add_paragraph()
                p.add_run().add_break(WD_BREAK.PAGE)
                self.mark(p, "break", idx)
            elif b.kind == "appendix":
                self.appendix(b, idx)
            elif b.kind == "fig":
                self.figure(b, idx, fresh_now)
            elif b.kind == "table":
                self.table(b, idx)
            elif b.kind == "listing":
                self.listing(b, idx)
            elif b.kind == "formula":
                self.formula(b, idx)

        if self.last_visual and not self.state["land"]:
            S.warn(f"документ не должен заканчиваться {self.last_visual} — добавьте "
                   f"заключительное предложение")
        return doc

    def _title_stub(self, pages=1):
        """Заглушка под титул (и задание, рабочий график…): pages страниц. Они считаются
        в нумерации, поэтому номера в содержании верны после вставки настоящих листов."""
        for k in range(pages):
            p = self.para(S_TEXT)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            if k:
                p.paragraph_format.page_break_before = True
            text = "[ ТИТУЛЬНЫЙ ЛИСТ — вставьте форму вашего вуза ]" if k == 0 else \
                f"[ ЛИСТ {k + 1} ИЗ {pages} ТИТУЛЬНОЙ ЧАСТИ — задание, рабочий график и т. п. ]"
            p.add_run(text).bold = True

    def page_break(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ------------------------------------------------------------ заголовок
    def heading(self, b, idx, up, no_break):
        S = self.S
        style = {1: S_H1, 2: S_H2, 3: S_H3}[b.level]
        if self.last_visual and self.prev_kind in ("fig", "table", "listing"):
            S.warn(f"раздел не должен заканчиваться {self.last_visual} — добавьте абзац "
                   f"текста перед «{b.text}»", b.line)
        p = self.para(style)
        self.mark(p, "h", idx)
        text = upper_plain(b.text) if b.level == 1 else b.text
        in_toc = idx in self.toc_entries
        bid = bookmark_start(p, toc_bookmark(idx)) if in_toc else None
        self.rich(p, text)
        if bid is not None:
            bookmark_end(p, bid)
        if not b.numbered:
            kill_numbering(p)
        if self.has_toc and not in_toc:
            # аннотация и само «СОДЕРЖАНИЕ» в оглавление не входят — и не попадут
            # в него, если поле оглавления обновят в Word (F9)
            put(ppr(p), "outlineLvl", {"val": 9})
        if is_toc_heading(up) or is_sources_heading(up):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            put(ppr(p), "ind", {"left": 0, "firstLine": 0})
        after_title = getattr(self, "after_title", False)
        self.after_title = False
        if no_break:
            put(ppr(p), "pageBreakBefore", {"val": 0})
        elif after_title and (b.level > 1 or (S.setting("разделы_с_новой_страницы").lower() == "нет"
                                              and not is_toc_heading(up))):
            put(ppr(p), "pageBreakBefore")
        elif after_title:
            pass
        elif b.level == 1 and S.setting("разделы_с_новой_страницы").lower() == "нет" \
                and not is_appendix_heading(up) and not is_sources_heading(up) \
                and not is_toc_heading(up):
            put(ppr(p), "pageBreakBefore", {"val": 0})
            p.paragraph_format.space_before = Pt(42)     # 15 мм, как у подразделов
        if is_toc_heading(up):
            self.write_toc()
        if is_sources_heading(up):
            self.write_sources()
        if is_appendix_heading(up):
            apps = [x for x in self.blocks if x.kind == "appendix"]
            if up == "ПРИЛОЖЕНИЯ" and apps:
                for a in apps:
                    pp = self.para(S_TEXT)
                    self.rich(pp, f"Приложение {a.letter} — {a.title.rstrip('.')}.")
        self.prev_kind, self.last_visual = "h", None
        self.cur_heading = b

    def write_toc(self):
        """Оглавление — настоящее поле Word (TOC \\o "1-3" \\h \\z \\u): каждая строка —
        ссылка на закладку заголовка, номер страницы — поле PAGEREF. Номера страниц
        подставлены из измеренной вёрстки, а при обновлении поля в Word (F9)
        оглавление пересобирается точно таким же."""
        entries = []
        for i in self.toc_entries:
            b = self.blocks[i]
            text = upper_plain(b.text) if b.level == 1 else b.text
            title = (b.num + " " + text) if b.num else text
            page = self.L.toc_pages.get(i, "99" if self.markers or not self.L.measured else "")
            entries.append((b.level, title, str(page), toc_bookmark(i)))
        instr = ' TOC \\o "1-3" \\h \\z \\u '
        if not entries:
            p = self.doc.add_paragraph(style=self.doc.styles["toc 1"])
            add_field(p, instr, "Обновите поле: Ctrl+A, F9")
            return

        def fld(parent, kind, text=None):
            r = OxmlElement("w:r")
            if kind == "instr":
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = text
                r.append(it)
            elif kind == "text":
                t = OxmlElement("w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = text
                r.append(t)
            elif kind == "tab":
                r.append(OxmlElement("w:tab"))
            else:
                f = OxmlElement("w:fldChar")
                f.set(qn("w:fldCharType"), kind)
                r.append(f)
            parent.append(r)
            return r

        last = None
        for n, (lvl, title, page, bm) in enumerate(entries):
            tp = self.doc.add_paragraph(style=self.doc.styles["toc %d" % lvl])
            if n == 0:
                fld(tp._p, "begin"); fld(tp._p, "instr", instr); fld(tp._p, "separate")
            link = OxmlElement("w:hyperlink")
            link.set(qn("w:anchor"), bm)
            link.set(qn("w:history"), "1")
            before = len(tp._p)
            self.rich(tp, title)
            for el in list(tp._p)[before:]:
                link.append(el)
            fld(link, "tab")
            fld(link, "begin"); fld(link, "instr", f" PAGEREF {bm} \\h "); fld(link, "separate")
            fld(link, "text", page)
            fld(link, "end")
            tp._p.append(link)
            last = tp
        fld(last._p, "end")

    def write_sources(self):
        for n, key in enumerate(self.src_order, 1):
            p = self.para(S_TEXT)
            self.mark(p, "src", key)
            bid = bookmark_start(p, self.bm("src", key))
            add_field(p, " SEQ Источник \\* ARABIC ", str(n))
            bookmark_end(p, bid)
            p.add_run(". ")
            desc = self.sources.get(key, (f"[ОПИСАНИЕ «{key}» НЕ ЗАДАНО]", 0))[0]
            self.rich(p, desc)

    # --------------------------------------------------------------- список
    def list_item(self, b, idx):
        prev = self.blocks[idx - 1] if idx > 0 else None
        cont = prev is not None and prev.kind == "li"
        if not cont:
            self._num_ctx = None
            self._let_no = 0
        if b.ordered == "num":
            style = S_NUMBER
            # новый список — нумерация с 1; маркированные вставки её не сбивают
            if self._num_ctx is None:
                self._num_ctx = self.num.new_num(self.num.number_abs)
            num_id, ilvl = self._num_ctx, b.level
        elif b.ordered == "let":
            style = S_NUMBER
            if prev is None or prev.kind != "li" or prev.ordered != "let":
                self._let_no = 0
            letter = LIST_LETTERS[min(self._let_no, len(LIST_LETTERS) - 1)]
            self._let_no += 1
            num_id, ilvl = self.num.new_num(self.num.letter_abs, letter=letter), 0
        else:
            style = S_BULLET
            num_id, ilvl = self.num.bullet_num, 0
        p = self.para(style)
        self.mark(p, "li", idx)
        set_numpr(p, num_id, ilvl)
        put(ppr(p), "ind", {"left": 1276, "hanging": 567})
        self.rich(p, b.text)
        self.prev_kind, self.last_visual = "li", None

    # ----------------------------------------------------------- приложение
    def appendix(self, b, idx):
        self.appendix_mode = True
        p = self.para(S_H3)
        self.mark(p, "app", idx)
        p.add_run("Приложение ")
        self.labelled(p, "app", b.letter, b.letter)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        put(ppr(p), "ind", {"left": 0, "firstLine": 0})
        put(ppr(p), "outlineLvl", {"val": 9})
        kill_numbering(p)
        p.paragraph_format.page_break_before = True
        cp = self.para(S_TEXT)
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.first_line_indent = Cm(0)
        cp.paragraph_format.keep_with_next = True
        self.rich(cp, b.title.rstrip("."))
        self.prev_kind, self.last_visual = "appendix", None

    # -------------------------------------------------------------- рисунок
    def figure(self, b, idx, fresh):
        S = self.S
        land = self.state["land"]
        if self.prev_kind in ("fig", "table") and not self.appendix_mode and not land:
            S.warn(f"Рисунок {b.number} идёт сразу за другим рисунком/таблицей — "
                   f"вставьте между ними абзац текста", b.line)
        max_w, max_h = self.text_box()
        reserve = 1.6 if not land else 1.4
        if self.prev_kind == "h":
            reserve += 2.6
        dflt = S.setting("ширина_рисунков")
        try:
            dflt = float(dflt.replace(",", ".")) if dflt else None
        except ValueError:
            dflt = None
        width = b.width or (24.0 if land else (dflt or 12.0))
        width = min(width, max_w)
        src = S.find_file(b.path)
        px = None
        frame = False
        if src:
            try:
                src, pw, ph, frame = prep_image(src, S, S.flag("обрезка_рисунков"),
                                                S.setting("рамка_рисунков"))
                aspect = ph / pw if pw else 0.62
                px = pw
            except Exception:
                aspect = 0.62
        else:
            aspect = 0.62 if not land else 0.55
        if width * aspect > max_h - reserve:
            width = (max_h - reserve) / aspect
        if not land and width * width * aspect > 0.75 * max_w * max_h:
            width = (0.75 * max_w * max_h / aspect) ** 0.5
        key = obj_key(b)
        orig_w = width
        if key in self.L.shrink and self.L.shrink[key] < width:
            width = self.L.shrink[key]
            self.stats["shrink"] += 1
        if px and width:
            dpi = px / (width / 2.54)
            if dpi < 140:
                S.warn(f"Рисунок {b.number}: низкое разрешение (~{int(dpi)} dpi при ширине "
                       f"{width:.1f} см) — нужен файл крупнее, иначе будет размыто", b.line)
        if not src:
            if b.path:
                S.warn(f"Рисунок {b.number}: файл «{b.path}» не найден — вставлена заглушка", b.line)
            tmpimg = os.path.join(tempfile.gettempdir(),
                                  re.sub(r"[^\w.-]", "_", f"ph_{b.id or b.number}.png"))
            make_placeholder(b.caption, width, tmpimg, aspect)
            src = tmpimg
        p = self.doc.add_paragraph(style=S_FIG)
        self.mark(p, "fig", key, {"idx": idx, "width": width, "aspect": aspect, "orig": orig_w})
        if land:
            if self.state["land_figs"] > 0:
                p.paragraph_format.page_break_before = True
            self.state["land_figs"] += 1
        run = p.add_run()
        if px:
            # Высота задаётся явно: в черновиках вёрстки вместо картинки стоит крошечная
            # заглушка того же размера — Word открывает и выгружает такой черновик быстрее
            pic_h = Emu(int(round(Cm(width) * aspect)))
            run.add_picture(tiny_image() if self.markers else src, width=Cm(width), height=pic_h)
        else:
            run.add_picture(src, width=Cm(width))
        if frame:
            add_outline(run)
        cap, cut = strip_end_period(b.caption)
        if cut:
            S.fix(f"убрана точка в конце подписи Рисунка {b.number}", b.line)
        if cap[:1].islower():
            cap = cap[0].upper() + cap[1:]
        cp = self.para(S_FIGCAP)
        self.mark(cp, "figcap", key)
        cp.add_run("Рисунок ")
        self.labelled(cp, "fig", b.id, b.number)
        cp.add_run(" — ")
        self.rich(cp, cap)
        self.prev_kind = "fig" if not land else "land_fig"
        self.last_visual = "рисунком" if not land else None

    # -------------------------------------------------------------- таблица
    def table(self, b, idx):
        S = self.S
        land = self.state["land"]
        if self.prev_kind in ("fig", "table") and not self.appendix_mode and not land:
            S.warn(f"Таблица {b.number} идёт сразу за другим объектом — вставьте между ними "
                   f"абзац текста", b.line)
        cap, cut = strip_end_period(b.caption)
        if cut:
            S.fix(f"убрана точка в конце названия Таблицы {b.number}", b.line)
        if cap[:1].islower():
            cap = cap[0].upper() + cap[1:]
        key = obj_key(b)
        p = self.para(S_TABCAP)
        self.mark(p, "tcap", key)
        p.add_run("Таблица ")
        self.labelled(p, "table", b.id, b.number)
        p.add_run(" — ")
        self.rich(p, cap)
        if not b.rows:
            S.warn(f"Таблица {b.number} пустая", b.line)
            return
        ncols = max(len(r) for r in b.rows)
        header = b.rows[0]
        data = b.rows[1:]
        hdr = []
        for c in header + [""] * (ncols - len(header)):
            c = c.strip()
            if re.match(r"^[а-яё]", c):
                c = c[0].upper() + c[1:]
                S.fix(f"заголовок графы Таблицы {b.number} начат с прописной буквы", b.line)
            if re.fullmatch(r"№(\s*п/п)?", c):
                S.warn(f"Таблица {b.number}: графу «№ п/п» в таблицу не включают", b.line)
            if c.endswith(".") and not re.search(r"\b\w{1,4}\.$", c):
                c = c.rstrip(".")
            hdr.append(c)
        # одинаковые единицы в каждой строке столбца
        for ci in range(ncols):
            col = [r[ci].strip() for r in data if ci < len(r) and r[ci].strip()]
            if len(col) >= 3 and all(re.search(r"\d\s*%$", c) for c in col):
                S.warn(f"Таблица {b.number}, графа «{hdr[ci]}»: знак % повторяется в каждой "
                       f"строке — укажите единицы в заголовке графы", b.line)
        repeat_head = S.setting("шапка_в_продолжении").lower() not in ("нет", "no", "0", "false")
        breaks = sorted(self.L.splits.get(key, []))
        segments, start = [], 0
        for br in breaks:
            if 0 < br < len(data):
                segments.append((start, br))
                start = br
        segments.append((start, len(data)))
        keep_whole = (len(segments) == 1 and len(data) <= KEEP_TABLE_ROWS and
                      key not in self.L.nokeep and not land)
        # первая строка не остаётся под шапкой одна: вместе с ней держится вторая
        keep_first = len(data) >= 3 and key not in self.L.nokeep and not land
        if keep_whole or keep_first:
            self.kept.add(key)
        for si, (a, z) in enumerate(segments):
            if si:
                cp = self.para(S_TABCAP)
                cp.paragraph_format.page_break_before = True
                cp.paragraph_format.space_before = Pt(0)
                cp.add_run(f"Продолжение Таблицы {b.number}")
                self.stats["split"] += 1
            t = self.doc.add_table(rows=0, cols=ncols)
            t.style = self.doc.styles["Table Grid"]
            full_width(t)
            last_seg = si == len(segments) - 1
            table_borders(t, {"top": "single", "left": "single", "right": "single",
                              "insideH": "single", "insideV": "single",
                              "bottom": "single" if last_seg else "none"})
            if si == 0 or repeat_head:
                cells = t.add_row().cells
                for ci in range(ncols):
                    self.cell(cells[ci], hdr[ci], bold=True, center=True, keep=True)
                row_flags(t.rows[0], header=True)
            for ri in range(a, z):
                row = data[ri]
                cells = t.add_row().cells
                # Метку строки ставим в самую короткую ячейку: в узкой ячейке с текстом
                # «впритык» даже метка шириной ~5 пт переносит слово на новую строку,
                # и черновик расходится с чистовиком.
                mark_ci = min(range(ncols), key=lambda c: len(plain_text(
                    row[c] if c < len(row) else "—").strip() or "—"))
                for ci in range(ncols):
                    val = row[ci] if ci < len(row) else ""
                    v = val.strip()
                    if v in ("", "-", "–", "—"):
                        v = "—"
                    center = bool(NUM_CELL.fullmatch(plain_text(v).replace("¤", "").strip() or "x")) \
                        or v == "—" or (v.startswith("$") and v.endswith("$"))
                    self.cell(cells[ci], v, center=center,
                              keep=(keep_whole and ri < z - 1) or (keep_first and ri == 0),
                              mark=("row", key, ri) if ci == mark_ci else None)
                row_flags(t.rows[-1])
        self.prev_kind = "table" if not land else "land_table"
        self.last_visual = "таблицей" if not land else None

    def cell(self, cell, text, bold=False, center=False, keep=False, mark=None, style=S_TABTEXT):
        cell.text = ""
        p = cell.paragraphs[0]
        p.style = self.doc.styles[style]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
        if keep:
            p.paragraph_format.keep_with_next = True
        if mark:
            self.mark(p, mark[0], mark[1], mark[2], size=12)
        self.rich(p, text, bold=bold, size=Pt(12))

    # -------------------------------------------------------------- листинг
    def listing(self, b, idx):
        S = self.S
        key = obj_key(b)
        cap, _ = strip_end_period(b.caption)
        if cap[:1].islower():
            cap = cap[0].upper() + cap[1:]
        lines = b.code.replace("\t", "    ").rstrip("\n").split("\n")
        if b.line_numbers:
            width = len(str(10 * len(lines)))
            lines = [f"{10 * (n + 1):>{width}}  {ln}" for n, ln in enumerate(lines)]
        if len(lines) > 45 and not self.appendix_mode:
            S.warn(f"Листинг {b.number}: {len(lines)} строк — в основном тексте листинг не "
                   f"больше 75 % страницы (~45 строк); вынесите код в приложение", b.line)
        p = self.para(S_TABCAP)
        self.mark(p, "lcap", key)
        p.add_run("Листинг ")
        self.labelled(p, "listing", b.id, b.number)
        p.add_run(" — ")
        self.rich(p, cap)
        breaks = sorted(self.L.splits.get(key, []))
        segments, start = [], 0
        for br in breaks:
            if 0 < br < len(lines):
                segments.append((start, br))
                start = br
        segments.append((start, len(lines)))
        keep_whole = (len(segments) == 1 and len(lines) <= KEEP_LISTING_LINES and
                      key not in self.L.nokeep)
        if keep_whole:
            self.kept.add(key)
        for si, (a, z) in enumerate(segments):
            if si:
                cp = self.para(S_TABCAP)
                cp.paragraph_format.page_break_before = True
                cp.paragraph_format.space_before = Pt(0)
                cp.add_run(f"Продолжение Листинга {b.number}")
                self.stats["split"] += 1
            t = self.doc.add_table(rows=0, cols=1)
            t.style = self.doc.styles["Table Grid"]
            full_width(t)
            last_seg = si == len(segments) - 1
            table_borders(t, {"top": "single", "left": "single", "right": "single",
                              "bottom": "single" if last_seg else "none"})
            for li in range(a, z):
                c = t.add_row().cells[0]
                c.text = ""
                pp = c.paragraphs[0]
                pp.style = self.doc.styles[S_LISTING]
                if (li == a and si == 0) or (keep_whole and li < z - 1):
                    pp.paragraph_format.keep_with_next = True
                self.mark(pp, "row", key, li)
                pp.add_run(lines[li] if lines[li].strip() else "\u00a0")   # обычный пробел схлопывается
                row_flags(t.rows[-1])
        self.prev_kind, self.last_visual = "listing", "листингом"

    # -------------------------------------------------------------- формула
    def formula(self, b, idx):
        empty = self.doc.add_paragraph(style=S_TEXT)
        empty.paragraph_format.keep_with_next = True
        tex = b.tex.strip()
        where = split_where(b.where)
        if where and not re.search(r"[,.;]\s*$", tex):
            tex = tex + ","
        el = self.math.get(tex, True)
        if b.number:
            t = self.doc.add_table(rows=1, cols=2)
            full_width(t)
            table_borders(t, {})
            c1, c2 = t.rows[0].cells
            for c, w in ((c1, 14.5), (c2, 2.5)):
                c.width = Cm(w)
                put(c._tc.get_or_add_tcPr(), "vAlign", {"val": "center"}, order=TCPR_ORDER)
            t.columns[0].width = Cm(14.5)
            t.columns[1].width = Cm(2.5)
            cells = ((c1, WD_ALIGN_PARAGRAPH.CENTER), (c2, WD_ALIGN_PARAGRAPH.RIGHT))
        else:
            fp = self.doc.add_paragraph(style=S_TEXT)
            cells = ((None, WD_ALIGN_PARAGRAPH.CENTER),)
        targets = []
        for c, align in cells:
            pp = c.paragraphs[0] if c is not None else fp
            pp.style = self.doc.styles[S_TEXT]
            pp.alignment = align
            pp.paragraph_format.first_line_indent = Cm(0)
            pp.paragraph_format.keep_with_next = bool(where)
            targets.append(pp)
        # Метку нельзя ставить в абзац с самой формулой: Word считает формулу с текстом
        # в абзаце строчной (∑ с индексами сбоку, мелкие дроби), строка становится ниже
        # на ~20 пт, и черновик расходится с чистовиком. Ставим её в ячейку номера,
        # а у формулы без номера — в пустой абзац перед ней (кегль метки = кегль
        # текста, поэтому высота пустого абзаца не меняется).
        self.mark(targets[1] if len(targets) > 1 else empty, "formula", idx)
        if el is not None:
            targets[0]._p.append(copy.deepcopy(el))
        else:
            if self.math.pandoc:
                self.S.warn(f"формула: не удалось собрать «{b.tex}» — вставлен текст", b.line)
            targets[0].add_run(delatex(tex)).italic = True
        if b.number:
            targets[1].add_run("(")
            self.labelled(targets[1], "formula", b.id, b.number)
            targets[1].add_run(")")
        # до и после формулы — по одной пустой строке; пояснения «где» идут после неё
        after = self.doc.add_paragraph(style=S_TEXT)
        if where:
            after.paragraph_format.keep_with_next = True
        for k, part in enumerate(where):
            wp = self.doc.add_paragraph(style=S_TEXT)
            pf = wp.paragraph_format
            pf.first_line_indent = Cm(-1.25) if k == 0 else Cm(0)
            pf.left_indent = Cm(1.25)
            pf.tab_stops.add_tab_stop(Cm(1.25))
            # С формулой держится только первая строка «где»: остальные пояснения могут
            # перейти на следующую страницу — иначе блок «пустая строка + формула +
            # пустая строка + пояснения» оставляет внизу страницы больше пяти строк.
            part = re.sub(r"[;.,]\s*$", "", part)
            tail = "." if k == len(where) - 1 else ";"
            if k == 0:
                wp.add_run("где\t")
            self.rich(wp, part + tail)
        self.prev_kind, self.last_visual = "formula", None


# ============================================================== вёрстка
class Layout:
    def __init__(self):
        self.toc_pages = {}
        self.splits = {}
        self.shrink = {}
        self.moves = {}
        self.ann_page = 2
        self.measured = False
        self.engine = None
        self.passes = 0
        self.pages = 0
        self.display_pages = 0
        self.pending = []
        self.nokeep = set()         # короткие таблицы/листинги, которые всё-таки пришлось разорвать
        self.frozen = set()         # объекты, правки которых зациклились: больше не трогаем
        self.ann_pages = 0          # сколько страниц занимает аннотация (в счёт не идёт)

    def display(self, pg):
        """Номер, который печатается на листе pg файла (аннотация не считается)."""
        if self.ann_pages and pg > self.ann_page - 1 + self.ann_pages:
            return pg - self.ann_pages
        return pg

    def state(self):
        """Решения вёрстки, от которых зависит черновик (для выбора лучшего прохода)."""
        return copy.deepcopy({"splits": self.splits, "shrink": self.shrink,
                              "moves": self.moves, "nokeep": sorted(self.nokeep)})

    def restore(self, st):
        self.splits, self.shrink, self.moves = (copy.deepcopy(st[k]) for k in
                                                ("splits", "shrink", "moves"))
        self.nokeep = set(st.get("nokeep", []))


def find_soffice():
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if os.path.exists(c):
            return c
    return None


PS_WORD = r'''
param([string]$InFile, [string]$OutFile, [int]$Bm = 0, [int]$Opt = 0)
$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($InFile)
    # 17 = PDF; последний аргумент: 2 — закладки Word в PDF (метки вёрстки),
    # 1 — закладки заголовков (навигация в готовом PDF), 0 — без них
    # 4-й аргумент: 0 — для печати, 1 — «для экрана» (черновики вёрстки: быстрее)
    $doc.ExportAsFixedFormat($OutFile, 17, $false, $Opt, 0, 1, 1, 0, $true, $true, $Bm)
    $doc.Saved = $true
    $doc.Close()
    $doc = $null
} finally {
    if ($doc -ne $null) { try { $doc.Saved = $true; $doc.Close() } catch {} }
    if ($word -ne $null) {
        try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
'''


def word_pids():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, timeout=20).stdout
        return set(re.findall(r'"WINWORD\.EXE","(\d+)"', out, re.I))
    except Exception:
        return set()


PS_WORD_LOOP = r"""
param([string]$JobFile)
$ErrorActionPreference = "Stop"
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    [Console]::Out.WriteLine("ready")
    [Console]::Out.Flush()
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($line -eq $null -or $line.Trim() -ne "go") { break }
        $job = @(Get-Content -LiteralPath $JobFile -Encoding UTF8)
        [string]$inFile = [string]$job[0]
        [string]$outFile = [string]$job[1]
        [int]$bm = [int]([string]$job[2])
        [int]$opt = [int]([string]$job[3])
        $doc = $null
        try {
            [Console]::Out.WriteLine("stage open"); [Console]::Out.Flush()
            $doc = $word.Documents.Open($inFile)
            [Console]::Out.WriteLine("stage export"); [Console]::Out.Flush()
            $doc.ExportAsFixedFormat($outFile, 17, $false, $opt, 0, 1, 1, 0, $true, $true, $bm)
            [Console]::Out.WriteLine("stage close"); [Console]::Out.Flush()
            $doc.Saved = $true
            $doc.Close()
            $doc = $null
            [Console]::Out.WriteLine("ok")
        } catch {
            if ($doc -ne $null) { try { $doc.Saved = $true; $doc.Close() } catch {} }
            [Console]::Out.WriteLine("fail")
        }
        [Console]::Out.Flush()
    }
} finally {
    if ($word -ne $null) {
        try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
"""


class WordSession:
    """Один скрытый Word на всю сборку: черновики вёрстки выгружаются в PDF без
    запуска Word на каждый проход. Если сеанс не поднялся или завис, сборщик
    возвращается к прежнему способу (отдельный запуск на каждый файл)."""

    def __init__(self):
        self.proc = None
        self.dead = os.name != "nt"
        self.q = queue.Queue()
        self.job = None
        self.before = set()
        self.seen = []              # всё, что написал сеанс (для отладки)

    def start(self):
        ps = shutil.which("powershell") or shutil.which("pwsh")
        if not ps:
            return False
        d = tempfile.mkdtemp(prefix="build_report_word_")
        script = os.path.join(d, "session.ps1")
        self.job = os.path.join(d, "job.txt")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(PS_WORD_LOOP)
        self.before = word_pids()
        try:
            self.proc = subprocess.Popen(
                [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-File", script, "-JobFile", self.job],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                text=True, encoding="utf-8", errors="replace")
        except Exception:
            self.proc = None
            return False

        def pump(out, q):
            try:
                for line in out:
                    q.put(line)
            except Exception:
                pass
            q.put("eof")

        threading.Thread(target=pump, args=(self.proc.stdout, self.q), daemon=True).start()
        return self.wait(("ready",), 90) == "ready"

    def wait(self, tokens, timeout):
        end = datetime.datetime.now() + datetime.timedelta(seconds=timeout)
        while True:
            left = (end - datetime.datetime.now()).total_seconds()
            if left <= 0:
                return None
            try:
                line = self.q.get(timeout=left).strip().lower()
            except queue.Empty:
                return None
            self.seen.append(line)
            if line == "eof":
                return None
            for t in tokens:
                if t in line:
                    return t

    def export(self, docx, pdf, bookmarks, optimize):
        if self.dead:
            return False
        if self.proc is None and not self.start():
            self.kill()
            return False
        try:
            with open(self.job, "w", encoding="utf-8-sig") as f:
                f.write("\n".join([os.path.abspath(docx), os.path.abspath(pdf),
                                   str(int(bookmarks)), str(int(optimize))]) + "\n")
            if os.path.exists(pdf):
                os.remove(pdf)
            self.proc.stdin.write("go\n")
            self.proc.stdin.flush()
        except Exception:
            self.kill()
            return False
        res = self.wait(("ok", "fail"), 120)
        try:
            with open(os.path.join(os.path.dirname(pdf), "session.log"), "a", encoding="utf-8") as f:
                f.write(f"{datetime.datetime.now():%H:%M:%S} {os.path.basename(docx)} -> {res}: "
                        f"{' | '.join(self.seen[-8:])}\n")
        except Exception:
            pass
        if res is None:              # завис или упал
            self.kill()
            return False
        return res == "ok" and os.path.exists(pdf)

    def close(self):
        if self.proc is None:
            return
        proc, self.proc = self.proc, None
        self.dead = True
        try:
            proc.stdin.write("quit\n")
            proc.stdin.flush()
            proc.stdin.close()
            proc.wait(timeout=40)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
            for pid in word_pids() - self.before:
                subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True)

    def kill(self):
        self.dead = True
        proc, self.proc = self.proc, None
        if proc is not None:
            try:
                proc.kill()
            except Exception:
                pass
            for pid in word_pids() - self.before:
                subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True)


WORD = WordSession()
if os.environ.get("REPORT_WORD_SESSION", "1") == "0":      # служебное: Word на каждый проход
    WORD.dead = True
# Черновики выгружаются так же, как чистовик («для печати»): по замеру в Word это не
# медленнее «для экрана», а страницы разбиваются в точности как в готовом PDF
DRAFT_OPT = 1 if os.environ.get("REPORT_DRAFT_OPT", "0") == "1" else 0
atexit.register(WORD.close)


def convert_word(docx, outdir, bookmarks=0, optimize=0):
    """bookmarks: 0 — без закладок, 1 — закладки заголовков, 2 — закладки Word (метки).
    optimize: 1 — «для экрана» (черновики), 0 — для печати (готовый PDF)."""
    if os.name != "nt":
        return None
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if WORD.export(docx, pdf, bookmarks, optimize):
        return pdf
    ps = shutil.which("powershell") or shutil.which("pwsh")
    if not ps:
        return None
    script = os.path.join(outdir, "export.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(PS_WORD)
    before = word_pids()
    try:
        if os.path.exists(pdf):
            os.remove(pdf)
        subprocess.run([ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-File", script, "-InFile", os.path.abspath(docx), "-OutFile",
                        os.path.abspath(pdf), "-Bm", str(int(bookmarks)),
                        "-Opt", str(int(optimize))],
                       capture_output=True, timeout=240)
    except subprocess.TimeoutExpired:
        for pid in word_pids() - before:
            subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True)
        return None
    except Exception:
        return None
    return pdf if os.path.exists(pdf) else None


def convert_lo(docx, outdir):
    so = find_soffice()
    if not so:
        return None
    prof = os.path.join(tempfile.gettempdir(), "build_report_lo_profile")
    from pathlib import Path
    try:
        subprocess.run([so, "-env:UserInstallation=" + Path(prof).as_uri(), "--headless",
                        "--convert-to", "pdf", "--outdir", outdir, os.path.abspath(docx)],
                       capture_output=True, timeout=300)
    except Exception:
        return None
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


def verify_toc(pdf, R, L):
    """Сверяет номера страниц оглавления с готовым PDF (Word кладёт в PDF закладки
    заголовков). -> {индекс блока: номер страницы} для строк, где номер разошёлся."""
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        d = fitz.open(pdf)
        outline = [e for e in d.get_toc(simple=True)
                   if not re.fullmatch(r"zm\d+", str(e[1]).strip())]
        d.close()
    except Exception:
        return None

    def key(t):
        return re.sub(r"[^0-9a-zа-я]+", "", norm(t).lower())

    want = []
    for i in R.toc_entries:
        b = R.blocks[i]
        want.append((i, key((b.num + " " if b.num else "") + plain_text(b.text))))
    # закладки PDF идут в порядке документа; сопоставляем по тексту, по порядку
    diff, pos = {}, 0
    matched = 0
    for i, k in want:
        for j in range(pos, len(outline)):
            ok = key(outline[j][1])
            if ok and (ok == k or ok.endswith(k) or k.endswith(ok)):
                page = L.display(outline[j][2])
                matched += 1
                if str(L.toc_pages.get(i, "")) != str(page):
                    diff[i] = page
                pos = j + 1
                break
    if matched < len(want):
        return None            # не все заголовки нашлись — сверка ненадёжна
    return diff


def convert(docx, outdir, S, L, marks=False, headings=False):
    pref = S.setting("вёрстка").lower()
    if pref in ("нет", "no"):
        return None
    order = []
    if L.engine:
        order = [L.engine]
    elif pref == "word":
        order = ["Word"]
    elif pref == "libreoffice":
        order = ["LibreOffice"]
    else:
        order = ["Word", "LibreOffice"]
    for eng in order:
        if eng == "Word":
            bm = 2 if (marks and getattr(L, "marker_mode", "") == "bm") else (1 if headings else 0)
            pdf = convert_word(docx, outdir, bm, optimize=DRAFT_OPT if marks else 0)
        else:
            pdf = convert_lo(docx, outdir)
        if pdf:
            L.engine = eng
            return pdf
    return None


def tiny_image():
    path = os.path.join(tempfile.gettempdir(), "build_report_tiny.png")
    if not os.path.exists(path):
        from PIL import Image
        Image.new("RGB", (8, 8), (225, 225, 225)).save(path)
    return path


MARK_RE = re.compile(r"#@(\d+)@#")


class Measure:
    def __init__(self):
        self.pages = []   # dict(no, w, h, land, bottom, limit, top, elems)
        self.marks = {}   # id -> (page, y0, y1)
        self.full = False


def analyze_pdf(pdf):
    M = Measure()
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
    except ImportError:
        fitz = None
    if fitz is not None:
        d = fitz.open(pdf)
        M.full = True
        bm_marks = {}
        for e in d.get_toc(simple=False):
            mm = re.fullmatch(r"zm(\d+)", str(e[1]).strip())
            if mm and len(e) > 3 and isinstance(e[3], dict) and e[3].get("to") is not None:
                bm_marks[int(mm.group(1))] = (e[2], float(e[3]["to"].y))
        for pno, page in enumerate(d, 1):
            W, H = page.rect.width, page.rect.height
            land = W > H
            limit = H - (1.0 if land else 2.0) * CM_PT
            top = (3.0 if land else 2.0) * CM_PT
            elems = []
            data = page.get_text("dict")
            for blk in data["blocks"]:
                if blk.get("type") == 1:
                    elems.append((blk["bbox"][1], blk["bbox"][3], "img", None))
                    continue
                for line in blk.get("lines", []):
                    for span in line.get("spans", []):
                        txt = span.get("text", "")
                        bb = span["bbox"]
                        for m in MARK_RE.finditer(txt):
                            M.marks.setdefault(int(m.group(1)), (pno, bb[1], bb[3]))
                        clean = MARK_RE.sub("", txt).strip()
                        if clean:
                            elems.append((bb[1], bb[3], "text", clean))
            for dr in page.get_drawings():
                r = dr["rect"]
                elems.append((r.y0, r.y1, "draw", None))
            if bm_marks:
                # закладка даёт верх строки; низ — по текстовому блоку (абзацу) на этом месте
                tblocks = [b["bbox"] for b in data["blocks"] if b.get("type") == 0]
                for mid, (bp, by) in bm_marks.items():
                    if bp != pno or mid in M.marks:
                        continue
                    y1 = by + 14
                    for bb in tblocks:
                        if bb[1] - 3 <= by <= bb[1] + 4:
                            y1 = max(y1, bb[3])
                    M.marks[mid] = (pno, by, y1)
            content = [e for e in elems if e[0] < limit - 8]   # нижний колонтитул не считаем
            bottom = max([e[1] for e in content], default=top)
            elems.sort(key=lambda e: e[0])
            M.pages.append(dict(no=pno, w=W, h=H, land=land, bottom=min(bottom, limit),
                                limit=limit, top=top, elems=[e for e in content]))
        d.close()
        return M
    # запасной путь: только номера страниц
    tool = shutil.which("pdftotext")
    if not tool:
        return None
    txt = subprocess.run([tool, "-layout", pdf, "-"], capture_output=True).stdout.decode("utf-8", "ignore")
    for pno, pg in enumerate(txt.split("\f"), 1):
        for m in MARK_RE.finditer(pg):
            M.marks.setdefault(int(m.group(1)), (pno, 0, 0))
        if pg.strip():
            M.pages.append(dict(no=pno, w=595, h=842, land=False, bottom=0, limit=0, top=0, elems=[]))
    return M


def final_gaps(pdf):
    """Проверка готового документа: пустое место перед «Продолжением» таблицы или
    листинга (так бывает, если чистовик сверстался не так, как измеренный черновик)."""
    M = analyze_pdf(pdf)
    notes = []
    if M is None or not M.full:
        return notes
    for i in range(len(M.pages) - 1):
        page, nxt = M.pages[i], M.pages[i + 1]
        if page["land"]:
            continue
        texts = [e for e in nxt["elems"] if e[2] == "text"]
        if not texts:
            continue
        first = min(texts, key=lambda e: e[0])[3]
        if not first.startswith("Продолжение"):
            continue
        free = page["limit"] - page["bottom"]
        if free > 2.5 * LINE_PT:
            notes.append(f"лист {page['no']} файла: перед «{first}» пусто ~{int(free / LINE_PT)} "
                         f"строк — пришлите этот вывод, чтобы поправить сборщик")
    return notes


def plan_layout(M, R, blocks, L, S, final_check=False):
    """По измерениям решает: номера страниц в содержании, переносы таблиц,
    уменьшение рисунков, перенос абзацев. Возвращает True, если что-то поменялось."""
    changed = False
    marks = R.marks
    pos = {}       # (kind, key) -> (page, y0, y1) первого вхождения
    rows = {}      # tkey -> {row: page}
    for mid, (kind, key, extra) in marks.items():
        if mid not in M.marks:
            continue
        pg, y0, y1 = M.marks[mid]
        if kind == "row":
            rows.setdefault(key, {})[extra] = (pg, y0)
        else:
            pos.setdefault((kind, key), (pg, y0, y1))
    L.pages = len(M.pages)

    # аннотация не учитывается в нумерации
    ann_pages = 0
    if R.ann_idx is not None and ("h", R.ann_idx) in pos:
        ann_start = pos[("h", R.ann_idx)][0]
        nxt = next((i for i in range(R.ann_idx + 1, len(blocks))
                    if blocks[i].kind == "h" and blocks[i].level == 1), None)
        if nxt is not None and ("h", nxt) in pos:
            ann_pages = pos[("h", nxt)][0] - ann_start
        if L.ann_page != ann_start:
            L.ann_page = ann_start
    L.ann_pages = ann_pages if R.ann_idx is not None else 0
    display = L.display

    new_toc = {}
    for i, b in enumerate(blocks):
        if b.kind == "h" and ("h", i) in pos:
            new_toc[i] = display(pos[("h", i)][0])
    if new_toc != L.toc_pages:
        L.toc_pages = new_toc
    L.display_pages = display(len(M.pages)) if M.pages else 0

    # Страницы, с которых вёрстка начинается заново (разрыв, приложение, смена
    # ориентации): правки в разных таких областях друг на друга не влияют.
    forced = {v[0] for k, v in pos.items() if k[0] in ("break", "app")}
    for i in range(1, len(M.pages)):
        if M.pages[i]["land"] != M.pages[i - 1]["land"]:
            forced.add(i + 1)
    forced = sorted(forced)

    def region(pg):
        return sum(1 for f in forced if f <= pg)

    # Таблицы и листинги. Любая правка сдвигает всё, что ниже, поэтому за проход
    # в каждой области применяется только самая верхняя правка (разрыв или пустое
    # место): выше неё вёрстка уже окончательная, и проходы сходятся сверху вниз.
    # Раньше правились все таблицы сразу, и снятые/добавленные «Продолжения»
    # раскачивали друг друга без конца.
    split_cands = {}   # область -> (страница, y, функция правки)

    def offer(pg, y, fn):
        reg = region(pg)
        if reg not in split_cands or (pg, y) < split_cands[reg][:2]:
            split_cands[reg] = (pg, y, fn)

    for tkey, rmap in rows.items():
        if tkey in L.frozen:
            continue
        cur = sorted(L.splits.get(tkey, []))
        stale_r = None
        if M.full:   # устаревший разрыв: перед «Продолжением» осталось пустое место
            for r in cur:
                if r - 1 in rmap:
                    pg_prev, y_prev = rmap[r - 1]
                    page = M.pages[pg_prev - 1]
                    if page["limit"] - page["bottom"] > 2.5 * LINE_PT:
                        stale_r = r
                        offer(pg_prev, y_prev,
                              lambda t=tkey, r=r, c=cur: L.splits.__setitem__(
                                  t, [x for x in c if x < r]))
                        break
        if stale_r is not None:
            continue
        order = sorted(rmap)
        for a, bidx in zip(order, order[1:]):   # строка ушла на новую страницу
            if rmap[bidx][0] != rmap[a][0] and bidx not in cur and bidx > 0:
                brk = bidx
                # последняя строка не переносится одна: вместе с ней уходит предыдущая,
                # если та невысокая (иначе перед «Продолжением» осталось бы пустое место)
                if M.full and bidx == order[-1] and a > max(cur + [0]) and a - 1 in rmap \
                        and rmap[a - 1][0] == rmap[a][0]:
                    page = M.pages[rmap[a][0] - 1]
                    if page["bottom"] - rmap[a][1] <= 2.0 * LINE_PT:
                        brk = a
                offer(rmap[a][0], rmap[a][1],
                      lambda t=tkey, b=brk, c=cur: L.splits.__setitem__(
                          t, sorted(set(c) | {b})))
                break

    L.pending = []     # страницы, где в этом проходе что-то правится (для отчёта)

    def apply_splits(skip=()):
        nonlocal changed
        if final_check:
            return
        for reg, (pg, y, fn) in split_cands.items():
            if reg not in skip:
                fn()
                changed = True
                L.pending.append(display(pg))

    if not M.full:
        apply_splits()
        return changed

    # полупустые страницы
    first_body = None
    if R.start_idx is not None and ("h", R.start_idx) in pos:
        first_body = pos[("h", R.start_idx)][0]
    elif R.start_idx is not None and ("p", R.start_idx) in pos:
        first_body = pos[("p", R.start_idx)][0]
    first_body = first_body or 1
    break_pages = {v[0] for k, v in pos.items() if k[0] == "break"}
    app_start = None
    for i, b in enumerate(blocks):
        if b.kind == "h" and is_appendix_heading(norm(b.text)) and ("h", i) in pos:
            app_start = pos[("h", i)][0]
            break
    by_page = {}
    for (kind, key), (pg, y0, y1) in pos.items():
        by_page.setdefault(pg, []).append((y0, y1, kind, key))
    for v in by_page.values():
        v.sort()

    def page_starter(pg):
        """Что открывает страницу: (kind, key) блока или None (продолжение текста)."""
        page = M.pages[pg - 1]
        if not page["elems"]:
            return None
        first = min(page["elems"], key=lambda e: e[0])
        top_group = [e for e in page["elems"] if e[0] <= first[0] + 3]
        imgs = [e for e in top_group if e[2] == "img"]
        texts = [e for e in top_group if e[2] == "text"]
        if imgs:                       # обводка рисунка рисуется линией — это всё равно рисунок
            first = imgs[0]
        elif texts:
            first = texts[0]
        cands = by_page.get(pg, [])
        if first[2] == "img":
            for y0, y1, kind, key in cands:
                if kind == "fig" and first[0] - 2 <= y0 <= first[1] + 6:
                    return (kind, key)
            return None
        for y0, y1, kind, key in cands:
            if first[0] - 3 <= y0 <= first[1] + 3:      # метка на той же строке
                return (kind, key)
        if first[2] == "draw":
            return ("table?", None)
        return None

    gap_done = set()   # области, где в этом проходе уже исправлено пустое место
    for pg in range(max(1, first_body), len(M.pages)):
        page = M.pages[pg - 1]
        nxt = M.pages[pg]
        if page["land"] or pg in break_pages:
            continue
        reg = region(pg)
        if not final_check and (reg in gap_done or
                                (reg in split_cands and split_cands[reg][0] <= pg)):
            continue   # ниже правки в этой области всё ещё сдвинется
        free = page["limit"] - page["bottom"]
        content_h = page["limit"] - page["top"]
        starter = page_starter(pg + 1)
        boundary = starter is not None and starter[0] in ("h", "app")
        land_block = None
        if nxt["land"]:
            land_block = next((b for b in blocks if b.kind == "land_on" and
                               _land_page(b, R, pos) == pg + 1), None)
            if land_block is not None:
                # альбомная страница открывается заголовком («Модели процесса» перед
                # [альбом]): предыдущий раздел закончен, это граница разделов
                j = blocks.index(land_block) + 1
                if j < len(blocks) and blocks[j].kind == "h":
                    boundary = True
                    starter = ("h", j)
        if boundary:
            if free > 0.5 * content_h and final_check and not (app_start and pg >= app_start):
                head = blocks[starter[1]].text if starter[0] == "h" else "приложения"
                need = max(1, int((free - 0.5 * content_h) / LINE_PT) + 1)
                L_note(S, f"с. {display(pg)}: свободно больше половины страницы перед "
                          f"«{head}» — между разделами допускается не более 50 % "
                          f"(допишите в конец раздела текста хотя бы на {n_lines(need)} "
                          f"или уберите разрыв)")
            continue
        if free <= (GAP_LINES + GAP_TOL) * LINE_PT:
            continue
        # внутри раздела пусто больше 5 строк
        fixed = False
        if final_check:
            pass
        elif starter and starter[0] == "fig" and not nxt["land"]:
            if starter[1] not in L.frozen:
                fixed = fix_figure_gap(starter[1], pg, free, M, R, blocks, L, S, pos)
        elif nxt["land"]:
            if land_block is not None and obj_key(land_block) not in L.frozen:
                fixed = fix_move(obj_key(land_block), free, M, R, blocks, L, pos)
        elif starter and starter[0] in ("tcap", "lcap") and starter[1] in R.kept \
                and starter[1] not in L.frozen:
            # короткая таблица (листинг) ушла на новую страницу целиком и оставила
            # больше пяти пустых строк — разрешаем её разорвать
            L.nokeep.add(starter[1])
            fixed = True
        if fixed:
            changed = True
            gap_done.add(reg)
            L.pending.append(display(pg))
        elif final_check:
            lines_free = f"{free / LINE_PT:.1f}".replace(".", ",")
            need = max(1, int(free / LINE_PT - GAP_LINES) + 1)      # сколько строк текста не хватает
            add = f"допишите перед ним текста хотя бы на {n_lines(need)}"
            if nxt["land"]:
                what, cure = "перед альбомной страницей", \
                    f"допишите перед [альбом] текста хотя бы на {n_lines(need)} или перенесите " \
                    f"[альбом] ниже по тексту"
            elif starter and starter[0] == "fig":
                what, cure = "рисунок не поместился", add + " или задайте рисунку меньшую ширину"
            elif starter and starter[0] in ("tcap", "lcap"):
                what, cure = "таблица (листинг) не поместилась", \
                    add.replace("перед ним", "перед ней")
            else:
                what, cure = "следующий блок не поместился", add
            L_note(S, f"с. {display(pg)}: внизу страницы пусто {lines_free} строки — {what}; "
                      f"внутри раздела допускается не более 5 строк ({cure})")
    apply_splits(skip=gap_done)
    return changed


def _land_page(b, R, pos):
    # первая страница альбомного блока = страница первого рисунка после метки
    idx = R.blocks.index(b)
    for j in range(idx + 1, len(R.blocks)):
        x = R.blocks[j]
        if x.kind == "fig" and ("fig", obj_key(x)) in pos:
            return pos[("fig", obj_key(x))][0]
        if x.kind == "land_off":
            break
    return None


def n_lines(n):
    """«1 строку», «3 строки», «7 строк»."""
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} строку"
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return f"{n} строки"
    return f"{n} строк"


def L_note(S, msg):
    if msg not in S.layout_notes:
        S.layout_notes.append(msg)


def fix_figure_gap(fkey, pg, free, M, R, blocks, L, S, pos):
    """Рисунок не влез внизу страницы pg. Уменьшаем его (если хватит до 15–30 %)
    или переносим следующий абзац перед ним."""
    info = None
    for mid, (kind, key, extra) in R.marks.items():
        if kind == "fig" and key == fkey:
            info = extra
            break
    if not info or ("figcap", fkey) not in pos or ("fig", fkey) not in pos:
        return False
    page = M.pages[pg]  # следующая страница (индекс pg = номер pg+1)
    cap_pg, cap_y0, cap_y1 = pos[("figcap", fkey)]
    img = [e for e in page["elems"] if e[2] == "img" and e[1] <= cap_y0 + 4]
    if not img:
        return False
    img_top = min(e[0] for e in img)
    img_h = max(e[1] for e in img) - img_top
    need = cap_y1 + 17 - img_top
    avail = free - 8
    if img_h <= 0 or need <= avail:
        return False
    other = need - img_h
    scale = (avail - other) / img_h
    width = info["width"]
    floor = max(7.0, 0.7 * info.get("orig", width))     # не мельче 70 % исходной ширины
    if scale >= 0.85 and width * scale - 0.05 >= floor:
        L.shrink[fkey] = round(width * scale - 0.05, 2)
        return True
    if fix_move(fkey, free, M, R, blocks, L, pos, fig_need=need):
        return True
    if scale >= 0.7 and width * scale - 0.05 >= floor:
        L.shrink[fkey] = round(width * scale - 0.05, 2)
        return True
    return False


def movable_after(blocks, idx):
    """Индексы абзацев, которые разрешено перенести перед объектом blocks[idx].

    Берутся только абзацы, стоявшие сразу после объекта в исходнике (по номерам строк:
    абзацы, уже перенесённые к следующему объекту, сюда не относятся). Перенос
    останавливается на абзаце, который вводит следующий объект («… показано на
    Рисунке 3.2»): иначе ссылка окажется раньше предыдущего рисунка, а два объекта —
    подряд. Если за абзацами идёт не текст, один абзац остаётся после объекта."""
    obj = blocks[idx]
    end = idx + 1
    last_line = obj.line
    if obj.kind == "land_on":
        while end < len(blocks) and blocks[end].kind != "land_off":
            end += 1
        if end < len(blocks):
            last_line = blocks[end].line
        end += 1
    later = {(x.kind, x.id) for x in blocks
             if x.kind in ("fig", "table", "listing") and x.line > last_line}
    # абзацы после следующего по исходнику объекта (их вёрстка могла подтянуть сюда) — чужие
    next_line = min([x.line for x in blocks if x.kind != "p" and x.line > last_line] or [10 ** 9])
    cands = []
    j = end
    stop_intro = False
    while j < len(blocks) and blocks[j].kind == "p" and last_line < blocks[j].line < next_line:
        refs = {(REF_KIND[m.group(1)], m.group(2).strip())
                for m in re.finditer(r"\{(рис|табл|лист):([^}]+)\}", blocks[j].text)}
        if refs & later:
            stop_intro = True
            break
        cands.append(j)
        j += 1
    after = blocks[j] if j < len(blocks) else None
    if not stop_intro and (after is None or after.kind != "p"):
        cands = cands[:-1]          # после объекта остаётся хотя бы один абзац
    return cands


def fix_move(okey, free, M, R, blocks, L, pos, fig_need=None):
    """Переносит абзацы, стоящие после объекта, перед ним, если они влезают
    в свободное место (какие абзацы можно трогать — см. movable_after).

    fig_need — высота рисунка с подписью: если первый абзац длиннее пустого места,
    он всё равно переносится (заполнит место и перетечёт на следующую страницу),
    когда рисунок после его остатка ещё помещается на странице."""
    idx = next((i for i, b in enumerate(blocks) if obj_key(b) == okey), None)
    if idx is None:
        return False
    cands = movable_after(blocks, idx)
    total, take = 0.0, 0
    for c in cands:
        a = pos.get(("p", c))
        nb = blocks[c + 1] if c + 1 < len(blocks) else None
        bkey = None
        if nb is not None:
            for kind in ("p", "h", "li", "formula"):
                if (kind, c + 1) in pos:
                    bkey = (kind, c + 1)
            if nb.kind == "fig" and ("fig", obj_key(nb)) in pos:
                bkey = ("fig", obj_key(nb))
            if nb.kind == "table" and ("tcap", obj_key(nb)) in pos:
                bkey = ("tcap", obj_key(nb))
            if nb.kind == "listing" and ("lcap", obj_key(nb)) in pos:
                bkey = ("lcap", obj_key(nb))
        if not a or not bkey:
            break
        b2 = pos[bkey]
        if b2[0] == a[0]:
            h = b2[1] - a[1]
        elif b2[0] == a[0] + 1 and M.full:
            # абзац перетекает на следующую страницу
            pa, pb = M.pages[a[0] - 1], M.pages[b2[0] - 1]
            h = (pa["bottom"] - a[1]) + (b2[1] - pb["top"])
        else:
            break
        if total + h > free - 10:
            if take == 0 and fig_need:
                page = M.pages[a[0] - 1]
                content_h = page["limit"] - page["top"]
                if (h - free) + fig_need <= content_h - LINE_PT:
                    take = 1
            break
        total += h
        take += 1
    if take <= 0:
        return False
    L.moves[okey] = L.moves.get(okey, 0) + take
    return True


# ================================================================ сборка
def prepare(src_text, S, layout):
    blocks, sources = parse(src_text, S)
    for b in blocks:
        if b.kind in ("p", "li"):
            b.text = prepare_text(b.text, S, b.line)
        elif b.kind == "table":
            b.rows = [[prepare_text(c, S, b.line) for c in r] for r in b.rows]
        elif b.kind == "fig":
            b.caption = prepare_text(b.caption, S, b.line)
        elif b.kind == "formula" and b.where:
            b.where = prepare_where(b.where)
    normalize_lists(blocks, S)
    blocks = apply_moves(blocks, layout.moves)
    labels, src_order, src_no, mode = enumerate_all(blocks, sources, S)
    return blocks, sources, labels, src_order, src_no, mode


def collect_all_math(blocks, sources):
    acc = set()
    for b in blocks:
        for t in all_texts(b):
            collect_math(t or "", acc)
        if b.kind == "formula":
            tex = b.tex.strip()
            where = b.where
            if where and not re.search(r"[,.;]\s*$", tex):
                tex += ","
            acc.add((tex, True))
        if b.kind == "h":
            collect_math(upper_plain(b.text) if b.level == 1 else b.text, acc)
    for desc, _ in sources.values():
        collect_math(desc, acc)
    return acc


def build_once(src_text, base_S, layout, math, markers):
    S = State()
    S.settings = dict(base_S.settings)
    S.search_dirs = base_S.search_dirs
    blocks, sources, labels, src_order, src_no, mode = prepare(src_text, S, layout)
    S.settings.update({k: v for k, v in base_S.settings.items()})
    math.S = S
    math.prepare(collect_all_math(blocks, sources))
    math.emit(S)
    lint_blocks(blocks, labels, S)
    R = Renderer(S, blocks, sources, labels, src_order, src_no, mode, layout, math, markers)
    doc = R.render()
    return S, R, doc, blocks


def main():
    ap = argparse.ArgumentParser(description="Сборка отчёта по нормоконтролю из report.md")
    ap.add_argument("source", nargs="?", default=None)
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--img", default=None, help="папка с рисунками")
    ap.add_argument("--init", nargs="?", const=".", default=None, metavar="ПАПКА",
                    help="новая работа: папка со сборщиком и заготовкой report.md")
    ap.add_argument("--update", action="store_true",
                    help="обновить сборщик в этой папке из общего комплекта")
    ap.add_argument("--pdf", action="store_true", help="сохранить ещё и PDF")
    ap.add_argument("--no-layout", action="store_true", help="не проверять вёрстку")
    ap.add_argument("--keep-temp", action="store_true", help=argparse.SUPPRESS)
    a = ap.parse_args()

    if a.init is not None:
        scaffold(a.init)
        return 0
    if a.update:
        return update_engine()
    if not os.path.exists(TEMPLATE):
        print("ОШИБКА: рядом с build.py нет файла template.docx")
        return 1
    source = a.source
    if source is None:
        for cand in (os.path.join(os.getcwd(), "report.md"), os.path.join(HERE, "report.md")):
            if os.path.exists(cand):
                source = cand
                break
    if not source or not os.path.exists(source):
        print(f"ОШИБКА: не найден исходник «{source or 'report.md'}».")
        print("Положите report.md рядом с build.py или укажите путь:")
        print("    python build.py путь/к/report.md")
        return 1

    src_dir = os.path.dirname(os.path.abspath(source)) or os.getcwd()
    imgdir = a.img or os.path.join(src_dir, "img")
    base = State()
    base.search_dirs = [src_dir, imgdir, os.path.join(src_dir, "code"), os.getcwd(), HERE]
    src_text = read_text(source)
    parse(src_text, base)                 # настройки
    base.warnings = []
    if a.no_layout:
        base.settings["вёрстка"] = "нет"

    fname = base.settings.get("имя_файла", "").strip() or "Отчёт"
    fname = re.sub(r'[\\/:*?"<>|]', "_", fname)
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    out = a.output or os.path.join(src_dir, fname)

    layout = Layout()
    math = MathCache(base)
    if a.keep_temp:   # отладка вёрстки: черновики и их PDF — рядом с исходником
        tmpdir = os.path.join(src_dir, "_отладка_вёрстки")
        shutil.rmtree(tmpdir, ignore_errors=True)
        os.makedirs(tmpdir, exist_ok=True)
    else:
        tmpdir = tempfile.mkdtemp(prefix="build_report_")
    pref = base.settings.get("вёрстка", "авто").lower()
    layout.marker_mode = "bm" if os.name == "nt" and pref in ("авто", "auto", "word", "") else "text"
    progress = sys.stdout.isatty()
    if pref not in ("нет", "no"):
        tried = []   # (правок осталось, проход, решения, содержание, страниц)
        seen = set()             # решения, по которым уже верстали (поиск зацикливания)
        converged = False

        def sig(st):
            return json.dumps(st, sort_keys=True, default=str, ensure_ascii=False)

        def measure(n):
            """Собирает черновик по текущим решениям и измеряет его. -> (S, R, blocks, M)"""
            S, R, doc, blocks = build_once(src_text, base, layout, math, markers=True)
            draft = os.path.join(tmpdir, f"draft{n}.docx")
            doc.save(draft)
            pdf = convert(draft, tmpdir, S, layout, marks=True)
            if pdf and layout.marker_mode == "bm" and layout.engine != "Word":
                # Word недоступен — LibreOffice закладки не выгружает: метки текстом
                layout.marker_mode = "text"
                S, R, doc, blocks = build_once(src_text, base, layout, math, markers=True)
                doc.save(draft)
                pdf = convert(draft, tmpdir, S, layout, marks=True)
            M = analyze_pdf(pdf) if pdf else None
            return S, R, blocks, M

        for n in range(MAX_PASSES):
            if progress:
                print(f"\rВёрстка: проход {n + 1}…", end="", flush=True)
            S, R, blocks, M = measure(n)
            if M is None:
                break
            layout.measured = True
            layout.passes = n + 1
            layout.full = M.full
            used = layout.state()          # решения, по которым свёрстан этот черновик
            seen.add(sig(used))
            while True:
                changed = plan_layout(M, R, blocks, layout, S)
                if not changed:
                    break
                if sig(layout.state()) not in seen:
                    break
                # Новые решения уже пробовались (или правка ничего не меняет): вёрстка
                # ходит по кругу. Объекты, вокруг которых она крутится, замораживаются,
                # и план строится заново по тому же черновику — без них.
                new = layout.state()
                loop = {k for name in ("splits", "shrink", "moves")
                        for k in set(used[name]) | set(new[name])
                        if used[name].get(k) != new[name].get(k)}
                loop |= set(used["nokeep"]) ^ set(new["nokeep"])
                layout.restore(used)
                if not loop or loop <= layout.frozen:
                    changed = False
                    break
                layout.frozen |= loop
            if not changed:
                plan_layout(M, R, blocks, layout, S, final_check=True)
                converged = True
                break
            # сколько ещё правок нужно этому черновику: меньше — лучше
            tried.append((len(layout.pending), n, used, dict(layout.toc_pages),
                          layout.display_pages))
        if progress and layout.measured:
            print("\r" + " " * 40 + "\r", end="", flush=True)
        if layout.measured and not converged and tried:
            # Не сошлось за отведённые проходы: в документ идёт измеренный черновик
            # с наименьшим числом оставшихся правок; он измеряется ещё раз, чтобы
            # перечислить в отчёте все страницы, которые нужно посмотреть глазами.
            left, n_best, st, toc, disp = min(tried, key=lambda t: (t[0], -t[1]))
            layout.restore(st)
            layout.toc_pages, layout.display_pages = toc, disp
            S, R, blocks, M = measure(MAX_PASSES)
            if M is not None:
                plan_layout(M, R, blocks, layout, S, final_check=True)
            S.layout_notes.insert(
                0, f"вёрстка не устоялась за {MAX_PASSES} проходов — взят лучший черновик; "
                   f"оставшиеся места перечислены ниже")
        if layout.measured and layout.frozen:
            names = {"fig": "рисунка", "table": "таблицы", "listing": "листинга", "land": "альбома"}
            what = ", ".join(f"{names.get(k.split(':')[0], 'объекта')} «{k.split(':', 1)[-1]}»"
                             for k in sorted(layout.frozen))
            S.layout_notes.append(f"вёрстка вокруг {what} не устоялась — посмотрите эти "
                                  f"места в документе глазами")
        layout_notes = list(S.layout_notes) if layout.measured else []
    else:
        layout_notes = []

    S, R, doc, blocks = build_once(src_text, base, layout, math, markers=False)
    try:
        doc.save(out)
    except PermissionError:
        print(f"ОШИБКА: файл «{out}» открыт в Word. Закройте его и повторите.")
        return 1

    pdf_out = None
    toc_fixed = False
    want_pdf = a.pdf or S.setting("pdf").lower() in ("да", "yes", "1")
    final_pdf = None
    if want_pdf or (layout.measured and getattr(layout, "full", False)):
        final_pdf = convert(out, tmpdir, S, layout, headings=True)
        # Сверка оглавления с готовым документом: номера страниц берутся из PDF
        # чистовика (закладки заголовков). Расхождение исправляется пересборкой.
        for _ in range(2):
            if not (final_pdf and layout.measured and layout.engine == "Word" and R.has_toc):
                break
            diff = verify_toc(final_pdf, R, layout)
            if not diff:
                break
            layout.toc_pages.update(diff)
            S, R, doc, blocks = build_once(src_text, base, layout, math, markers=False)
            try:
                doc.save(out)
            except PermissionError:
                break
            final_pdf = convert(out, tmpdir, S, layout, headings=True)
            toc_fixed = True
        if final_pdf and layout.measured:
            layout_notes += final_gaps(final_pdf)
    if want_pdf:
        pdf = final_pdf
        if pdf:
            pdf_out = os.path.splitext(out)[0] + ".pdf"
            try:
                shutil.copy(pdf, pdf_out)
            except PermissionError:
                print(f"PDF не сохранён: «{pdf_out}» открыт в другой программе")
                pdf_out = None
    WORD.close()
    if not a.keep_temp:
        shutil.rmtree(tmpdir, ignore_errors=True)

    # объём
    vol = S.setting("объём")
    if vol and layout.measured:
        m = re.match(r"^\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*$", vol)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2) or 10 ** 6)
            n = layout.display_pages
            if not lo <= n <= hi:
                layout_notes.append(f"объём {n} с. — по требованиям {vol}")

    # ---------------------------------------------------------------- итог
    print(f"Готово: {out}   (сборщик {VERSION})")
    if pdf_out:
        print(f"PDF:    {pdf_out}")
    nfig = sum(1 for b in blocks if b.kind == "fig")
    ntab = sum(1 for b in blocks if b.kind == "table")
    nfor = sum(1 for b in blocks if b.kind == "formula" and b.number)
    napp = sum(1 for b in blocks if b.kind == "appendix")
    print(f"Рисунков: {nfig}, таблиц: {ntab}, формул с номером: {nfor}, приложений: {napp}, "
          f"источников: {len(R.src_order)}; нумерация объектов: {R.mode.replace('_', ' ')}")
    if layout.measured:
        extra = []
        if R.stats["split"]:
            extra.append(f"продолжений таблиц/листингов: {R.stats['split']}")
        if R.stats["shrink"]:
            extra.append(f"уменьшено рисунков: {R.stats['shrink']}")
        moved = sum(layout.moves.values())
        if moved:
            extra.append(f"абзацев перенесено перед рисунком: {moved}")
        if toc_fixed:
            extra.append("номера страниц в содержании уточнены по готовому документу")
        print(f"Страниц: {layout.display_pages} (вёрстка: {layout.engine}, проходов: {layout.passes})"
              + ("; " + ", ".join(extra) if extra else ""))
        if not getattr(layout, "full", True):
            print("  (для проверки пустых мест установите PyMuPDF: python -m pip install pymupdf)")
    elif S.setting("вёрстка").lower() not in ("нет", "no"):
        print("Вёрстка не проверена: не найден ни Word, ни LibreOffice. Номера страниц в "
              "содержании: откройте документ, Ctrl+A, F9.")
    if S.fixes:
        print(f"\nИсправлено автоматически ({len(S.fixes)}):")
        for line, msg in S.fixes[:40]:
            print(f"  - {'стр. ' + str(line) + ': ' if line else ''}{msg}")
        if len(S.fixes) > 40:
            print(f"  … и ещё {len(S.fixes) - 40}")
    has_toc = any(b.kind == "h" and is_toc_heading(norm(b.text)) for b in blocks)
    if has_toc and not layout.measured:
        print("Номера страниц в содержании не проставлены: откройте документ в Word, "
              "Ctrl+A, затем F9 (обновить целиком).")
    if S.warnings:
        print(f"\nЗАМЕЧАНИЯ НОРМОКОНТРОЛЯ ({len(S.warnings)}):")
        for line, msg in S.warnings:
            print(f"  - {'стр. ' + str(line) + ' report.md: ' if line else ''}{msg}")
    if layout_notes:
        print(f"\nВЁРСТКА — проверьте эти страницы ({len(layout_notes)}):")
        for msg in layout_notes:
            print(f"  - {msg}")
    if not S.warnings and not layout_notes:
        print("\nЗамечаний нет.")
    return 0


# Файлы сборщика, которые кладутся в папку работы и обновляются командой --update.
# report.md, img/ и code/ — это сама работа, их сборщик не трогает никогда.
ENGINE = ["build.py", "template.docx", "Собрать.bat", "Отладка.bat", "Обновить сборщик.bat",
          "AGENTS.md", "CLAUDE.md", "GEMINI.md", "ФОРМАТ.md", "НОРМОКОНТРОЛЬ.md"]

UPDATER = ("@echo off\r\n"
           "rem Updates build.py and docs in this folder from the shared kit folder\r\n"
           "rem (the nearest folder with build.py and primery/ found above this one).\r\n"
           "chcp 65001 >nul\r\n"
           "set \"PY=python\"\r\n"
           "where python >nul 2>nul || set \"PY=py\"\r\n"
           "%PY% \"%~dp0build.py\" --update\r\n"
           "echo.\r\n"
           "pause\r\n")

STUB = """[настройки]
// Состав работы и настройки предмета — из Кафедры.md (профиль предмета).
// Всё необязательно; полный список и значения по умолчанию — в ФОРМАТ.md.
титул = да                        // да | нет | N — заглушка на N листов (титул, задание…)
нумерация_объектов = по_разделам  // по_разделам (Рисунок 1.1) | сквозная (Рисунок 1)
разделы_с_новой_страницы = нет    // да — каждый раздел 1-го уровня с новой страницы
имя_файла = ФамилияИО_ПРn         // имя готового .docx
pdf = да                          // сохранить ещё и PDF
// обязательно = СОДЕРЖАНИЕ; ВВЕДЕНИЕ; ЗАКЛЮЧЕНИЕ   // проверка, что разделы есть
// нумерация_страниц_с = ВВЕДЕНИЕ
// мин_источников = 5
// свежесть_источников = 5

# ЦЕЛЬ РАБОТЫ {-}

**Цель работы:** сформулируйте цель.

# Ход работы

## Постановка задачи

Опишите задание своими словами.

## Выполнение

Рисунки кладутся в папку img, ссылка пишется до рисунка: результат показан на
Рисунке {рис:result}.

[рис: result | img/result.png | Результат выполнения работы]

После каждого рисунка — хотя бы одно предложение с выводом по нему.

# ВЫВОДЫ {-}

1. Первый вывод.
2. Второй вывод.

Завершающее предложение после списка.
"""


def _copy_engine(src_dir, dst_dir):
    """Копирует файлы сборщика из src_dir в dst_dir. -> список скопированных."""
    done = []
    for name in ENGINE:
        src = os.path.join(src_dir, name)
        dst = os.path.join(dst_dir, name)
        if name == "Обновить сборщик.bat" and not os.path.exists(src):
            with open(dst, "w", encoding="ascii", newline="") as f:
                f.write(UPDATER)
            done.append(name)
            continue
        if not os.path.exists(src):
            continue
        if os.path.abspath(src) == os.path.abspath(dst):
            continue
        if os.path.exists(dst):              # «только чтение» не мешает обновлению
            os.chmod(dst, 0o666)
        shutil.copyfile(src, dst)
        os.chmod(dst, 0o666)
        done.append(name)
    return done


def scaffold(target="."):
    """Новая работа: самостоятельная папка со сборщиком (build.py, шаблон, .bat,
    инструкции для ИИ) и заготовкой report.md. Её можно копировать куда угодно;
    свежий сборщик потом подтягивается «Обновить сборщик.bat» (= --update)."""
    root = os.path.abspath(target)
    os.makedirs(root, exist_ok=True)
    for d in ("img", "code"):
        os.makedirs(os.path.join(root, d), exist_ok=True)
    made = _copy_engine(HERE, root)
    rp = os.path.join(root, "report.md")
    if not os.path.exists(rp):
        with open(rp, "w", encoding="utf-8") as f:
            f.write(STUB)
        made.append("report.md")
    print(f"Папка работы: {root}   (сборщик {VERSION})")
    print("Создано: " + ", ".join(made + ["img/", "code/"]))
    print("Дальше: заполните report.md, положите рисунки в img/ и дважды щёлкните "
          "Собрать.bat")


def is_kit(d):
    """Общая папка комплекта: build.py рядом с примерами."""
    return os.path.isfile(os.path.join(d, "build.py")) and \
        os.path.isdir(os.path.join(d, "примеры"))


def find_kit(start):
    """Ищет комплект: переменная REPORT_KIT, затем папки-соседи на каждом уровне
    вверх от start (ВУЗ\\Предмет\\ПР1 → ВУЗ\\Генерация отчётов)."""
    env = os.environ.get("REPORT_KIT")
    if env and is_kit(env):
        return os.path.abspath(env)
    cur = os.path.abspath(start)
    for _ in range(8):
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        try:
            names = sorted(os.listdir(parent))
        except OSError:
            names = []
        for n in names:
            d = os.path.join(parent, n)
            if os.path.isdir(d) and os.path.abspath(d) != os.path.abspath(HERE) and is_kit(d):
                return d
        cur = parent
    return None


def kit_version(d):
    try:
        with open(os.path.join(d, "build.py"), encoding="utf-8") as f:
            m = re.search(r'^VERSION = "([^"]+)"', f.read(), re.M)
        return m.group(1) if m else "?"
    except OSError:
        return "?"


def update_engine():
    """Обновляет сборщик в папке работы из общего комплекта."""
    if is_kit(HERE):
        print("Это сама папка комплекта — обновлять нечего (новая версия берётся с GitHub).")
        return 0
    kit = find_kit(HERE)
    if not kit:
        print("Не найдена папка комплекта (build.py + примеры/) выше этой папки.\n"
              "Положите работу внутрь той же папки, где лежит комплект (например, ВУЗ),\n"
              "или задайте путь: set REPORT_KIT=C:\\путь\\к\\Генерация отчётов")
        return 1
    old, new = VERSION, kit_version(kit)
    done = _copy_engine(kit, HERE)
    print(f"Комплект: {kit}")
    print(f"Сборщик: {old} → {new}" if old != new else f"Сборщик: {new} (уже свежий, файлы освежены)")
    print("Обновлено: " + ", ".join(done))
    print("report.md, img/ и code/ не тронуты.")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
