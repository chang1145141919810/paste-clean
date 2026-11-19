#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""clean_paste —— 把模型输出的原文，转成可以直接粘贴的 .docx

一条命令，一个功能：
    python3 clean_paste.py <输入文件>            # 输出同目录 <名字>.clean.docx
    python3 clean_paste.py <输入文件> -o <目录>   # 输出到指定目录
    type raw.txt | python3 clean_paste.py - -o out

设计原则（贯穿全文）：
  1. 只转换「语法上确定」的结构。拿不准的一律原样保留，绝不猜语义。
  2. 纯字符串变换 + 固定排序 + 固定 zip 时间戳 => 同样输入永远字节级同样输出。
  3. 不联网、不写配置、不留状态，跑完只在磁盘上多一个 .docx。
"""

from __future__ import annotations

import argparse
import hashlib
import io
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover - 只为了把错误说清楚
    sys.stderr.write(
        "缺少 python-docx。请用 DSH 自带的运行时执行：\n"
        "  python3 clean_paste.py <输入文件>\n"
    )
    raise SystemExit(2)


# ---------------------------------------------------------------------------
# 1. 常量：所有「认定规则」都集中在这里，方便逐条讲
# ---------------------------------------------------------------------------

# 句末标点。上一行以它结尾 => 换行是真的断段；否则视为硬折行，接回去。
TERMINAL = "。．！？!?；;：:…"

# 零宽字符：从聊天窗口复制出来时最常见的隐形垃圾。
ZERO_WIDTH = "\ufeff\u200b\u200c\u200d"

ATX_RE = re.compile(r"^ {0,3}(#{1,6})([ \t]*)(.*)$")
SETEXT_H1_RE = re.compile(r"^ {0,3}=+[ \t]*$")
HR_RE = re.compile(r"^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$")
FENCE_OPEN_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})(.*)$")
FENCE_CLOSE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})[ \t]*$")
BULLET_RE = re.compile(r"^([ \t]*)[-*+•·・][ \t]+(.*)$")
ORDERED_RE = re.compile(r"^([ \t]*)(\d{1,3})([.、)）])[ \t]*(.*)$")
QUOTE_RE = re.compile(r"^ {0,3}>[ \t]?(.*)$")
CHECKBOX_RE = re.compile(r"^\[([ xX])\][ \t]+(.*)$")
DELIM_CELL_RE = re.compile(r"^:?-+:?$")

HEADING_SIZES = {1: 16.0, 2: 14.0, 3: 12.5, 4: 11.5, 5: 11.0, 6: 10.5}
BODY_SIZE = 10.5           # 五号
CODE_SIZE = 9.0
ASCII_FONT = "Segoe UI"
EAST_ASIA_FONT = "微软雅黑"
CODE_FONT = "Consolas"

MAX_LIST_LEVEL = 5


# ---------------------------------------------------------------------------
# 2. 数据模型：先解析成「块」，再渲染。解析和渲染分开，
#    同一个解析结果以后要出 .md / .txt 也不用重写解析器。
# ---------------------------------------------------------------------------

@dataclass
class Block:
    kind: str                       # heading/para/bullet/number/table/code
    text: str = ""
    level: int = 0
    indent: int = 0                 # 原始缩进宽度，用来排列表层级
    marker: str = ""                # 编号列表的原文序号，如 "3."
    lines: list = field(default_factory=list)      # code 的行
    rows: list = field(default_factory=list)       # table 的单元格
    aligns: list = field(default_factory=list)     # table 的对齐


# ---------------------------------------------------------------------------
# 3. 行级清洗
# ---------------------------------------------------------------------------

def _strip_zero_width(s: str) -> str:
    for ch in ZERO_WIDTH:
        if ch in s:
            s = s.replace(ch, "")
    return s


def _split_leading(line: str):
    """拆成（缩进, 正文）。行首的全角空格也算缩进，不污染正文。"""
    i = 0
    while i < len(line) and line[i] in " \t\u3000":
        i += 1
    lead = line[:i].replace("\u3000", " ").expandtabs(4)
    return lead, line[i:]


def _is_cjk(ch: str) -> bool:
    return bool(ch) and ("\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef")


def _ends_terminal(s: str) -> bool:
    s = s.rstrip()
    while s and s[-1] in "”\"』」）)】]":
        s = s[:-1]
    return bool(s) and s[-1] in TERMINAL


def _join_sep(a: str, b: str) -> str:
    """中文折行接回去不加空格，英文才加。"""
    if not a or not b:
        return ""
    return "" if (_is_cjk(a[-1]) or _is_cjk(b[0])) else " "


def _title_like(line: str) -> bool:
    """能不能当标题：短，且没有句末标点。"""
    s = line.strip()
    return bool(s) and len(s) <= 40 and not _ends_terminal(s)


_ORPHAN_HEAD = re.compile(r"^\*{1,3}(?=\S)")
_ORPHAN_TAIL = re.compile(r"(?<=\S)\*{1,3}$")


def _strip_orphan_markers(line: str) -> str:
    """只处理「孤零零、没有配对」的星号：成对的交给 parse_inline。"""
    if "*" not in line:
        return line
    m = _ORPHAN_HEAD.match(line)
    if m and "*" not in line[m.end():]:
        line = line[m.end():]
    m = _ORPHAN_TAIL.search(line)
    if m and "*" not in line[:m.start()]:
        line = line[:m.start()]
    return line


# ---------------------------------------------------------------------------
# 4. 表格
# ---------------------------------------------------------------------------

def _table_line(line: str) -> str:
    """全角竖线 ｜ 也当分隔符（从中文窗口复制时常带着）。"""
    return line.replace("｜", "|")


def _has_pipe(line: str) -> bool:
    return "|" in _table_line(line)


def _split_row(line: str) -> list:
    s = _table_line(line).strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    cells = re.split(r"(?<!\\)\|", s)
    return [c.strip().replace("\\|", "|") for c in cells]


def _is_delim_row(line: str) -> bool:
    if not _has_pipe(line):
        return False
    cells = _split_row(line)
    if len(cells) < 2:
        return False
    return all(DELIM_CELL_RE.match(c or "") for c in cells)


def _read_aligns(line: str) -> list:
    out = []
    for c in _split_row(line):
        left = c.startswith(":")
        right = c.endswith(":")
        out.append("center" if left and right else "right" if right else "left")
    return out


def _pipe_run(lines, i) -> int:
    """从第 i 行起，连续多少行的竖线列数完全一致。"""
    width = len(_split_row(lines[i]))
    n = len(lines)
    j = i
    while j < n and lines[j].strip() and _has_pipe(lines[j]) \
            and len(_split_row(lines[j])) == width:
        j += 1
    return j - i


def _find_fence_close(lines, i, fence) -> int:
    """找配对的反引号围栏。找不到就返回 -1：宁可不当代码块，
    也不要把后面整篇文档都吞进代码里。"""
    ch, count = fence[0], len(fence)
    n = len(lines)
    j = i + 1
    while j < n:
        cm = FENCE_CLOSE_RE.match(lines[j])
        if cm and cm.group(1)[0] == ch and len(cm.group(1)) >= count:
            return j
        j += 1
    return -1


# ---------------------------------------------------------------------------
# 5. 行内格式：**粗** *斜* ~~删除~~ `代码` [文字](链接)
#    规则刻意保守：边界不满足就当普通文字。宁可少变，不要错变。
# ---------------------------------------------------------------------------

_SAFE_EDGE = " \t（(【[「『“‘\"'，。：；、,.:;"


def _can_open(text, i):
    prev = text[i - 1] if i > 0 else " "
    if prev.isascii() and prev.isalnum():
        return False
    return prev in _SAFE_EDGE or not prev.isascii()


def _find_close(text, start, marker):
    j = start
    while True:
        j = text.find(marker, j)
        if j == -1:
            return -1
        nxt = text[j + len(marker)] if j + len(marker) < len(text) else " "
        if (nxt.isascii() and nxt.isalnum()) or text[j - 1] == " ":
            j += len(marker)
            continue
        return j


def parse_inline(text, bold=False, italic=False, strike=False):
    runs = []
    buf = []

    def flush():
        if buf:
            runs.append({
                "text": "".join(buf), "bold": bold, "italic": italic,
                "strike": strike, "code": False, "link": None,
            })
            del buf[:]

    i, n = 0, len(text)
    while i < n:
        ch = text[i]

        if ch == "`":
            j = text.find("`", i + 1)
            if j > i + 1:
                flush()
                runs.append({"text": text[i + 1:j], "bold": bold, "italic": italic,
                             "strike": strike, "code": True, "link": None})
                i = j + 1
                continue

        if text.startswith("***", i):
            j = text.find("***", i + 3)
            if j > i + 3:
                flush()
                runs.extend(parse_inline(text[i + 3:j], True, True, strike))
                i = j + 3
                continue

        if text.startswith("**", i) or text.startswith("__", i):
            m = text[i:i + 2]
            j = text.find(m, i + 2)
            while j != -1 and (j == i + 2 or text[j - 1].isspace()):
                j = text.find(m, j + 2)
            if j > i + 2:
                inner = text[i + 2:j]
                # __init__ 这类标识符不要误判成加粗
                if m == "__" and "_" in inner and inner.replace("_", "").isalnum():
                    pass
                else:
                    flush()
                    runs.extend(parse_inline(inner, True, italic, strike))
                    i = j + 2
                    continue

        if text.startswith("~~", i):
            j = text.find("~~", i + 2)
            if j > i + 2:
                flush()
                runs.extend(parse_inline(text[i + 2:j], bold, italic, True))
                i = j + 2
                continue

        if ch == "[":
            m = re.match(r"\[([^\]]*)\]\(\s*([^)\s]+)\s*\)", text[i:])
            if m:
                flush()
                runs.append({"text": m.group(1) or m.group(2), "bold": bold,
                             "italic": italic, "strike": strike, "code": False,
                             "link": m.group(2)})
                i += m.end()
                continue

        if ch == "<":
            m = re.match(r"<(https?://[^>\s]+)>", text[i:])
            if m:
                flush()
                runs.append({"text": m.group(1), "bold": bold, "italic": italic,
                             "strike": strike, "code": False, "link": m.group(1)})
                i += m.end()
                continue

        if ch == "*" and not text.startswith("**", i):
            if _can_open(text, i):
                j = _find_close(text, i + 1, "*")
                if j > i + 1 and not text[i + 1].isspace() and "*" not in text[i + 1:j]:
                    flush()
                    runs.extend(parse_inline(text[i + 1:j], bold, True, strike))
                    i = j + 1
                    continue

        if ch == "_" and not text.startswith("__", i):
            prev = text[i - 1] if i > 0 else " "
            if not (prev.isascii() and prev.isalnum()):
                j = _find_close(text, i + 1, "_")
                if j > i + 1 and not text[i + 1].isspace() and "_" not in text[i + 1:j]:
                    nxt = text[j + 1] if j + 1 < len(text) else " "
                    if not (nxt.isascii() and nxt.isalnum()):
                        flush()
                        runs.extend(parse_inline(text[i + 1:j], bold, True, strike))
                        i = j + 1
                        continue

        buf.append(ch)
        i += 1

    flush()
    return runs


# ---------------------------------------------------------------------------
# 6. 块解析：一行一行过状态机
# ---------------------------------------------------------------------------

def _build_para(lines):
    cur = lines[0]
    for nxt in lines[1:]:
        cur = cur + _join_sep(cur, nxt) + nxt
    return cur


def parse_blocks(text: str):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _strip_zero_width(text)
    raw_lines = text.split("\n")
    lines = [ln.rstrip() for ln in raw_lines]

    blocks = []
    para = []

    def flush_para():
        if para:
            blocks.append(Block("para", text=_build_para(para)))
            del para[:]

    def push_text(line):
        if para and _ends_terminal(para[-1]):
            flush_para()
        para.append(_strip_orphan_markers(line))

    i, n = 0, len(lines)
    while i < n:
        line = lines[i]

        if not line.strip():
            flush_para()
            i += 1
            continue

        # 围栏代码块：内容原样保留，只在块边界上收口
        m = FENCE_OPEN_RE.match(line)
        if m:
            flush_para()
            close = _find_fence_close(lines, i, m.group(1))
            if close < 0:
                # 没有配对的围栏标记，当成一条残留符号丢掉，正文照常解析
                i += 1
                continue
            blocks.append(Block("code", lines=raw_lines[i + 1:close]))
            i = close + 1
            continue

        # 表格 1：标准 GFM，表头 + |---| 分隔行
        if (_has_pipe(line) and i + 1 < n and _is_delim_row(lines[i + 1])
                and not _is_delim_row(line)):
            flush_para()
            rows = [_split_row(line)]
            aligns = _read_aligns(lines[i + 1])
            j = i + 2
            while (j < n and lines[j].strip() and _has_pipe(lines[j])
                   and not _is_delim_row(lines[j]) and not ATX_RE.match(lines[j])
                   and not FENCE_OPEN_RE.match(lines[j])):
                rows.append(_split_row(lines[j]))
                j += 1
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            aligns = (aligns + ["left"] * width)[:width]
            blocks.append(Block("table", rows=rows, aligns=aligns))
            i = j
            continue

        # 表格 2：丢了分隔行的竖线表。连续 >=2 行、列数一致才算，
        # 否则普通带竖线的句子会被误吞。
        if _has_pipe(line) and _pipe_run(lines, i) >= 2:
            flush_para()
            width = len(_split_row(line))
            rows = []
            j = i
            while (j < n and lines[j].strip() and _has_pipe(lines[j])
                   and not _is_delim_row(lines[j]) and not ATX_RE.match(lines[j])
                   and not FENCE_OPEN_RE.match(lines[j])
                   and len(_split_row(lines[j])) == width):
                rows.append(_split_row(lines[j]))
                j += 1
            blocks.append(Block("table", rows=rows, aligns=["left"] * width))
            i = j
            continue

        # ATX 标题：#、##……  允许 #标题 这种省略空格的写法，但不吃 #1
        m = ATX_RE.match(line)
        if m:
            hashes, gap, body = m.group(1), m.group(2), m.group(3)
            if bool(gap) or (bool(body) and not body[0].isascii()):
                flush_para()
                body = re.sub(r"[ \t]*#+[ \t]*$", "", body).strip()
                blocks.append(Block("heading", text=body, level=len(hashes)))
                i += 1
                continue

        # 分隔线：整行只有 *** / --- / ___
        if HR_RE.match(line):
            flush_para()
            i += 1
            continue

        # setext 一级标题：短行 + 下一行 ===。长句或带句号的上一行不算标题，
        # 那条 === 只当分隔线丢掉。
        if SETEXT_H1_RE.match(line):
            if len(para) == 1 and _title_like(para[-1]):
                title = para.pop()
                blocks.append(Block("heading", text=title.strip(), level=1))
            i += 1
            continue

        # 引用：只摘掉 >，不猜它的语义
        m = QUOTE_RE.match(line)
        if m:
            inner = m.group(1).strip()
            if inner:
                push_text(inner)
            i += 1
            continue

        lead, _body = _split_leading(line)
        indent = len(lead)

        # 无序列表
        m = BULLET_RE.match(line)
        if m:
            flush_para()
            rest = m.group(2).strip()
            cm = CHECKBOX_RE.match(rest)
            if cm:
                rest = ("☑ " if cm.group(1).lower() == "x" else "☐ ") + cm.group(2).strip()
            blocks.append(Block("bullet", text=rest, indent=indent))
            i += 1
            continue

        # 有序列表：1. / 1、/ 1) / 1）
        m = ORDERED_RE.match(line)
        if m:
            sep = m.group(3)
            after = m.group(4)
            # "1.5 倍" 这种数字不要当列表
            number_ok = True
            if sep == "." and after and after[0].isdigit() and not m.group(0)[len(m.group(1)) + len(m.group(2)) + 1:].startswith(" "):
                number_ok = False
            if number_ok:
                flush_para()
                blocks.append(Block("number", text=after.strip(), indent=indent,
                                    marker="%s%s" % (m.group(2), sep)))
                i += 1
                continue

        push_text(line.strip())
        i += 1

    flush_para()
    _assign_list_levels(blocks)
    return blocks


def _assign_list_levels(blocks):
    """缩进宽度按出现顺序排名 => 最浅的一级算第 1 层。
    比「每 2 空格一层」稳，因为模型有时 2 空格、有时 4 空格。"""
    widths = sorted({b.indent for b in blocks if b.kind in ("bullet", "number")})
    rank = {w: min(idx, MAX_LIST_LEVEL) for idx, w in enumerate(widths)}
    for b in blocks:
        if b.kind in ("bullet", "number"):
            b.level = rank.get(b.indent, 0)


# ---------------------------------------------------------------------------
# 7. docx 渲染
# ---------------------------------------------------------------------------

def _set_style_fonts(style, ascii_name, ea_name):
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        pos = 1 if rpr.find(qn("w:rStyle")) is not None else 0
        rpr.insert(pos, rfonts)
    rfonts.set(qn("w:ascii"), ascii_name)
    rfonts.set(qn("w:hAnsi"), ascii_name)
    rfonts.set(qn("w:cs"), ascii_name)
    rfonts.set(qn("w:eastAsia"), ea_name)


def _set_run_fonts(run, ascii_name, ea_name):
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), ascii_name)
    rfonts.set(qn("w:hAnsi"), ascii_name)
    rfonts.set(qn("w:eastAsia"), ea_name)


_PPR_AFTER_SHD = (
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)


def _shade_paragraph(paragraph, fill):
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn("w:shd"))
    if old is not None:
        ppr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.insert_element_before(shd, *_PPR_AFTER_SHD)


def _box_paragraph(paragraph, color="D9D9D9"):
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn("w:pBdr"))
    if old is not None:
        ppr.remove(old)
    pbdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "6")
        el.set(qn("w:color"), color)
        pbdr.append(el)
    ppr.insert_element_before(pbdr, "w:shd", *_PPR_AFTER_SHD)


def _add_hyperlink(paragraph, url, text, bold=False, italic=False):
    part = paragraph.part
    try:
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
    except Exception:
        # 链接不合法就退化成「文字（URL）」，内容不丢
        run = paragraph.add_run("%s（%s）" % (text, url))
        run.bold, run.italic = bold, italic
        return
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    run_el = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(u)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    run_el.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run_el.append(t)
    hl.append(run_el)
    paragraph._p.append(hl)


def _shade_run(run, fill):
    rpr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    rpr.append(shd)


def _emit_runs(paragraph, text, base_size=None):
    for r in parse_inline(text):
        if r["link"]:
            _add_hyperlink(paragraph, r["link"], r["text"], r["bold"], r["italic"])
            continue
        run = paragraph.add_run(r["text"])
        run.bold = r["bold"]
        run.italic = r["italic"]
        if r["strike"]:
            run.font.strike = True
        if r["code"]:
            _set_run_fonts(run, CODE_FONT, EAST_ASIA_FONT)
            _shade_run(run, "F2F2F2")
        if base_size:
            run.font.size = Pt(base_size)
        elif r["code"]:
            run.font.size = Pt(BODY_SIZE - 1.0)


def _setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.size = Pt(BODY_SIZE)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    _set_style_fonts(normal, ASCII_FONT, EAST_ASIA_FONT)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.25

    for lvl, size in HEADING_SIZES.items():
        style = doc.styles["Heading %d" % lvl]
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.italic = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        _set_style_fonts(style, ASCII_FONT, EAST_ASIA_FONT)
        hpf = style.paragraph_format
        hpf.space_before = Pt(14 if lvl == 1 else 10 if lvl == 2 else 8)
        hpf.space_after = Pt(6)
        hpf.line_spacing = 1.2

    sec = doc.sections[0]
    sec.left_margin = Cm(2.6)
    sec.right_margin = Cm(2.6)
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.4)


def _tight_spacer(paragraph, points):
    """空段不占一整行：精确行高 + 1pt 字，视觉上只是一道缝。"""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(points)
    run = paragraph.add_run("")
    run.font.size = Pt(1)


def _add_table(doc, block):
    rows, aligns = block.rows, block.aligns
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = doc.styles["Table Grid"]
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # 表头跨页重复
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))

    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            paragraph = table.cell(r_idx, c_idx).paragraphs[0]
            fmt = paragraph.paragraph_format
            fmt.space_before = Pt(2)
            fmt.space_after = Pt(2)
            fmt.line_spacing = 1.15
            if aligns[c_idx] == "center":
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif aligns[c_idx] == "right":
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _emit_runs(paragraph, value, base_size=9.5)

    # 表格后补一点点空，避免和下一块黏在一起
    spacer = doc.add_paragraph()
    _tight_spacer(spacer, 6)


def build_docx(blocks) -> bytes:
    doc = Document()
    _setup_styles(doc)

    for b in blocks:
        if b.kind == "heading":
            p = doc.add_paragraph(style="Heading %d" % max(1, min(b.level, 6)))
            _emit_runs(p, b.text)
        elif b.kind == "para":
            p = doc.add_paragraph()
            _emit_runs(p, b.text)
        elif b.kind in ("bullet", "number"):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Cm(0.75 * (b.level + 1))
            pf.first_line_indent = Cm(-0.45)
            pf.space_after = Pt(3)
            if b.kind == "bullet":
                label = "• "
            elif b.marker.endswith("、"):
                label = b.marker
            else:
                label = b.marker + " "
            p.add_run(label)
            _emit_runs(p, b.text)
        elif b.kind == "code":
            keep = len(b.lines) <= 20
            for idx, line in enumerate(b.lines):
                p = doc.add_paragraph()
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                pf.left_indent = Cm(0.25)
                pf.right_indent = Cm(0.15)
                if keep and idx < len(b.lines) - 1:
                    pf.keep_with_next = True
                run = p.add_run(line if line else " ")
                _set_run_fonts(run, CODE_FONT, EAST_ASIA_FONT)
                run.font.size = Pt(CODE_SIZE)
                _box_paragraph(p)
                _shade_paragraph(p, "F5F5F5")
            tail = doc.add_paragraph()
            _tight_spacer(tail, 6)
        elif b.kind == "table":
            _add_table(doc, b)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# 8. 确定性：重写 zip，固定时间戳与顺序 => 同样输入字节级同样输出
# ---------------------------------------------------------------------------

_FIXED_TIME = (1980, 1, 1, 0, 0, 0)


def pack_deterministic(blob: bytes) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(blob))
    names = src.namelist()
    order = ["[Content_Types].xml"] + sorted(x for x in names if x != "[Content_Types].xml")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            info = zipfile.ZipInfo(name, date_time=_FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            info.create_system = 0
            zf.writestr(info, src.read(name))
    return out.getvalue()


# ---------------------------------------------------------------------------
# 9. 入口
# ---------------------------------------------------------------------------

def _read_input(path: str) -> str:
    if path == "-":
        return sys.stdin.buffer.read().decode("utf-8", errors="replace")
    with open(path, "rb") as fh:
        data = fh.read()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("utf-8", errors="replace")


def _stats(blocks) -> str:
    count = {}
    for b in blocks:
        count[b.kind] = count.get(b.kind, 0) + 1
    names = [("heading", "标题"), ("para", "段落"), ("bullet", "项目符号"),
             ("number", "编号列表"), ("table", "表格"), ("code", "代码块")]
    parts = ["%s %d" % (label, count[k]) for k, label in names if count.get(k)]
    return "、".join(parts) if parts else "（没有可识别结构）"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="clean_paste",
        description="把模型输出的原文清洗成可直接粘贴的 .docx（真标题/真列表/真表格/真代码块）。",
    )
    parser.add_argument("input", help="输入文本文件；用 - 表示从标准输入读")
    parser.add_argument("-o", "--out", default=None, help="输出目录（默认与输入同目录）")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    try:
        source = _read_input(args.input)
    except OSError as exc:
        sys.stderr.write("读不到输入文件：%s\n" % exc)
        return 1

    blocks = parse_blocks(source)
    if not blocks:
        sys.stderr.write("输入是空的，没有可转换的内容。\n")
        return 1

    blob = pack_deterministic(build_docx(blocks))

    if args.input == "-":
        stem, in_dir = "pasted", os.getcwd()
    else:
        stem = os.path.splitext(os.path.basename(args.input))[0]
        in_dir = os.path.dirname(os.path.abspath(args.input))
    out_dir = os.path.abspath(args.out) if args.out else in_dir
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, stem + ".clean.docx")
    with open(out_path, "wb") as fh:
        fh.write(blob)

    digest = hashlib.sha256(blob).hexdigest()[:16]
    print("输入：%s（%d 字符）" % (args.input, len(source)))
    print("识别：%s" % _stats(blocks))
    print("输出：%s（%d 字节，sha256:%s）" % (out_path, len(blob), digest))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
