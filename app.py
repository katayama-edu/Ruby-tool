#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ルビふりツール - Webアプリ版（Streamlit）
"""

import io
import sys
from datetime import datetime
from pathlib import Path
from copy import deepcopy

import streamlit as st
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from sudachipy import tokenizer, dictionary

# ────────────────────────────────────────────────
# ページ設定
# ────────────────────────────────────────────────

st.set_page_config(
    page_title="ルビふりツール",
    page_icon="📝",
    layout="centered",
)

st.title("📝 ルビふりツール")
st.caption("Word（.docx）ファイルの漢字にルビを自動付与します")

# ────────────────────────────────────────────────
# SudachiPy 初期化（キャッシュ）
# ────────────────────────────────────────────────

@st.cache_resource
def load_tokenizer():
    for dict_type in ("full", "core", "small"):
        try:
            tok = dictionary.Dictionary(dict_type=dict_type).create()
            return tok
        except Exception:
            continue
    raise RuntimeError("SudachiPy辞書の読み込みに失敗しました。")

# ────────────────────────────────────────────────
# ルビ書式パラメータ
# ────────────────────────────────────────────────

DEFAULT_BASE_TEXT_SIZE = 24
DEFAULT_RUBY_FONT = "游明朝"

# ────────────────────────────────────────────────
# 横書き：ルビ書式パラメータ
# ────────────────────────────────────────────────

def get_ruby_params(sz_hpt, szCs_hpt, doc_default_hpt=24):
    if sz_hpt is not None:
        base = sz_hpt
    elif szCs_hpt is not None:
        base = szCs_hpt
    else:
        base = doc_default_hpt
    hps_base_text = base
    hps = max(8, base // 2)
    if base <= 20:
        hps_raise = 20
    elif base <= 24:
        hps_raise = 24
    else:
        hps_raise = max(24, int(base * 1.0))
    return hps, hps_raise, hps_base_text

# ────────────────────────────────────────────────
# 縦書き：ルビ書式パラメータ
# ────────────────────────────────────────────────

def get_ruby_params_tate(sz_hpt, szCs_hpt, doc_default_hpt=24, hps_raise_tate=17):
    """
    縦書き用ルビパラメータ。実ファイル解析より：
      hpsBaseText = sz（またはフォールバック）
      hps         = ceil(szCs / 2)  szCsがNoneならsz//2
      hpsRaise    = 20固定
    """
    sz   = sz_hpt   if sz_hpt   is not None else doc_default_hpt
    szCs = szCs_hpt if szCs_hpt is not None else sz  # szCsがなければszで代用

    hps_base_text = sz
    hps = max(8, -(-szCs // 2))  # ceil(szCs/2)
    hps_raise = hps_raise_tate  # 縦書き：UIから調整可能
    rt_sz   = max(7, szCs // 3)
    rt_szCs = szCs
    return hps, hps_raise, hps_base_text, rt_sz, rt_szCs


def get_run_color(rpr_elem):
    """<w:rPr> から文字色を取得する。戻り値: 16進カラーコード文字列 or None"""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    if rpr_elem is None:
        return None
    color_elem = rpr_elem.find(f"{{{W}}}color")
    if color_elem is not None:
        val = color_elem.get(f"{{{W}}}val")
        if val and val.upper() != "AUTO":
            return val
    return None

# ────────────────────────────────────────────────
# 漢字判定
# ────────────────────────────────────────────────

def is_kanji(char):
    cp = ord(char)
    return (
        (0x4E00 <= cp <= 0x9FFF) or
        (0x3400 <= cp <= 0x4DBF) or
        (0x20000 <= cp <= 0x2A6DF) or
        (0xF900 <= cp <= 0xFAFF)
    )

def contains_kanji(text):
    return any(is_kanji(c) for c in text)

# ────────────────────────────────────────────────
# 読み取得・送り仮名分離
# ────────────────────────────────────────────────

def kata_to_hira(text):
    result = ""
    for ch in text:
        cp = ord(ch)
        if 0x30A1 <= cp <= 0x30F6:
            result += chr(cp - 0x60)
        else:
            result += ch
    return result

def is_non_kanji_char(ch):
    """漢字以外の文字か判定（ひらがな・カタカナ・英数字・記号など）"""
    return not is_kanji(ch)

def split_surface_reading(surface, reading):
    """
    表層形と読みを「漢字連続部分」と「非漢字部分」に分離し、
    それぞれに対応する読みを割り当てる。

    例:
      "炭酸水素ナトリウム" / "たんさんすいそナトリウム"
        → [("炭酸水素", "たんさんすいそ"), ("ナトリウム", None)]

      "書く" / "かく"
        → [("書", "か"), ("く", None)]

      "二酸化炭素" / "にさんかたんそ"
        → [("二酸化炭素", "にさんかたんそ")]
    """
    # surface を「漢字連続」と「非漢字連続」に分割
    segments_surf = []
    i = 0
    while i < len(surface):
        if is_kanji(surface[i]):
            j = i
            while j < len(surface) and is_kanji(surface[j]):
                j += 1
            segments_surf.append((surface[i:j], True))
            i = j
        else:
            j = i
            while j < len(surface) and not is_kanji(surface[j]):
                j += 1
            segments_surf.append((surface[i:j], False))
            i = j

    # 非漢字部分が読みに含まれているか確認して読みを分配する
    result = []
    remaining_reading = reading

    for seg_text, is_k in segments_surf:
        if not is_k:
            # 非漢字部分（カタカナ・ひらがな等）
            # 読みの先頭または末尾から対応部分を除去
            seg_hira = kata_to_hira(seg_text)
            if remaining_reading.startswith(seg_hira):
                remaining_reading = remaining_reading[len(seg_hira):]
            elif remaining_reading.endswith(seg_hira):
                remaining_reading = remaining_reading[: len(remaining_reading) - len(seg_hira)]
            result.append((seg_text, None))
        else:
            # 漢字部分：残りの読みのうち次の非漢字セグメントの読みを除いた部分を割り当て
            # 次の非漢字セグメントを先読みして読みの末尾から除去
            next_non_kanji = ""
            idx = segments_surf.index((seg_text, True))
            for next_seg, next_is_k in segments_surf[idx + 1:]:
                if not next_is_k:
                    next_non_kanji = kata_to_hira(next_seg)
                    break

            if next_non_kanji and remaining_reading.endswith(next_non_kanji):
                kanji_reading = remaining_reading[: len(remaining_reading) - len(next_non_kanji)]
                remaining_reading = next_non_kanji
            else:
                kanji_reading = remaining_reading
                remaining_reading = ""

            if kanji_reading and kanji_reading != seg_text:
                result.append((seg_text, kanji_reading))
            else:
                result.append((seg_text, None))

    return result


def get_ruby_segments(text, tok):
    mode = tokenizer.Tokenizer.SplitMode.C
    morphemes = tok.tokenize(text, mode)
    segments = []

    for m in morphemes:
        surface = m.surface()
        if not contains_kanji(surface):
            segments.append((surface, None))
            continue
        reading = m.reading_form()
        if not reading:
            segments.append((surface, None))
            continue
        hira = kata_to_hira(reading)

        # 漢字と非漢字が混在する場合（例：炭酸水素ナトリウム）を分離
        sub_segs = split_surface_reading(surface, hira)
        segments.extend(sub_segs)

    return segments

# ────────────────────────────────────────────────
# フォントサイズ・情報取得
# ────────────────────────────────────────────────

def get_run_sz_szcs(run_elem):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    def read_pair(rpr_elem):
        if rpr_elem is None:
            return None, None
        sz = szCs = None
        sz_e = rpr_elem.find(f"{{{W}}}sz")
        if sz_e is not None:
            v = sz_e.get(f"{{{W}}}val")
            if v:
                sz = int(v)
        szCs_e = rpr_elem.find(f"{{{W}}}szCs")
        if szCs_e is not None:
            v = szCs_e.get(f"{{{W}}}val")
            if v:
                szCs = int(v)
        return sz, szCs
    rpr = run_elem.find(f"{{{W}}}rPr")
    sz, szCs = read_pair(rpr)
    if sz is not None or szCs is not None:
        return sz, szCs
    parent = run_elem.getparent()
    if parent is not None:
        ppr = parent.find(f"{{{W}}}pPr")
        if ppr is not None:
            sz, szCs = read_pair(ppr.find(f"{{{W}}}rPr"))
            if sz is not None or szCs is not None:
                return sz, szCs
    return None, None

def get_theme_fonts(doc):
    """
    theme1.xml から日本語フォント（majorFont/minorFont の Jpan）を取得する。
    戻り値: {'major': str or None, 'minor': str or None}
    """
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    result = {'major': None, 'minor': None}
    try:
        theme_part = doc.part.part_related_by(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
        )
        root = etree.fromstring(theme_part.blob)
        for kind in ('major', 'minor'):
            tag = 'majorFont' if kind == 'major' else 'minorFont'
            font_elem = root.find(f".//{{{A}}}{tag}")
            if font_elem is not None:
                for f in font_elem.findall(f"{{{A}}}font"):
                    if f.get('script') == 'Jpan':
                        result[kind] = f.get('typeface')
                        break
                # Jpanがなければlatinを使用
                if result[kind] is None:
                    latin = font_elem.find(f"{{{A}}}latin")
                    if latin is not None:
                        result[kind] = latin.get('typeface')
    except Exception:
        pass
    return result


def get_run_font_info(rpr_elem, theme_fonts=None):
    """
    <w:rPr> からフォント・kern・spacing 情報を返す。
    eastAsiaTheme が指定されている場合はテーマフォントに解決する。
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    info = {'eastAsia': None, 'ascii': None, 'hAnsi': None, 'kern': None, 'spacing': None}
    if rpr_elem is None:
        return info
    rfonts = rpr_elem.find(f"{{{W}}}rFonts")
    if rfonts is not None:
        ea = rfonts.get(f"{{{W}}}eastAsia")
        ea_theme = rfonts.get(f"{{{W}}}eastAsiaTheme")
        # eastAsia が直接指定されている場合はそのまま使用
        if ea:
            info['eastAsia'] = ea
        # eastAsiaTheme でテーマフォントを参照している場合は解決する
        elif ea_theme and theme_fonts:
            if 'minor' in ea_theme.lower():
                info['eastAsia'] = theme_fonts.get('minor')
            elif 'major' in ea_theme.lower():
                info['eastAsia'] = theme_fonts.get('major')
        info['ascii'] = rfonts.get(f"{{{W}}}ascii")
        info['hAnsi'] = rfonts.get(f"{{{W}}}hAnsi")
    kern = rpr_elem.find(f"{{{W}}}kern")
    if kern is not None:
        info['kern'] = kern.get(f"{{{W}}}val")
    spacing = rpr_elem.find(f"{{{W}}}spacing")
    if spacing is not None:
        info['spacing'] = spacing.get(f"{{{W}}}val")
    return info

# ────────────────────────────────────────────────
# ルビ要素生成
# ────────────────────────────────────────────────

def make_ruby_element(base_text, ruby_text, sz_hpt, szCs_hpt, doc_default_hpt, rpr_elem=None, theme_fonts=None, color_mode="black"):
    hps, hps_raise, hps_base_text = get_ruby_params(sz_hpt, szCs_hpt, doc_default_hpt)
    font_info = get_run_font_info(rpr_elem, theme_fonts)

    ea_font = font_info['eastAsia'] or DEFAULT_RUBY_FONT
    ruby_font  = ea_font
    ruby_ascii = font_info['ascii'] or ea_font
    ruby_hansi = font_info['hAnsi'] or ea_font

    # ルビ色の決定
    ruby_color = None
    if color_mode == "match":
        ruby_color = get_run_color(rpr_elem)  # 本文と同じ色（Noneなら黒）

    ruby = OxmlElement("w:ruby")
    ruby_pr = OxmlElement("w:rubyPr")
    ruby_align = OxmlElement("w:rubyAlign")
    ruby_align.set(qn("w:val"), "distributeSpace")
    ruby_pr.append(ruby_align)
    hps_elem = OxmlElement("w:hps")
    hps_elem.set(qn("w:val"), str(hps))
    ruby_pr.append(hps_elem)
    hps_raise_elem = OxmlElement("w:hpsRaise")
    hps_raise_elem.set(qn("w:val"), str(hps_raise))
    ruby_pr.append(hps_raise_elem)
    hps_base_elem = OxmlElement("w:hpsBaseText")
    hps_base_elem.set(qn("w:val"), str(hps_base_text))
    ruby_pr.append(hps_base_elem)
    lid = OxmlElement("w:lid")
    lid.set(qn("w:val"), "ja-JP")
    ruby_pr.append(lid)
    ruby.append(ruby_pr)

    rt = OxmlElement("w:rt")
    rt_run = OxmlElement("w:r")
    rt_rpr = OxmlElement("w:rPr")
    rt_rfonts = OxmlElement("w:rFonts")
    rt_rfonts.set(qn("w:ascii"),    ruby_ascii)
    rt_rfonts.set(qn("w:hAnsi"),    ruby_hansi)
    rt_rfonts.set(qn("w:eastAsia"), ruby_font)
    rt_rpr.append(rt_rfonts)
    if font_info['kern'] is not None:
        rt_kern = OxmlElement("w:kern")
        rt_kern.set(qn("w:val"), font_info['kern'])
        rt_rpr.append(rt_kern)
    if font_info['spacing'] is not None:
        rt_spacing = OxmlElement("w:spacing")
        rt_spacing.set(qn("w:val"), font_info['spacing'])
        rt_rpr.append(rt_spacing)
    # 色指定
    if ruby_color:
        rt_color = OxmlElement("w:color")
        rt_color.set(qn("w:val"), ruby_color)
        rt_rpr.append(rt_color)
    rt_sz = OxmlElement("w:sz")
    rt_sz.set(qn("w:val"), str(hps))
    rt_rpr.append(rt_sz)
    rt_sz_cs = OxmlElement("w:szCs")
    rt_sz_cs.set(qn("w:val"), str(hps))
    rt_rpr.append(rt_sz_cs)
    rt_run.append(rt_rpr)
    rt_t = OxmlElement("w:t")
    rt_t.text = ruby_text
    if ruby_text.startswith(" ") or ruby_text.endswith(" "):
        rt_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    rt_run.append(rt_t)
    rt.append(rt_run)
    ruby.append(rt)

    ruby_base_elem = OxmlElement("w:rubyBase")
    base_run = OxmlElement("w:r")
    if rpr_elem is not None:
        try:
            base_run.append(deepcopy(rpr_elem))
        except Exception:
            pass
    base_t = OxmlElement("w:t")
    base_t.text = base_text
    if base_text.startswith(" ") or base_text.endswith(" "):
        base_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    base_run.append(base_t)
    ruby_base_elem.append(base_run)
    ruby.append(ruby_base_elem)
    return ruby


def make_ruby_element_tate(base_text, ruby_text, sz_hpt, szCs_hpt, doc_default_hpt, rpr_elem=None, theme_fonts=None, color_mode="black", hps_raise_tate=17):
    """
    縦書き用ルビ要素生成。
    実ファイル解析より：rt に sz/szCs を指定しない（Wordがhpsから自動計算）
    hpsBaseText = sz、hps = ceil(szCs/2)、hpsRaise = UIで調整可能
    """
    hps, hps_raise, hps_base_text, rt_sz, rt_szCs = get_ruby_params_tate(sz_hpt, szCs_hpt, doc_default_hpt, hps_raise_tate)
    font_info = get_run_font_info(rpr_elem, theme_fonts)

    ea_font    = font_info['eastAsia'] or DEFAULT_RUBY_FONT
    ruby_font  = ea_font
    ruby_ascii = font_info['ascii'] or ea_font
    ruby_hansi = font_info['hAnsi'] or ea_font

    ruby_color = None
    if color_mode == "match":
        ruby_color = get_run_color(rpr_elem)

    ruby = OxmlElement("w:ruby")
    ruby_pr = OxmlElement("w:rubyPr")
    ruby_align = OxmlElement("w:rubyAlign")
    ruby_align.set(qn("w:val"), "distributeSpace")
    ruby_pr.append(ruby_align)
    hps_elem = OxmlElement("w:hps")
    hps_elem.set(qn("w:val"), str(hps))
    ruby_pr.append(hps_elem)
    hps_raise_elem = OxmlElement("w:hpsRaise")
    hps_raise_elem.set(qn("w:val"), str(hps_raise))
    ruby_pr.append(hps_raise_elem)
    hps_base_elem = OxmlElement("w:hpsBaseText")
    hps_base_elem.set(qn("w:val"), str(hps_base_text))
    ruby_pr.append(hps_base_elem)
    lid = OxmlElement("w:lid")
    lid.set(qn("w:val"), "ja-JP")
    ruby_pr.append(lid)
    ruby.append(ruby_pr)

    rt = OxmlElement("w:rt")
    rt_run = OxmlElement("w:r")
    rt_rpr = OxmlElement("w:rPr")
    rt_rfonts = OxmlElement("w:rFonts")
    rt_rfonts.set(qn("w:ascii"),    ruby_ascii)
    rt_rfonts.set(qn("w:hAnsi"),    ruby_hansi)
    rt_rfonts.set(qn("w:eastAsia"), ruby_font)
    rt_rpr.append(rt_rfonts)
    if ruby_color:
        rt_color = OxmlElement("w:color")
        rt_color.set(qn("w:val"), ruby_color)
        rt_rpr.append(rt_color)
    # 縦書きは sz/szCs を指定しない（Wordがhpsから自動計算）
    rt_run.append(rt_rpr)
    rt_t = OxmlElement("w:t")
    rt_t.text = ruby_text
    if ruby_text.startswith(" ") or ruby_text.endswith(" "):
        rt_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    rt_run.append(rt_t)
    rt.append(rt_run)
    ruby.append(rt)

    ruby_base_elem = OxmlElement("w:rubyBase")
    base_run = OxmlElement("w:r")
    if rpr_elem is not None:
        try:
            base_run.append(deepcopy(rpr_elem))
        except Exception:
            pass
    base_t = OxmlElement("w:t")
    base_t.text = base_text
    if base_text.startswith(" ") or base_text.endswith(" "):
        base_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    base_run.append(base_t)
    ruby_base_elem.append(base_run)
    ruby.append(ruby_base_elem)
    return ruby


def detect_text_direction(file_bytes):
    """
    docx ファイルの書き方向を自動検出する。
    sectPr の textDirection が tbRl なら縦書き。
    戻り値: 'tate' or 'yoko'
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    try:
        doc = Document(io.BytesIO(file_bytes))
        body = doc.element.body
        sect = body.find(f"{{{W}}}sectPr")
        if sect is not None:
            td = sect.find(f"{{{W}}}textDirection")
            if td is not None:
                val = td.get(f"{{{W}}}val", "")
                if val in ("tbRl", "tbLr", "btLr"):
                    return "tate"
    except Exception:
        pass
    return "yoko"

# ────────────────────────────────────────────────
# 段落・ファイル処理
# ────────────────────────────────────────────────

def process_run(run, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black", tate=False, hps_raise_tate=17):
    text = run.text
    if not text or not contains_kanji(text):
        return None
    sz_hpt, szCs_hpt = get_run_sz_szcs(run._r)
    rpr_elem = run._r.find(qn("w:rPr"))
    segments = get_ruby_segments(text, tok)
    new_elements = []
    for seg_text, reading in segments:
        if reading is not None:
            if tate:
                ruby_elem = make_ruby_element_tate(seg_text, reading, sz_hpt, szCs_hpt, doc_default_hpt, rpr_elem, theme_fonts, color_mode, hps_raise_tate)
            else:
                ruby_elem = make_ruby_element(seg_text, reading, sz_hpt, szCs_hpt, doc_default_hpt, rpr_elem, theme_fonts, color_mode)
            new_elements.append(ruby_elem)
        else:
            plain_run = deepcopy(run._r)
            for t in plain_run.findall(qn("w:t")):
                t.text = seg_text
                if seg_text.startswith(" ") or seg_text.endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            new_elements.append(plain_run)
    return new_elements

def apply_ruby_to_paragraph(para, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black", tate=False, hps_raise_tate=17):
    for run in para.runs:
        new_elems = process_run(run, tok, doc_default_hpt, theme_fonts, color_mode, tate, hps_raise_tate)
        if new_elems is None:
            continue
        r_elem = run._r
        parent = r_elem.getparent()
        if parent is None:
            continue
        idx = list(parent).index(r_elem)
        parent.remove(r_elem)
        for i, elem in enumerate(new_elems):
            parent.insert(idx + i, elem)

def apply_ruby_to_textboxes(doc, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black", tate=False, hps_raise_tate=17):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    from docx.text.paragraph import Paragraph as DocxParagraph
    body = doc.element.body
    txbx_contents = body.findall(f".//{{{W}}}txbxContent")
    txbx_contents += body.findall(f".//{{{MC}}}Fallback//{{{W}}}txbxContent")
    for txbx in txbx_contents:
        for p_elem in txbx.findall(f"{{{W}}}p"):
            para = DocxParagraph(p_elem, doc)
            apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode, tate, hps_raise_tate)

def get_doc_default_font_size(doc):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    try:
        root = doc.part.styles._element
        sz = root.find(f".//{{{W}}}rPrDefault/{{{W}}}rPr/{{{W}}}sz")
        if sz is not None:
            val = sz.get(f"{{{W}}}val")
            if val:
                return int(val)
    except Exception:
        pass
    return DEFAULT_BASE_TEXT_SIZE

def process_docx(file_bytes, filename, tok, color_mode="black", tate=False, hps_raise_tate=17):
    doc = Document(io.BytesIO(file_bytes))
    doc_default_hpt = get_doc_default_font_size(doc)
    theme_fonts = get_theme_fonts(doc)
    for para in doc.paragraphs:
        apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode, tate, hps_raise_tate)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode, tate, hps_raise_tate)
    apply_ruby_to_textboxes(doc, tok, doc_default_hpt, theme_fonts, color_mode, tate, hps_raise_tate)
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ────────────────────────────────────────────────
# UI
# ────────────────────────────────────────────────

# 辞書の読み込み
with st.spinner("辞書を読み込んでいます（初回のみ時間がかかります）..."):
    try:
        tok = load_tokenizer()
    except Exception as e:
        st.error(f"辞書の読み込みに失敗しました: {e}")
        st.stop()

st.success("辞書の読み込み完了！")
st.divider()

uploaded_file = st.file_uploader(
    "Wordファイル（.docx）をアップロードしてください",
    type=["docx"],
    help=".docx 形式のファイルのみ対応しています"
)

if uploaded_file is not None:
    st.info(f"📄 ファイル名: {uploaded_file.name}")

    # 書き方向の自動検出
    file_bytes_preview = uploaded_file.read()
    auto_direction = detect_text_direction(file_bytes_preview)
    uploaded_file.seek(0)  # 読み直しのためにリセット

    direction_default = 0 if auto_direction == "yoko" else 1
    st.caption(f"🔍 自動検出: {'横書き' if auto_direction == 'yoko' else '縦書き'}")

    # 書き方向の選択
    direction_choice = st.radio(
        "書き方向",
        options=["yoko", "tate"],
        format_func=lambda x: "↔️ 横書き" if x == "yoko" else "↕️ 縦書き",
        index=direction_default,
        horizontal=True,
    )

    # ルビ色の選択
    color_choice = st.radio(
        "ルビの文字色",
        options=["black", "match"],
        format_func=lambda x: "⬛ すべて黒色" if x == "black" else "🎨 本文の色に合わせる",
        horizontal=True,
    )

    # 縦書きのみ：ルビ距離スライダー
    hps_raise_tate = 17
    if direction_choice == "tate":
        offset = st.slider(
            "ルビの距離調整（縦書き用）",
            min_value=-5,
            max_value=5,
            value=0,
            step=1,
            help="0がデフォルト  \n＋で離れる　－で近づく"
        )
        hps_raise_tate = 17 + offset

    if st.button("✨ ルビをふる", type="primary", use_container_width=True):
        with st.spinner("処理中です..."):
            try:
                file_bytes = uploaded_file.read()
                tate = (direction_choice == "tate")
                result = process_docx(file_bytes, uploaded_file.name, tok, color_mode=color_choice, tate=tate, hps_raise_tate=hps_raise_tate)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = Path(uploaded_file.name).stem
                out_name = f"ルビ版_{timestamp}_{stem}.docx"

                st.success("✅ 処理完了！")
                st.download_button(
                    label="📥 ダウンロード",
                    data=result,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")

st.divider()
st.caption("ルビの読みはSudachiPy（全辞書）を使用しています。付与後に内容をご確認ください。")

st.markdown("""
<style>
.rubifuri-kun {
    position: fixed;
    bottom: 16px;
    right: 16px;
    width: 80px;
    opacity: 0.88;
    z-index: 999;
    pointer-events: none;
}
</style>
<div class="rubifuri-kun">
<svg width="100%" viewBox="0 0 680 500" xmlns="http://www.w3.org/2000/svg">
  <line x1="308" y1="168" x2="295" y2="205" stroke="#555" stroke-width="2" stroke-dasharray="4,3"/>
  <line x1="372" y1="168" x2="385" y2="205" stroke="#555" stroke-width="2" stroke-dasharray="4,3"/>
  <rect x="225" y="205" width="230" height="210" rx="28" fill="#FFD93D" stroke="#333" stroke-width="3.5"/>
  <path d="M228 270 Q175 280 168 330" fill="none" stroke="#FFD93D" stroke-width="20" stroke-linecap="round"/>
  <path d="M228 270 Q175 280 168 330" fill="none" stroke="#333" stroke-width="3.5" stroke-linecap="round"/>
  <circle cx="167" cy="337" r="16" fill="#FFD93D" stroke="#333" stroke-width="3.5"/>
  <path d="M452 260 Q500 230 510 185" fill="none" stroke="#FFD93D" stroke-width="20" stroke-linecap="round"/>
  <path d="M452 260 Q500 230 510 185" fill="none" stroke="#333" stroke-width="3.5" stroke-linecap="round"/>
  <circle cx="512" cy="178" r="16" fill="#FFD93D" stroke="#333" stroke-width="3.5"/>
  <ellipse cx="300" cy="278" rx="20" ry="22" fill="#333"/>
  <ellipse cx="380" cy="278" rx="20" ry="22" fill="#333"/>
  <ellipse cx="307" cy="271" rx="7" ry="8" fill="white"/>
  <ellipse cx="387" cy="271" rx="7" ry="8" fill="white"/>
  <ellipse cx="272" cy="308" rx="20" ry="13" fill="#FF9A8B" opacity="0.55"/>
  <ellipse cx="408" cy="308" rx="20" ry="13" fill="#FF9A8B" opacity="0.55"/>
  <path d="M302 328 Q340 355 378 328" fill="none" stroke="#333" stroke-width="3.5" stroke-linecap="round"/>
  <rect x="268" y="408" width="42" height="62" rx="21" fill="#FFD93D" stroke="#333" stroke-width="3.5"/>
  <rect x="370" y="408" width="42" height="62" rx="21" fill="#FFD93D" stroke="#333" stroke-width="3.5"/>
  <ellipse cx="289" cy="469" rx="30" ry="16" fill="#333"/>
  <ellipse cx="391" cy="469" rx="30" ry="16" fill="#333"/>
  <circle cx="340" cy="108" r="62" fill="#FF6B9D" stroke="#333" stroke-width="3.5"/>
  <ellipse cx="320" cy="97" rx="11" ry="13" fill="#333"/>
  <ellipse cx="360" cy="97" rx="11" ry="13" fill="#333"/>
  <ellipse cx="326" cy="91" rx="4.5" ry="5.5" fill="white"/>
  <ellipse cx="366" cy="91" rx="4.5" ry="5.5" fill="white"/>
  <ellipse cx="303" cy="112" rx="12" ry="8" fill="#FF9A8B" opacity="0.55"/>
  <ellipse cx="377" cy="112" rx="12" ry="8" fill="#FF9A8B" opacity="0.55"/>
  <path d="M320 125 Q340 140 360 125" fill="none" stroke="#333" stroke-width="2.5" stroke-linecap="round"/>
  <ellipse cx="272" cy="118" rx="20" ry="14" fill="#FF6B9D" stroke="#333" stroke-width="3"/>
  <ellipse cx="410" cy="95" rx="20" ry="14" fill="#FF6B9D" stroke="#333" stroke-width="3" transform="rotate(-25 410 95)"/>
</svg>
</div>
""", unsafe_allow_html=True)
