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
        hps_raise = max(24, int(base * 0.6))
    return hps, hps_raise, hps_base_text


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

def split_okuri(surface, reading):
    okuri = ""
    i = len(surface)
    while i > 0:
        ch = surface[i - 1]
        cp = ord(ch)
        if 0x3041 <= cp <= 0x3096:
            okuri = ch + okuri
            i -= 1
        else:
            break
    kanji_part = surface[:i]
    okuri_part = okuri
    if not okuri_part:
        return surface, "", reading
    if reading.endswith(okuri_part):
        reading_kanji = reading[: len(reading) - len(okuri_part)]
    else:
        return surface, "", reading
    return kanji_part, okuri_part, reading_kanji

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
        kanji_part, okuri_part, reading_kanji = split_okuri(surface, hira)
        if not kanji_part or not contains_kanji(kanji_part):
            segments.append((surface, None))
            continue
        if reading_kanji and reading_kanji != kanji_part:
            segments.append((kanji_part, reading_kanji))
        else:
            segments.append((kanji_part, None))
        if okuri_part:
            segments.append((okuri_part, None))
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

# ────────────────────────────────────────────────
# 段落・ファイル処理
# ────────────────────────────────────────────────

def process_run(run, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black"):
    text = run.text
    if not text or not contains_kanji(text):
        return None
    sz_hpt, szCs_hpt = get_run_sz_szcs(run._r)
    rpr_elem = run._r.find(qn("w:rPr"))
    segments = get_ruby_segments(text, tok)
    new_elements = []
    for seg_text, reading in segments:
        if reading is not None:
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

def apply_ruby_to_paragraph(para, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black"):
    for run in para.runs:
        new_elems = process_run(run, tok, doc_default_hpt, theme_fonts, color_mode)
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

def apply_ruby_to_textboxes(doc, tok, doc_default_hpt=DEFAULT_BASE_TEXT_SIZE, theme_fonts=None, color_mode="black"):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    from docx.text.paragraph import Paragraph as DocxParagraph
    body = doc.element.body
    txbx_contents = body.findall(f".//{{{W}}}txbxContent")
    txbx_contents += body.findall(f".//{{{MC}}}Fallback//{{{W}}}txbxContent")
    for txbx in txbx_contents:
        for p_elem in txbx.findall(f"{{{W}}}p"):
            para = DocxParagraph(p_elem, doc)
            apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode)

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

def process_docx(file_bytes, filename, tok, color_mode="black"):
    doc = Document(io.BytesIO(file_bytes))
    doc_default_hpt = get_doc_default_font_size(doc)
    theme_fonts = get_theme_fonts(doc)
    for para in doc.paragraphs:
        apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    apply_ruby_to_paragraph(para, tok, doc_default_hpt, theme_fonts, color_mode)
    apply_ruby_to_textboxes(doc, tok, doc_default_hpt, theme_fonts, color_mode)
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

    # ルビ色の選択
    color_choice = st.radio(
        "ルビの文字色",
        options=["black", "match"],
        format_func=lambda x: "⬛ すべて黒色" if x == "black" else "🎨 本文の色に合わせる",
        horizontal=True,
    )

    if st.button("✨ ルビをふる", type="primary", use_container_width=True):
        with st.spinner("処理中です..."):
            try:
                file_bytes = uploaded_file.read()
                result = process_docx(file_bytes, uploaded_file.name, tok, color_mode=color_choice)

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
