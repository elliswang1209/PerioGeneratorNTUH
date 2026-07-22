# pptx_engine.py
"""
簡報生成引擎：完全匹配專科規範之順序與對稱佈局。
Upper: Mobility -> KM -> PD(B) -> Rec(B) -> CAL(B) -> Furcation -> PD(P) -> Rec(P) -> CAL(P)
Lower: PD(L) -> Rec(L) -> CAL(L) -> Mobility -> KM -> PD(B) -> Rec(B) -> CAL(B) -> Furcation
"""
from io import BytesIO
from typing import Dict, Set, Any, List
import pandas as pd

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

import config
from core_parser import (
    find_tooth_rows, 
    get_tooth_start_columns, 
    collect_comparison_row_indices, 
    get_three_digit_raw_list,
    clean_cell,
    find_furcation_rows
)

def _rgb(color_tuple):
    return RGBColor(*color_tuple)

def _apply_cell_density(cell):
    cell.margin_top = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_bottom = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_left = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_right = Inches(config.CELL_INTERNAL_MARGIN)

def _set_cell_border_to_white(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    border_xml = (
        f'<a:ln {nsdecls("a")} w="12700" cmpd="s" algn="ctr">'
        f'  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'  <a:prstDash val="solid"/>'
        f'</a:ln>'
    )
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{side}')
        if existing is not None:
            tcPr.remove(existing)
        tcPr.append(parse_xml(f'<a:{side} {nsdecls("a")}>{border_xml}</a:{side}>'))

def _add_diagonal_strikethrough(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    diagonal_xml = (
        f'<a:lnDiagonalDown {nsdecls("a")} w="12700" cmpd="s" algn="ctr">'
        f'  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'  <a:prstDash val="solid"/>'
        f'</a:lnDiagonalDown>'
    )
    existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}lnDiagonalDown')
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(parse_xml(diagonal_xml))

def _remove_default_table_style(table):
    try:
        tblPr = table._tbl.tblPr
        tableStyleId = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
        if tableStyleId is not None:
            tblPr.remove(tableStyleId)
        tblPr.append(parse_xml(f'<a:tableStyleId {nsdecls("a")}>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>'))
    except Exception:
        pass

def create_six_sextants_presentation(df, missing_teeth: Set[int]) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name}", teeth, df, missing_teeth, is_comparison=False)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

def create_comparison_presentation(df, missing_teeth: Set[int]) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name}", teeth, df, missing_teeth, is_comparison=False)

    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name}", teeth, df, missing_teeth, is_comparison=True)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

def _draw_sextant_slide(slide, title_text: str, teeth: List[int], df, missing_teeth: Set[int], is_comparison: bool):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(config.COLOR_BG_DARK)

    # 1. 標題
    tx_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(12), Inches(0.8))
    p = tx_box.text_frame.paragraphs[0]
    p.text = title_text
    p.font.name = config.FONT_PRIMARY
    p.font.size = Pt(config.FONT_SIZE_TITLE)
    p.font.bold = True
    p.font.italic = True
    p.font.underline = True
    p.font.color.rgb = _rgb(config.COLOR_TITLE_WHEAT)
    p.alignment = PP_ALIGN.LEFT

    tooth_rows = find_tooth_rows(df)
    if len(tooth_rows) < 2: return
    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    row_map = collect_comparison_row_indices(df)
    is_upper = (teeth[0] < 30)

    # 2. 🚀 完全匹配指定的 18/17 列項目與順序
    if not is_comparison:
        active_row_metas = [
            {"label": "Probing Depth (B)" if is_upper else "Probing Depth (L)", "r_key": "up_b_pd_i" if is_upper else "lo_l_pd_i", "type": "pd"},
            {"label": "Probing Depth (P)" if is_upper else "Probing Depth (B)", "r_key": "up_p_pd_i" if is_upper else "lo_b_pd_i", "type": "pd"},
            {"label": "Recession", "r_key": "up_b_gm_i" if is_upper else "lo_l_gm_i", "type": "gm"},
            {"label": "", "r_key": "up_p_gm_i" if is_upper else "lo_b_gm_i", "type": "gm"},
            {"label": "Masticatory Mucosa", "r_key": "up_km_i" if is_upper else "lo_km_i", "type": "km"},
            {"label": "Furcation Involvement", "type": "furc"},
            {"label": "Mobility", "r_key": "up_mob_i" if is_upper else "lo_mob_i", "type": "mob"},
            {"label": "Prognosis", "type": "prog"}
        ]
    else:
        if is_upper:
            # 🚀 上顎 Initial + Re-eval (17 列)
            active_row_metas = [
                {"label": "Mobility (I)", "r_key": "up_mob_i", "type": "mob"},
                {"label": "Mobility (R)", "r_key": "up_mob_r", "type": "mob"},
                {"label": "Masticatory Mucosa (I)", "r_key": "up_km_i", "type": "km"},
                {"label": "Masticatory Mucosa (R)", "r_key": "up_km_r", "type": "km"},
                {"label": "Probing Depth (B) (I)", "r_key": "up_b_pd_i", "type": "pd"},
                {"label": "Probing Depth (B) (R)", "r_key": "up_b_pd_r", "type": "pd"},
                {"label": "Recession (B) (I)", "r_key": "up_b_gm_i", "type": "gm"},
                {"label": "Recession (B) (R)", "r_key": "up_b_gm_r", "type": "gm"},
                {"label": "CAL (B) (I)", "r_key": "up_b_cal_i", "type": "cal"},
                {"label": "CAL (B) (R)", "r_key": "up_b_cal_r", "type": "cal"},
                {"label": "Furcation Involvement", "type": "furc"},
                {"label": "Probing Depth (P) (I)", "r_key": "up_p_pd_i", "type": "pd"},
                {"label": "Probing Depth (P) (R)", "r_key": "up_p_pd_r", "type": "pd"},
                {"label": "Recession (P) (I)", "r_key": "up_p_gm_i", "type": "gm"},
                {"label": "Recession (P) (R)", "r_key": "up_p_gm_r", "type": "gm"},
                {"label": "CAL (P) (I)", "r_key": "up_p_cal_i", "type": "cal"},
                {"label": "CAL (P) (R)", "r_key": "up_p_cal_r", "type": "cal"},
            ]
        else:
            # 🚀 下顎 Initial + Re-eval (17 列)
            active_row_metas = [
                {"label": "Probing Depth (L) (I)", "r_key": "lo_l_pd_i", "type": "pd"},
                {"label": "Probing Depth (L) (R)", "r_key": "lo_l_pd_r", "type": "pd"},
                {"label": "Recession (L) (I)", "r_key": "lo_l_gm_i", "type": "gm"},
                {"label": "Recession (L) (R)", "r_key": "lo_l_gm_r", "type": "gm"},
                {"label": "CAL (L) (I)", "r_key": "lo_l_cal_i", "type": "cal"},
                {"label": "CAL (L) (R)", "r_key": "lo_l_cal_r", "type": "cal"},
                {"label": "Mobility (I)", "r_key": "lo_mob_i", "type": "mob"},
                {"label": "Mobility (R)", "r_key": "lo_mob_r", "type": "mob"},
                {"label": "Masticatory Mucosa (I)", "r_key": "lo_km_i", "type": "km"},
                {"label": "Masticatory Mucosa (R)", "r_key": "lo_km_r", "type": "km"},
                {"label": "Probing Depth (B) (I)", "r_key": "lo_b_pd_i", "type": "pd"},
                {"label": "Probing Depth (B) (R)", "r_key": "lo_b_pd_r", "type": "pd"},
                {"label": "Recession (B) (I)", "r_key": "lo_b_gm_i", "type": "gm"},
                {"label": "Recession (B) (R)", "r_key": "lo_b_gm_r", "type": "gm"},
                {"label": "CAL (B) (I)", "r_key": "lo_b_cal_i", "type": "cal"},
                {"label": "CAL (B) (R)", "r_key": "lo_b_cal_r", "type": "cal"},
                {"label": "Furcation Involvement", "type": "furc"},
            ]

    total_rows = 1 + len(active_row_metas)
    total_cols = 1 + len(teeth)

    if not is_comparison:
        col_width_left, col_width_data = Inches(2.2), Inches(1.8)
        row_height, top_pos = Inches(config.TABLE_ROW_HEIGHT), Inches(1.3)
    else:
        col_width_left, col_width_data = Inches(2.1), Inches(0.85)
        row_height, top_pos = Inches(0.32), Inches(1.2)

    total_table_width = col_width_left + col_width_data * len(teeth)
    total_table_height = row_height * total_rows
    left_pos = Inches(config.PPT_SLIDE_WIDTH) - total_table_width - Inches(0.4)

    table_shape = slide.shapes.add_table(total_rows, total_cols, left_pos, top_pos, total_table_width, total_table_height)
    table = table_shape.table
    _remove_default_table_style(table)

    table.columns[0].width = col_width_left
    for c in range(1, total_cols): table.columns[c].width = col_width_data
    for r in range(total_rows): table.rows[r].height = row_height

    text_white, text_alert = _rgb(config.COLOR_TEXT_WHITE), _rgb(config.COLOR_TEXT_ALERT)

    # 3. Header
    c_t0 = table.cell(0, 0); c_t0.vertical_anchor = MSO_ANCHOR.MIDDLE
    _apply_cell_density(c_t0); c_t0.fill.background()
    p_t0 = c_t0.text_frame.paragraphs[0]; p_t0.alignment = PP_ALIGN.CENTER
    r_t0 = p_t0.add_run(); r_t0.text = "Tooth"
    r_t0.font.name, r_t0.font.size, r_t0.font.bold, r_t0.font.color.rgb = config.FONT_PRIMARY, Pt(11 if is_comparison else 14), True, text_white

    for c_i, t in enumerate(teeth):
        c_t = table.cell(0, c_i + 1); c_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        _apply_cell_density(c_t); c_t.fill.background()
        p_t = c_t.text_frame.paragraphs[0]; p_t.alignment = PP_ALIGN.CENTER
        r_t = p_t.add_run(); r_t.text = str(t)
        r_t.font.name, r_t.font.size, r_t.font.bold, r_t.font.color.rgb = config.FONT_PRIMARY, Pt(11 if is_comparison else 14), True, text_white

    f_rows = find_furcation_rows(df)
    dynamic_missing = set(missing_teeth)

    # 4. 填寫資料列
    for r_i, meta in enumerate(active_row_metas):
        r_ppt = r_i + 1
        c_lbl = table.cell(r_ppt, 0); c_lbl.vertical_anchor = MSO_ANCHOR.MIDDLE
        _apply_cell_density(c_lbl); c_lbl.fill.background()
        
        lbl_text = meta["label"]
        if lbl_text:
            p_lbl = c_lbl.text_frame.paragraphs[0]; p_lbl.alignment = PP_ALIGN.CENTER
            r_lbl = p_lbl.add_run(); r_lbl.text = lbl_text
            r_lbl.font.name, r_lbl.font.size, r_lbl.font.bold, r_lbl.font.color.rgb = config.FONT_PRIMARY, Pt(9.5 if is_comparison else 12), True, text_white

        for c_i, t in enumerate(teeth):
            cell_d = table.cell(r_ppt, c_i + 1); cell_d.vertical_anchor = MSO_ANCHOR.MIDDLE
            _apply_cell_density(cell_d); cell_d.fill.background()

            col = upper_cols.get(t) if is_upper else lower_cols.get(t)
            p_d = cell_d.text_frame.paragraphs[0]; p_d.alignment = PP_ALIGN.CENTER

            if col is not None:
                m_type = meta["type"]
                r_idx = row_map.get(meta.get("r_key"))

                if m_type in ["pd", "gm", "cal"]:
                    digits = get_three_digit_raw_list(df, r_idx, col)
                    if m_type == "pd" and all(d.strip() in ['?', ''] for d in digits):
                        dynamic_missing.add(t)

                    for d in digits:
                        display_val = clean_cell(d)
                        run = p_d.add_run(); run.text = f" {display_val} " if len(display_val)>=2 else display_val
                        run.font.name, run.font.size = config.FONT_PRIMARY, Pt(10 if is_comparison else 14)
                        if m_type == "pd" and display_val.isdigit() and int(display_val) >= 5:
                            run.font.bold = True; run.font.color.rgb = text_alert
                        else:
                            run.font.bold = False; run.font.color.rgb = text_white

                elif m_type == "km":
                    val = clean_cell(df.iloc[r_idx, col]) if r_idx is not None else "0"
                    run = p_d.add_run(); run.text = val if val!="" else "0"
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(10 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "furc":
                    text = "-"
                    if len(f_rows) >= 2:
                        f_info = f_rows[0] if is_upper else f_rows[-1]
                        raw_f = {clean_cell(df.iloc[f_info["label_row"], col+o]).upper(): clean_cell(df.iloc[f_info["value_row"], col+o]) for o in range(3)}
                        pref = ["M", "B", "D", "L"] if is_upper else ["B", "L", "M", "D"]
                        extracted = "".join([f"{l}{raw_f[l]}" for l in pref if l in raw_f and raw_f[l] in ["1", "2", "3"]])
                        if extracted: text = extracted
                    run = p_d.add_run(); run.text = text
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(10 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "mob":
                    v = clean_cell(df.iloc[r_idx, col]) if r_idx is not None else ""
                    val = {"1": "I", "2": "II", "3": "III"}.get(v, "WNL")
                    run = p_d.add_run(); run.text = val
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(10 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "prog":
                    run = p_d.add_run(); run.text = "G"
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(10 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, _rgb((0, 255, 0))

    # 5. 缺牙大合併 + 純白斜線
    for c_i, t in enumerate(teeth):
        if t in dynamic_missing:
            start_cell = table.cell(1, c_i + 1)
            end_cell = table.cell(total_rows - 1, c_i + 1)
            start_cell.merge(end_cell)
            
            merged = table.cell(1, c_i + 1)
            merged.vertical_anchor = MSO_ANCHOR.MIDDLE
            merged.text_frame.clear()
            _apply_cell_density(merged)
            merged.fill.background()
            _add_diagonal_strikethrough(merged)

    # 6. 純白邊框刷新
    for r in range(total_rows):
        for c in range(total_cols):
            _set_cell_border_to_white(table.cell(r, c))
