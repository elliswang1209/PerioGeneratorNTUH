# pptx_engine.py
"""
簡報生成引擎：支援缺牙整欄合併 + 對角斜線，以及數值去除浮點數格式化。
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
    collect_absolute_row_indices, 
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
    """🚀 為缺牙合併儲存格塗上物理 100% 純白斜對角線 (lnDiagonalDown)"""
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

# ==============================================================================
# 簡報直出對外接口
# ==============================================================================

def create_six_sextants_presentation(df, missing_teeth: Set[int]) -> BytesIO:
    """模式 A：生成 6 頁 Initial 六象限簡報"""
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial Charting", teeth, df, missing_teeth, is_comparison=False)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

def create_comparison_presentation(df, missing_teeth: Set[int]) -> BytesIO:
    """模式 B：生成 12 頁 (6 頁 Initial + 6 頁 Initial vs Re-eval 對比) 簡報"""
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    # 前 6 頁：Initial Stage
    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial Stage", teeth, df, missing_teeth, is_comparison=False)

    # 後 6 頁：Initial vs Re-evaluation
    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial vs Re-evaluation", teeth, df, missing_teeth, is_comparison=True)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

# ==============================================================================
# 核心繪圖與數據真實填入邏輯
# ==============================================================================

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

    # 2. 數據對照列地圖
    tooth_rows = find_tooth_rows(df)
    if len(tooth_rows) < 2:
        return
    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    anatomy_map_I = collect_absolute_row_indices(df, target_stage="I")
    anatomy_map_R = collect_absolute_row_indices(df, target_stage="R") if is_comparison else {}

    row_meta_single = [
        {"type": "pd", "s_row": "Line1", "header": "Probing Depth (B)" if teeth[0] < 30 else "Probing Depth (L)"},
        {"type": "pd", "s_row": "Line2", "header": "Probing Depth (P)" if teeth[0] < 30 else "Probing Depth (B)"},
        {"type": "gm", "s_row": "Line1", "header": "Recession"},
        {"type": "gm", "s_row": "Line2", "header": ""},
        {"type": "km", "header": "Masticatory Mucosa"},
        {"type": "furc", "header": "Furcation Involvement"},
        {"type": "mob", "header": "Mobility"},
        {"type": "prog", "header": "Prognosis"}
    ]

    total_rows = 1 + (len(row_meta_single) * 2 if is_comparison else len(row_meta_single))
    total_cols = 1 + len(teeth)

    if not is_comparison:
        col_width_left, col_width_data = Inches(2.2), Inches(1.8)
        row_height, top_pos = Inches(config.TABLE_ROW_HEIGHT), Inches(1.3)
    else:
        col_width_left, col_width_data = Inches(1.6), Inches(0.9)
        row_height, top_pos = Inches(0.35), Inches(1.2)

    total_table_width = col_width_left + col_width_data * len(teeth)
    total_table_height = row_height * total_rows
    left_pos = Inches(config.PPT_SLIDE_WIDTH) - total_table_width - Inches(0.5)

    table_shape = slide.shapes.add_table(total_rows, total_cols, left_pos, top_pos, total_table_width, total_table_height)
    table = table_shape.table
    _remove_default_table_style(table)

    table.columns[0].width = col_width_left
    for c in range(1, total_cols): table.columns[c].width = col_width_data
    for r in range(total_rows): table.rows[r].height = row_height

    text_white, text_alert = _rgb(config.COLOR_TEXT_WHITE), _rgb(config.COLOR_TEXT_ALERT)

    # 3. Header (Tooth)
    c_t0 = table.cell(0, 0); c_t0.vertical_anchor = MSO_ANCHOR.MIDDLE
    _apply_cell_density(c_t0); c_t0.fill.background()
    p_t0 = c_t0.text_frame.paragraphs[0]; p_t0.alignment = PP_ALIGN.CENTER
    r_t0 = p_t0.add_run(); r_t0.text = "Tooth"
    r_t0.font.name, r_t0.font.size, r_t0.font.bold, r_t0.font.color.rgb = config.FONT_PRIMARY, Pt(13 if is_comparison else 14), True, text_white

    for c_i, t in enumerate(teeth):
        c_t = table.cell(0, c_i + 1); c_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        _apply_cell_density(c_t); c_t.fill.background()
        p_t = c_t.text_frame.paragraphs[0]; p_t.alignment = PP_ALIGN.CENTER
        r_t = p_t.add_run(); r_t.text = str(t)
        r_t.font.name, r_t.font.size, r_t.font.bold, r_t.font.color.rgb = config.FONT_PRIMARY, Pt(13 if is_comparison else 14), True, text_white

    active_row_metas = []
    if not is_comparison:
        for m in row_meta_single:
            active_row_metas.append({**m, "stage": "I"})
    else:
        for m in row_meta_single:
            active_row_metas.append({**m, "stage": "I"})
            active_row_metas.append({**m, "stage": "R"})

    f_rows = find_furcation_rows(df)

    # 🚀 動態缺失牙補充集合（當讀取出來都是 ??? 時，自動標記為缺失牙）
    dynamic_missing = set(missing_teeth)

    # 4. 數據填寫
    for r_i, meta in enumerate(active_row_metas):
        r_ppt = r_i + 1
        c_lbl = table.cell(r_ppt, 0); c_lbl.vertical_anchor = MSO_ANCHOR.MIDDLE
        _apply_cell_density(c_lbl); c_lbl.fill.background()
        
        lbl_text = meta["header"]
        if is_comparison:
            tag = " (I)" if meta["stage"] == "I" else " (R)"
            lbl_text = f"{lbl_text}{tag}" if lbl_text else ""
            
        if lbl_text:
            p_lbl = c_lbl.text_frame.paragraphs[0]; p_lbl.alignment = PP_ALIGN.CENTER
            r_lbl = p_lbl.add_run(); r_lbl.text = lbl_text
            r_lbl.font.name, r_lbl.font.size, r_lbl.font.bold, r_lbl.font.color.rgb = config.FONT_PRIMARY, Pt(10 if is_comparison else 12), True, text_white

        for c_i, t in enumerate(teeth):
            cell_d = table.cell(r_ppt, c_i + 1); cell_d.vertical_anchor = MSO_ANCHOR.MIDDLE
            _apply_cell_density(cell_d); cell_d.fill.background()

            q = t // 10
            col = upper_cols.get(t) if q in [1, 2] else lower_cols.get(t)
            anatomy = anatomy_map_I if meta["stage"] == "I" else anatomy_map_R
            p_d = cell_d.text_frame.paragraphs[0]; p_d.alignment = PP_ALIGN.CENTER

            if col is not None:
                m_type = meta["type"]

                if m_type == "pd":
                    row_num = (anatomy.get("up_b_pd") if meta["s_row"] == "Line1" else anatomy.get("up_p_pd")) if q in [1,2] else (anatomy.get("lo_l_pd") if meta["s_row"] == "Line1" else anatomy.get("lo_b_pd"))
                    digits = get_three_digit_raw_list(df, row_num, col)
                    
                    # 🚀 如果讀出來是 ['?', '?', '?']，自動補入動態缺牙集合
                    if all(d.strip() in ['?', ''] for d in digits):
                        dynamic_missing.add(t)

                    for d in digits:
                        display_val = clean_cell(d)
                        run = p_d.add_run(); run.text = f" {display_val} " if len(display_val)>=2 else display_val
                        run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                        if display_val.isdigit() and int(display_val) >= 5:
                            run.font.bold = True
                            run.font.color.rgb = text_alert
                        else:
                            run.font.bold = False
                            run.font.color.rgb = text_white

                elif m_type == "gm":
                    row_num = (anatomy.get("up_b_gm") if meta["s_row"] == "Line1" else anatomy.get("up_p_gm")) if q in [1,2] else (anatomy.get("lo_l_gm") if meta["s_row"] == "Line1" else anatomy.get("lo_b_gm"))
                    digits = get_three_digit_raw_list(df, row_num, col)
                    for d in digits:
                        display_val = clean_cell(d)
                        run = p_d.add_run(); run.text = f" {display_val} " if len(display_val)>=2 else display_val
                        run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                        run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "km":
                    row_num = anatomy.get("up_km") if q in [1, 2] else anatomy.get("lo_km")
                    val = clean_cell(df.iloc[row_num, col]) if row_num is not None else "0"
                    run = p_d.add_run(); run.text = val if val!="" else "0"
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "furc":
                    text = "-"
                    if len(f_rows) >= 2:
                        f_info = f_rows[0] if q in [1, 2] else f_rows[-1]
                        raw_f = {clean_cell(df.iloc[f_info["label_row"], col+o]).upper(): clean_cell(df.iloc[f_info["value_row"], col+o]) for o in range(3)}
                        pref = ["M", "B", "D", "L"] if q in [1, 2] else ["B", "L", "M", "D"]
                        extracted = "".join([f"{l}{raw_f[l]}" for l in pref if l in raw_f and raw_f[l] in ["1", "2", "3"]])
                        if extracted: text = extracted
                    run = p_d.add_run(); run.text = text
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "mob":
                    row_num = anatomy.get("up_mob") if q in [1, 2] else anatomy.get("lo_mob")
                    v = clean_cell(df.iloc[row_num, col]) if row_num is not None else ""
                    val = {"1": "I", "2": "II", "3": "III"}.get(v, "WNL")
                    run = p_d.add_run(); run.text = val
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, text_white

                elif m_type == "prog":
                    run = p_d.add_run(); run.text = "G"
                    run.font.name, run.font.size = config.FONT_PRIMARY, Pt(11 if is_comparison else 14)
                    run.font.bold, run.font.color.rgb = False, _rgb((0, 255, 0))

    # 5. 🚀 處理缺牙：整欄大合併 + 繪製跨欄對角斜線 (Diagonal Down)
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
            
            # 🚀 畫上對角斜線
            _add_diagonal_strikethrough(merged)

    # 6. 刷新 100% 純白實線外邊框
    for r in range(total_rows):
        for c in range(total_cols):
            _set_cell_border_to_white(table.cell(r, c))
