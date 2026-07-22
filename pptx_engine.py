# pptx_engine.py
"""
簡報生成引擎：負責將結構化的牙周資料繪製成 PPTX 檔案
特點：表格100%靠右對齊、純黑背景、100%純白邊框、單期與雙期對比尺寸動態最佳化。
"""
from io import BytesIO
from typing import Dict, Set, Any, List
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

import config

def _rgb(color_tuple):
    return RGBColor(*color_tuple)

def _apply_cell_density(cell):
    """消除儲存格內部的 Margin 留白"""
    cell.margin_top = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_bottom = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_left = Inches(config.CELL_INTERNAL_MARGIN)
    cell.margin_right = Inches(config.CELL_INTERNAL_MARGIN)

def _set_cell_border_to_white(cell):
    """強行將儲存格四周刷上 100% 飽和度的純白實線邊框"""
    tcPr = cell._tc.get_or_add_tcPr()
    border_xml = (
        f'<a:ln {nsdecls("a")} w="12700" cmpd="s" algn="ctr">'
        f'  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'  # 100% 純白
        f'  <a:prstDash val="solid"/>'                           # 實線
        f'</a:ln>'
    )
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{side}')
        if existing is not None:
            tcPr.remove(existing)
        tcPr.append(parse_xml(f'<a:{side} {nsdecls("a")}>{border_xml}</a:{side}>'))

def _remove_default_table_style(table):
    """移除 PowerPoint 預設表格樣式 (如藍底白字)，植入透明無樣式空 ID"""
    try:
        tblPr = table._tbl.tblPr
        tableStyleId = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
        if tableStyleId is not None:
            tblPr.remove(tableStyleId)
        tblPr.append(parse_xml(f'<a:tableStyleId {nsdecls("a")}>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>'))
    except Exception:
        pass

def _set_cell_text(cell, text: str, font_size: int, font_color: RGBColor, bold: bool = False, italic: bool = False):
    """統一樣式設定輔助函式"""
    cell.text = str(text)
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.name = config.FONT_PRIMARY
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = font_color

# ==============================================================================
# 簡報直出對外接口
# ==============================================================================

def create_six_sextants_presentation(records: Dict[int, Dict[str, Any]], missing_teeth: Set[int]) -> BytesIO:
    """模式 A：生成 6 頁 Initial 六象限簡報"""
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial Charting", teeth, records, missing_teeth, is_comparison=False)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

def create_comparison_presentation(records: Dict[int, Dict[str, Any]], missing_teeth: Set[int]) -> BytesIO:
    """模式 B：生成 12 頁 (6 頁 Initial + 6 頁 Initial vs Re-eval 對比) 簡報"""
    prs = Presentation()
    prs.slide_width = Inches(config.PPT_SLIDE_WIDTH)
    prs.slide_height = Inches(config.PPT_SLIDE_HEIGHT)
    blank_layout = prs.slide_layouts[6]

    # 前 6 頁：Initial Stage
    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial Stage", teeth, records, missing_teeth, is_comparison=False)

    # 後 6 頁：Initial vs Re-evaluation
    for sextant_name, teeth in config.SEXTANTS.items():
        slide = prs.slides.add_slide(blank_layout)
        _draw_sextant_slide(slide, f"{sextant_name} - Initial vs Re-evaluation", teeth, records, missing_teeth, is_comparison=True)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

# ==============================================================================
# 核心繪圖邏輯 (靠右對齊 & 尺寸動態適應)
# ==============================================================================

def _draw_sextant_slide(slide, title_text: str, teeth: List[int], records: Dict[int, Dict[str, Any]], missing_teeth: Set[int], is_comparison: bool):
    # 1. 投影片背景設置純黑
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(config.COLOR_BG_DARK)

    # 2. 標題：麥黃色、斜體、靠左對齊
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

    # 3. 定義列屬性
    row_labels = [
        "Probing Depth (B/P)" if teeth[0] < 30 else "Probing Depth (L/B)",
        "Recession",
        "Masticatory Mucosa",
        "Furcation Involvement",
        "Mobility",
        "Prognosis"
    ]
    
    rows_per_metric = 2 if is_comparison else 1
    total_rows = 1 + len(row_labels) * rows_per_metric
    total_cols = 1 + len(teeth)

    # 4. 根據模式 (Initial vs Initial+Re) 計算靠右對齊座標與尺寸
    if not is_comparison:
        # --- 單期模式 (維持原本寬度與高度) ---
        col_width_left = Inches(2.2)
        col_width_data = Inches(1.8)
        row_height = Inches(config.TABLE_ROW_HEIGHT)
        top_pos = Inches(1.3)
    else:
        # --- 雙期對比模式 (高度與投影片切齊、寬度比一半少一點) ---
        col_width_left = Inches(1.6)
        col_width_data = Inches(0.9)  # 5顆牙總寬約 1.6 + 0.9*5 = 6.1 英吋 (< 6.666)
        row_height = Inches(0.45)     # 13 列 * 0.45 = 5.85 英吋高度
        top_pos = Inches(1.2)         # 高度滿版切齊 (1.2 + 5.85 = 7.05 / 7.5 英吋)

    total_table_width = col_width_left + col_width_data * len(teeth)
    total_table_height = row_height * total_rows

    # 🚀 精確計算靠右對齊的 Left 座標 (右邊固定留白 0.5 英吋)
    right_margin = Inches(0.5)
    left_pos = Inches(config.PPT_SLIDE_WIDTH) - total_table_width - right_margin

    # 5. 新增表格並清除預設樣式
    table_shape = slide.shapes.add_table(total_rows, total_cols, left_pos, top_pos, total_table_width, total_table_height)
    table = table_shape.table
    _remove_default_table_style(table)

    table.columns[0].width = col_width_left
    for c in range(1, total_cols):
        table.columns[c].width = col_width_data
    for r in range(total_rows):
        table.rows[r].height = row_height

    text_white = _rgb(config.COLOR_TEXT_WHITE)

    # 6. 填入 Header (第一列牙位)
    cell_tooth = table.cell(0, 0)
    cell_tooth.vertical_anchor = MSO_ANCHOR.MIDDLE
    _apply_cell_density(cell_tooth)
    cell_tooth.fill.background()
    _set_cell_text(cell_tooth, "Tooth", 13 if is_comparison else 14, text_white, bold=True)

    for c_i, tooth in enumerate(teeth):
        cell_t = table.cell(0, c_i + 1)
        cell_t.vertical_anchor = MSO_ANCHOR.MIDDLE
        _apply_cell_density(cell_t)
        cell_t.fill.background()
        _set_cell_text(cell_t, str(tooth), 13 if is_comparison else 14, text_white, bold=True)

    # 7. 填入左側標籤與背景設定
    for r in range(total_rows):
        for c in range(total_cols):
            cell = table.cell(r, c)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            _apply_cell_density(cell)
            cell.fill.background()  # 透出簡報純黑背景
            
            if r > 0 and c == 0:
                metric_idx = (r - 1) // rows_per_metric
                stage_tag = " (I)" if (is_comparison and (r - 1) % 2 == 0) else (" (R)" if is_comparison else "")
                label_text = f"{row_labels[metric_idx]}{stage_tag}"
                _set_cell_text(cell, label_text, 10 if is_comparison else 12, text_white, bold=True)

    # 8. 處理缺失牙 (Missing Teeth) 大合併
    for c_i, tooth in enumerate(teeth):
        if tooth in missing_teeth:
            start_cell = table.cell(1, c_i + 1)
            end_cell = table.cell(total_rows - 1, c_i + 1)
            start_cell.merge(end_cell)
            
            merged_cell = table.cell(1, c_i + 1)
            merged_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            merged_cell.text_frame.clear()
            _apply_cell_density(merged_cell)
            merged_cell.fill.background()

    # 9. 最終重新覆蓋一圈 100% 純白實線邊框
    for r in range(total_rows):
        for c in range(total_cols):
            _set_cell_border_to_white(table.cell(r, c))
