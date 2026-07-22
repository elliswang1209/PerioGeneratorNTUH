# config.py
"""
全域設定檔：集中管理 Streamlit UI 與 PPTX 簡報的所有排版、顏色、字體與象限分組常數。
"""

# ==============================================================================
# PPTX 簡報佈局與尺寸設定
# ==============================================================================
PPT_SLIDE_WIDTH = 13.333  # 16:9 寬螢幕標準寬度 (英吋)
PPT_SLIDE_HEIGHT = 7.5    # 16:9 寬螢幕標準高度 (英吋)

# 簡報字體設定
FONT_PRIMARY = 'Times New Roman'
FONT_SIZE_TITLE = 36
FONT_SIZE_LABEL = 13
FONT_SIZE_DATA = 12

# 簡報色彩計畫 (RGB 格式)
COLOR_BG_DARK = (0, 0, 0)            # 暗黑主題背景
COLOR_TEXT_WHITE = (255, 255, 255)   # 標頭與一般數據白字
COLOR_TEXT_ALERT = (255, 50, 50)     # PD >= 5mm 警示紅
COLOR_TITLE_WHEAT = (245, 222, 179)  # 簡報標題麥黃色

# 雙期對比配色 (Initial vs Re-evaluation)
COLOR_INITIAL_TAG = (255, 180, 0)    # Initial (I) 橘黃色標示
COLOR_REEVAL_TAG = (100, 220, 100)   # Re-evaluation (R) 亮綠色標示

# 簡報表格內邊距與列高
TABLE_ROW_HEIGHT = 0.42
CELL_INTERNAL_MARGIN = 0.015

# ==============================================================================
# 六象限牙位分組 (FDI 記法)
# ==============================================================================
SEXTANTS = {
    "Upper Right Sextant": [18, 17, 16, 15, 14],
    "Upper Anterior Sextant": [13, 12, 11, 21, 22, 23],
    "Upper Left Sextant": [24, 25, 26, 27, 28],
    "Lower Left Sextant": [34, 35, 36, 37, 38],
    "Lower Anterior Sextant": [43, 42, 41, 31, 32, 33],
    "Lower Right Sextant": [48, 47, 46, 45, 44],
}

# 臨床警示閾值
ALERT_PD_THRESHOLD = 5
