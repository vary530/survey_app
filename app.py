import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
import io
import os
import re
import pdfplumber
import streamlit.components.v1 as components
from PIL import Image, ImageOps # 新增影像處理套件

# --- 1. 頁面設定 ---
st.set_page_config(
    page_title="永義物調整合", 
    page_icon="🏠", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- 核心邏輯 ---
TEMPLATE_FILE = "template.xlsx"
MAIN_ORDER = [
    "物件類型", "案名", "地址", "社區名稱", 
    "地上層", "地下層", "位於樓層", "格局", 
    "售價", "登記總建坪", "主建物坪數", "附屬建坪數", "公設坪數", "不含車位坪數", 
    "車位坪數", "車位形式", "車位樓層", "汽車編號", "機車位樓層", "機車編號", 
    "使用現況", "總戶數", "同層戶數", "電梯數", "有無警衛", "管理費", "繳納方式", 
    "建築完成日", "瓦斯", "學校", "市場", "公園", "公設比", 
    "建物KEY", "座向", "土地面積", "權利範圍", 
    "冒泡位置圖", "承辦人電話", "委託契約書編號" 
]
OTHER_ORDER = [
    "房地合一", "面道路", "貸款設定", "車位價格", "房屋單價"
]

# --- 2. 視覺設計 (修正版：核彈級隱藏介面雜訊) ---
def inject_custom_styles():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');
            
            /* 全域重置 */
            * {
                box-sizing: border-box;
            }

            /* --- 強力隱藏 Streamlit 預設介面 --- */
            
            /* 1. 徹底隱藏上方 Header (包含漢堡選單、Deploy 按鈕、裝飾條) */
            header, [data-testid="stHeader"], .stAppHeader {
                display: none !important;
                visibility: hidden !important;
                height: 0px !important;
                opacity: 0 !important;
                pointer-events: none !important;
            }

            /* 2. 徹底隱藏下方 Footer (Hosted with Streamlit) */
            footer, [data-testid="stFooter"] {
                display: none !important;
                visibility: hidden !important;
                height: 0px !important;
            }

            /* 3. 隱藏選單按鈕與開發者工具 */
            #MainMenu {
                display: none !important;
                visibility: hidden !important;
            }
            [data-testid="stToolbar"] {
                display: none !important;
                visibility: hidden !important;
            }
            [data-testid="stDecoration"] {
                display: none !important;
            }
            [data-testid="stStatusWidget"] {
                display: none !important;
            }
            .stDeployButton {
                display: none !important;
            }
            
            /* 隱藏右下角可能出現的 Manage app 按鈕區域 */
            div[class*="viewerBadge"] {
                display: none !important;
            }

            /* 全域背景設定 */
            .stApp {
                background-color: #050505;
                background-image: radial-gradient(circle at 50% 0%, #1a1a1a 0%, #050505 80%);
                background-attachment: fixed; 
                background-size: cover;
                font-family: 'Inter', 'Noto Sans TC', sans-serif;
                color: #d1d5db;
                /* 因為 header 已經 display:none，不需要負 margin */
                margin-top: 0px !important;
            }

            /* 修正主要內容區域的 padding，去除上方留白 */
            .block-container { 
                padding-top: 2rem !important; /* 確保內容不會貼齊頂部太緊，保留適當呼吸空間 */
                padding-left: 1rem;
                padding-right: 1rem;
                padding-bottom: 5rem;
                max-width: 600px; 
                margin: 0 auto;
            }

            /* --- 表單區塊樣式 --- */
            [data-testid="stForm"] {
                background: rgba(30, 30, 30, 0.4);
                border: 1px solid rgba(197, 160, 101, 0.2);
                border-radius: 16px; 
                padding: 20px 24px;
                backdrop-filter: blur(12px);
                margin-top: 10px; /* 稍微縮小頂部間距 */
                width: 100%;
                box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
            }

            /* 上傳區塊美化 */
            [data-testid="stFileUploader"] {
                background: rgba(30, 30, 30, 0.3);
                border: 1px dashed rgba(197, 160, 101, 0.4);
                border-radius: 10px;
                padding: 12px;
                transition: border-color 0.3s;
            }
            [data-testid="stFileUploader"]:hover {
                border-color: #c5a065;
            }

            /* 輸入框美化 */
            .stTextInput > div > div > input, 
            .stSelectbox > div > div > div, 
            .stTextArea > div > div > textarea,
            .stNumberInput > div > div > input {
                background-color: #121212 !important; 
                color: #e5e5e5 !important;
                border: 1px solid #333 !important;
                border-radius: 6px;
                font-size: 16px;
                padding: 8px 12px;
            }
            
            ::placeholder {
                color: #555 !important;
                opacity: 1;
            }
            
            /* 呼吸燈特效 */
            @keyframes breathe {
                0% { border-color: rgba(197, 160, 101, 0.3); box-shadow: 0 0 2px rgba(197, 160, 101, 0.1); }
                50% { border-color: rgba(197, 160, 101, 0.9); box-shadow: 0 0 8px rgba(197, 160, 101, 0.2); }
                100% { border-color: rgba(197, 160, 101, 0.3); box-shadow: 0 0 2px rgba(197, 160, 101, 0.1); }
            }

            .stTextInput > div > div > input:focus, 
            .stSelectbox > div > div > div:focus,
            .stTextArea > div > div > textarea:focus,
            .stNumberInput > div > div > input:focus {
                background-color: #080808 !important;
                outline: none;
                animation: breathe 3s infinite ease-in-out;
            }

            /* 按鈕美化 */
            .stButton > button {
                width: 100%;
                background: linear-gradient(to bottom, #c5a065, #8e733b);
                color: #000;
                font-weight: 700;
                border: none;
                border-radius: 6px;
                padding: 12px 0;
                letter-spacing: 1px;
                margin-top: 20px;
                font-size: 16px;
                box-shadow: 0 4px 15px rgba(197, 160, 101, 0.2);
            }
            .stButton > button:hover { 
                filter: brightness(1.15); 
                transform: translateY(-1px);
            }
            .stButton > button:active {
                transform: translateY(1px);
            }

            /* 標題與文字 */
            h1 {
                text-align: center !important; 
                color: #e5e5e5; 
                font-weight: 400; 
                letter-spacing: 4px; 
                margin-bottom: 8px;
                font-size: 1.8rem !important; 
                width: 100%;
                display: flex;
                justify-content: center;
                align-items: center;
                text-shadow: 0 2px 4px rgba(0,0,0,0.5);
                padding-top: 0px;
            }
            .subtitle {
                text-align: center !important;
                color: #c5a065;
                font-size: 0.75rem; 
                letter-spacing: 3px;
                font-weight: 500;
                margin-bottom: 35px;
                font-family: 'Inter', sans-serif;
                width: 100%;
                display: flex;
                justify-content: center;
                opacity: 0.9;
            }

            /* 儀表板 */
            .dashboard-grid {
                display: grid;
                grid-template-columns: repeat(3, 1fr); 
                gap: 10px;
                margin-top: 15px;
                margin-bottom: 15px;
            }
            .dash-item {
                background: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 8px;
                padding: 10px;
                text-align: center;
                min-width: 0;
            }
            .dash-label {
                font-size: 11px;
                color: #888;
                margin-bottom: 4px;
            }
            .dash-value {
                font-size: 14px;
                color: #c5a065;
                font-weight: 600;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            
            ::-webkit-scrollbar { display: none; }
            
        </style>
    """, unsafe_allow_html=True)

# --- 輔助函式 ---
def full_to_half(s):
    if not s: return ""
    return s.translate(str.maketrans('０１２３４５６７８９', '0123456789'))

def chinese_to_arabic(cn_str):
    if not cn_str: return ""
    cn_map = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10, '0':0, '1':1, '2':2, '3':3, '4':4, '5':5, '6':6, '7':7, '8':8, '9':9}
    clean_str = cn_str.replace('層', '').replace('樓', '').strip()
    if clean_str.isdigit(): return str(int(clean_str))
    try:
        val = 0
        if len(clean_str) == 1: val = cn_map.get(clean_str, 0)
        elif len(clean_str) == 2:
            if clean_str[0] == '十': val = 10 + cn_map.get(clean_str[1], 0)
            elif clean_str[1] == '十': val = cn_map.get(clean_str[0], 0) * 10
        elif len(clean_str) == 3:
             val = cn_map.get(clean_str[0], 0) * 10 + cn_map.get(clean_str[2], 0)
        return str(val) if val > 0 else cn_str
    except: return cn_str

def format_date_roc(date_str):
    if not date_str: return ""
    match = re.match(r'(\d+)[/.-](\d+)[/.-](\d+)', date_str)
    if match:
        y, m, d = match.groups()
        return f"民國{y}年{m}月{d}日"
    return date_str

def format_layout(layout_str):
    if not layout_str: return ""
    parts = re.split(r'[/, .]', layout_str)
    parts = [p for p in parts if p.strip()]
    result = ""
    if len(parts) >= 1: result += f"{parts[0]}房"
    if len(parts) >= 2: result += f"{parts[1]}廳"
    if len(parts) >= 3: result += f"{parts[2]}衛浴"
    if len(parts) >= 4: result += f"{parts[3]}陽台"
    return result if result else layout_str

def safe_float_convert(value):
    """安全轉換字串為浮點數，失敗回傳 0.0"""
    try:
        if not value: return 0.0
        clean_val = re.sub(r'[^\d.]', '', str(value))
        return float(clean_val)
    except:
        return 0.0

def crop_image_to_ratio(image, target_ratio_w=27, target_ratio_h=16):
    """將圖片置中剪裁為指定長寬比"""
    original_w, original_h = image.size
    target_aspect = target_ratio_w / target_ratio_h
    current_aspect = original_w / original_h

    if current_aspect > target_aspect:
        new_w = int(original_h * target_aspect)
        offset = (original_w - new_w) // 2
        box = (offset, 0, offset + new_w, original_h)
    else:
        new_h = int(original_w / target_aspect)
        offset = (original_h - new_h) // 2
        box = (0, offset, original_w, offset + new_h)
    
    return image.crop(box)

def calculate_cell_pixels(ws, coord):
    """計算 Excel 儲存格 (含合併) 的像素大小"""
    target_range = None
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            target_range = merged_range
            break
    
    if target_range:
        min_col, min_row, max_col, max_row = target_range.min_col, target_range.min_row, target_range.max_col, target_range.max_row
    else:
        c = ws[coord]
        min_col, min_row, max_col, max_row = c.column, c.row, c.column, c.row

    total_width = 0
    for col_idx in range(min_col, max_col + 1):
        col_letter = get_column_letter(col_idx)
        cw = ws.column_dimensions[col_letter].width
        if cw is None: cw = 9 
        total_width += cw * 7.7 
        
    total_height = 0
    for row_idx in range(min_row, max_row + 1):
        rh = ws.row_dimensions[row_idx].height
        if rh is None: rh = 15
        total_height += rh * 1.34 
        
    return total_width, total_height

def parse_transcript_pdf(pdf_file):
    data = {}
    full_text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                full_text += page.extract_text() + "\n"
        
        lines = full_text.split('\n')
        address_prefix = ""
        address_road = ""
        for i, line in enumerate(lines):
            line = line.strip()
            if "建物標示部" in line:
                for offset in range(1, 5):
                    if i + offset < len(lines):
                        txt = lines[i+offset]
                        match = re.search(r'(.+?[市縣].+?[區鄉鎮市])', txt)
                        if match:
                            address_prefix = match.group(1)
                            break
            if "建物門牌" in line:
                parts = line.split("建物門牌")
                if len(parts) > 1 and parts[1].strip():
                    address_road = parts[1].strip()
                elif i+1 < len(lines):
                    address_road = lines[i+1].strip()

        if address_prefix or address_road:
            full_addr = f"{address_prefix}{address_road}"
            data["地址"] = full_to_half(full_addr).replace(" ", "")

        date_match = re.search(r'建築完成日期\s*([民國\d]+年\d+月\d+日)', full_text)
        if date_match: data["建築完成日"] = date_match.group(1)

        layer_m2_matches = re.findall(r'層次面積\s*([\d\.]+)\s*平方公尺', full_text)
        if layer_m2_matches:
            total_main_m2 = sum(float(x) for x in layer_m2_matches)
            data["主建物坪數"] = str(round(total_main_m2 * 0.3025, 3))

        try:
            start = full_text.find("附屬建物用途")
            end = full_text.find("共有部分")
            if start != -1:
                sub_text = full_text[start:end] if end != -1 else full_text[start:]
                annex_matches = re.findall(r'面積\s*([\d\.]+)\s*平方公尺', sub_text)
                if annex_matches:
                    total_annex_m2 = sum(float(x) for x in annex_matches)
                    data["附屬建坪數"] = str(round(total_annex_m2 * 0.3025, 3))
        except: pass

        floors_match = re.search(r'層數\s*(\d+)層', full_text)
        if floors_match: data["地上層"] = str(int(floors_match.group(1)))

        layer_match = re.search(r'層次\s*([^\d\s]+)層', full_text)
        if layer_match and "面積" not in layer_match.group(0):
            data["位於樓層"] = chinese_to_arabic(layer_match.group(1))
        
    except Exception as e:
        st.error(f"PDF 解析錯誤: {e}")
    return data

def main():
    inject_custom_styles()

    if not os.path.exists(TEMPLATE_FILE):
        st.error(f"系統錯誤：找不到 {TEMPLATE_FILE}")
        return
    try:
        wb = load_workbook(TEMPLATE_FILE)
        target_sheet = None
        for sheetname in wb.sheetnames:
            if "物調表" in sheetname:
                target_sheet = wb[sheetname]
                break
        if target_sheet is None: target_sheet = wb.active 
    except Exception as e:
        st.error(f"系統錯誤：讀取模板失敗 {e}")
        return

    label_to_coord = {}
    scanned_items = []
    for row in target_sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and '"""' in cell.value:
                raw_txt = cell.value
                label_name = ""
                content_part = ""
                
                match_star = re.search(r'\*(.*?)\*(.*)', raw_txt)
                if match_star:
                    label_name = match_star.group(1).strip()
                    content_part = match_star.group(2).replace('"""', '')
                else:
                    label_name = raw_txt.replace('"""', '').strip()
                    content_part = label_name

                options = []
                input_type = "text"
                
                if "□" in content_part:
                    input_type = "select"
                    segments = content_part.split('□')
                    options = [s.strip() for s in segments if s.strip()]
                    options.insert(0, "請選擇...")
                
                if "特色" in label_name or "說明" in label_name:
                    input_type = "textarea"
                elif "冒泡" in label_name:
                    input_type = "image_upload"

                item_data = {
                    "label": label_name,
                    "coordinate": cell.coordinate,
                    "type": input_type,
                    "options": options
                }
                label_to_coord[label_name] = cell.coordinate
                scanned_items.append(item_data)

    st.markdown("<h1>永義物調整合</h1>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>YUNGYI PROPERTY INTEGRATION</div>", unsafe_allow_html=True)

    st.markdown("<div style='color:#c5a065; font-size:15px; font-weight:bold; margin-bottom:10px; margin-top:20px;'>智慧匯入中心</div>", unsafe_allow_html=True)
    
    uploaded_pdf = st.file_uploader("點此上傳建物謄本 (PDF)", type=['pdf'])
    
    if uploaded_pdf:
        if 'last_uploaded_pdf' not in st.session_state or st.session_state.last_uploaded_pdf != uploaded_pdf.name:
            with st.spinner("分析中..."):
                parsed = parse_transcript_pdf(uploaded_pdf)
                st.session_state.pdf_parsed_data = parsed
                st.session_state.last_uploaded_pdf = uploaded_pdf.name
        
        if 'pdf_parsed_data' in st.session_state:
            data = st.session_state.pdf_parsed_data
            
            grid_html = f"""
            <div class="dashboard-grid">
                <div class="dash-item"><div class="dash-label">地址</div><div class="dash-value">{data.get('地址', '-')}</div></div>
                <div class="dash-item"><div class="dash-label">建築完成日</div><div class="dash-value">{data.get('建築完成日', '-')}</div></div>
                <div class="dash-item"><div class="dash-label">主建物坪數</div><div class="dash-value">{data.get('主建物坪數', '-')}</div></div>
                <div class="dash-item"><div class="dash-label">附屬建坪數</div><div class="dash-value">{data.get('附屬建坪數', '-')}</div></div>
                <div class="dash-item"><div class="dash-label">地上層</div><div class="dash-value">{data.get('地上層', '-')}</div></div>
                <div class="dash-item"><div class="dash-label">位於樓層</div><div class="dash-value">{data.get('位於樓層', '-')}</div></div>
            </div>
            """
            st.markdown(grid_html, unsafe_allow_html=True)
            
            if st.button("匯入建物基本資料", type="primary"):
                count = 0
                for pdf_key, pdf_val in st.session_state.pdf_parsed_data.items():
                    target_coord = None
                    if pdf_key in label_to_coord: target_coord = label_to_coord[pdf_key]
                    if not target_coord:
                        for lbl, coord in label_to_coord.items():
                            if pdf_key in lbl or lbl in pdf_key:
                                target_coord = coord
                                break
                    if target_coord:
                        st.session_state[target_coord] = pdf_val
                        count += 1
                if count > 0:
                    st.success("資料已匯入")

    user_inputs = {} 
    uploaded_map_image = None
    scanned_dict = {item["label"]: item for item in scanned_items}

    with st.form("survey_form"):
        st.markdown("<div style='color:#c5a065; font-size:15px; font-weight:bold; margin-bottom:15px;'>不動產基本資料</div>", unsafe_allow_html=True)

        for label in MAIN_ORDER:
            found_key = label if label in scanned_dict else None
            if not found_key:
                for k in scanned_dict.keys():
                    if label in k or k in label:
                        found_key = k
                        break
            
            if found_key:
                item = scanned_dict[found_key]
                coord = item["coordinate"]
                
                if label == "地址":
                    val = st.text_input(label, key=coord)
                    user_inputs[coord] = val
                    if val:
                        map_url = f"https://www.google.com/maps/search/?api=1&query={val}"
                        st.markdown(f"<div style='text-align:right; margin-top:-5px; margin-bottom:10px;'><a href='{map_url}' target='_blank' style='font-size:12px; color:#888; text-decoration:none;'>📍 開啟地圖</a></div>", unsafe_allow_html=True)
                
                elif item["type"] == "select":
                    val = st.selectbox(found_key, item["options"], key=coord)
                    user_inputs[coord] = val if val != "請選擇..." else ""
                
                elif item["type"] == "textarea":
                    val = st.text_area(found_key, key=coord, height=120)
                    user_inputs[coord] = val

                elif item["type"] == "image_upload":
                    st.markdown(f"<div style='margin-top:15px; margin-bottom:5px; font-size:14px; color:#c5a065;'>{found_key}</div>", unsafe_allow_html=True)
                    uploaded_map_image = st.file_uploader("", type=['jpg', 'png', 'jpeg'], key=coord, label_visibility="collapsed")
                    st.markdown("<div style='font-size:12px; color:#666; margin-top:-5px;'>* 圖片將自動「置中剪裁 (27:16)」並拉伸填滿 Excel 儲存格</div>", unsafe_allow_html=True)
                    user_inputs[coord] = ""
                else:
                    placeholder_txt = ""
                    if "房屋單價" in found_key or "公設比" in found_key:
                        placeholder_txt = "輸入數字0系統匯出自動計算"
                    elif "登記總建坪" in found_key or "不含車位坪數" in found_key:
                        placeholder_txt = "輸入數字0系統匯出自動計算"
                    
                    val = st.text_input(found_key, key=coord, placeholder=placeholder_txt)
                    user_inputs[coord] = val
                
                if found_key in scanned_dict: del scanned_dict[found_key]

        if any(k in scanned_dict for k in OTHER_ORDER):
            st.markdown("<hr style='border-color: rgba(255,255,255,0.05); margin: 30px 0;'>", unsafe_allow_html=True)
            for label in OTHER_ORDER:
                if label in scanned_dict:
                    item = scanned_dict[label]
                    coord = item["coordinate"]
                    
                    if item["type"] == "select":
                        val = st.selectbox(label, item["options"], key=coord)
                        user_inputs[coord] = val if val != "請選擇..." else ""
                    elif item["type"] == "textarea":
                        val = st.text_area(label, key=coord, height=100)
                        user_inputs[coord] = val
                    else:
                        placeholder_txt = ""
                        if "房屋單價" in label or "公設比" in label:
                            placeholder_txt = "輸入數字0系統匯出自動計算"
                        elif "登記總建坪" in label or "不含車位坪數" in label:
                            placeholder_txt = "輸入數字0系統匯出自動計算"
                        
                        val = st.text_input(label, key=coord, placeholder=placeholder_txt)
                        user_inputs[coord] = val
                    
                    del scanned_dict[label]

        if scanned_dict:
            st.markdown("<hr style='border-color: rgba(255,255,255,0.05); margin: 30px 0;'>", unsafe_allow_html=True)
            for label, item in scanned_dict.items():
                coord = item["coordinate"]
                if item["type"] == "select":
                    val = st.selectbox(label, item["options"], key=coord)
                    user_inputs[coord] = val if val != "請選擇..." else ""
                elif item["type"] == "textarea":
                    val = st.text_area(label, key=coord, height=100)
                    user_inputs[coord] = val
                else:
                    placeholder_txt = ""
                    if "房屋單價" in label or "公設比" in label:
                        placeholder_txt = "輸入數字0系統匯出自動計算"
                    elif "登記總建坪" in label or "不含車位坪數" in label:
                        placeholder_txt = "輸入數字0系統匯出自動計算"

                    val = st.text_input(label, key=coord, placeholder=placeholder_txt)
                    user_inputs[coord] = val

        st.markdown("<br>", unsafe_allow_html=True)
        submitted = st.form_submit_button("匯出至Excel")

    if submitted:
        wb_output = load_workbook(TEMPLATE_FILE)
        ws_output = wb_output[target_sheet.title]

        coord_to_header = {item["coordinate"]: item["label"] for item in scanned_items}
        
        image_coords = [item["coordinate"] for item in scanned_items if item["type"] == "image_upload"]

        coord_price = next((k for k, v in coord_to_header.items() if "售價" in v), None)
        coord_total_area = next((k for k, v in coord_to_header.items() if "登記總建坪" in v), None)
        coord_area_no_parking = next((k for k, v in coord_to_header.items() if "不含車位" in v), None)
        coord_parking_area = next((k for k, v in coord_to_header.items() if "車位坪數" in v), None)
        coord_public_area = next((k for k, v in coord_to_header.items() if "公設坪數" in v), None)
        coord_unit_price = next((k for k, v in coord_to_header.items() if "房屋單價" in v), None)
        coord_public_ratio = next((k for k, v in coord_to_header.items() if "公設比" in v), None)
        coord_main_area = next((k for k, v in coord_to_header.items() if "主建物" in v), None)
        coord_annex_area = next((k for k, v in coord_to_header.items() if "附屬" in v), None)

        # 1. 計算不含車位坪數 (主+附+公)
        if coord_area_no_parking and user_inputs.get(coord_area_no_parking) == "0":
            try:
                a_main = safe_float_convert(user_inputs.get(coord_main_area))
                a_annex = safe_float_convert(user_inputs.get(coord_annex_area))
                a_pub = safe_float_convert(user_inputs.get(coord_public_area))
                user_inputs[coord_area_no_parking] = str(round(a_main + a_annex + a_pub, 3))
            except: pass

        # 2. 計算登記總建坪 (主+附+公+車)
        if coord_total_area and user_inputs.get(coord_total_area) == "0":
            try:
                a_main = safe_float_convert(user_inputs.get(coord_main_area))
                a_annex = safe_float_convert(user_inputs.get(coord_annex_area))
                a_pub = safe_float_convert(user_inputs.get(coord_public_area))
                a_park = safe_float_convert(user_inputs.get(coord_parking_area))
                user_inputs[coord_total_area] = str(round(a_main + a_annex + a_pub + a_park, 3))
            except: pass

        # 3. 計算房屋單價
        if coord_unit_price and user_inputs.get(coord_unit_price) == "0":
            try:
                p = safe_float_convert(user_inputs.get(coord_price))
                a = safe_float_convert(user_inputs.get(coord_area_no_parking))
                if a > 0:
                    res = round(p / a, 2)
                    user_inputs[coord_unit_price] = str(res)
            except: pass

        # 4. 計算公設比
        if coord_public_ratio and user_inputs.get(coord_public_ratio) == "0":
            try:
                pub = safe_float_convert(user_inputs.get(coord_public_area))
                a = safe_float_convert(user_inputs.get(coord_area_no_parking))
                if a > 0:
                    res = round((pub / a) * 100, 1)
                    user_inputs[coord_public_ratio] = f"{res}%"
            except: pass

        for coord, value in user_inputs.items():
            if coord in image_coords:
                continue

            cell = ws_output[coord]
            final_val = value if value else ""
            header = coord_to_header.get(coord, "")
            
            if "完成日" in header or "日期" in header:
                final_val = format_date_roc(final_val)
            elif "格局" in header:
                final_val = format_layout(final_val)
            
            # 自動加萬
            keywords_for_wan = ["售價", "單價", "價格", "貸款"]
            if any(k in header for k in keywords_for_wan) and final_val:
                v_str = str(final_val).strip()
                if v_str.replace('.', '', 1).isdigit() and "萬" not in v_str:
                    final_val = f"{v_str}萬"
            
            # 管理費自動加元
            if "管理費" in header and final_val:
                v_str = str(final_val).strip()
                if v_str and "元" not in v_str:
                    final_val = f"{v_str}元"

            cell.value = final_val

        if uploaded_map_image:
            try:
                target_map_coord = None
                for item in scanned_items:
                    if "冒泡" in item["label"]:
                        target_map_coord = item["coordinate"]
                        break
                
                if target_map_coord:
                    ws_output[target_map_coord].value = ""

                    pil_img = Image.open(uploaded_map_image)
                    pil_img = ImageOps.exif_transpose(pil_img)
                    cropped_img = crop_image_to_ratio(pil_img, 27, 16)
                    
                    img_byte_arr = io.BytesIO()
                    cropped_img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    
                    img = ExcelImage(img_byte_arr)
                    calc_w, calc_h = calculate_cell_pixels(ws_output, target_map_coord)
                    img.width = calc_w
                    img.height = calc_h
                    
                    ws_output.add_image(img, target_map_coord)
            except Exception as e:
                st.warning(f"圖片處理異常: {e}")

        id_coord = None
        name_coord = None
        for item in scanned_items:
            if "委託" in item["label"] and "編號" in item["label"]:
                id_coord = item["coordinate"]
            if "案名" in item["label"]:
                name_coord = item["coordinate"]
        
        file_id = user_inputs.get(id_coord, "無編號") if id_coord else "無編號"
        file_name = user_inputs.get(name_coord, "無案名") if name_coord else "無案名"
        safe_filename = f"{file_id}{file_name}.xlsx"
        safe_filename = "".join([c for c in safe_filename if c.isalpha() or c.isdigit() or c in " ._-()[\u4e00-\u9fa5]"])

        output_buffer = io.BytesIO()
        wb_output.save(output_buffer)
        output_buffer.seek(0)

        st.success(f"整合完成 目前已可供下載Excel：{safe_filename}")
        
        st.download_button(
            label="下載Excel檔案",
            data=output_buffer,
            file_name=safe_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
