import streamlit as st

def inject_custom_styles():
    # 這是你 GitHub 上圖片的原始連結
    icon_url = "https://raw.githubusercontent.com/vary530/survey_app/main/my_logo.png" 

    # 針對 iOS Safari 的特殊 Meta 設定
    st.markdown(f"""
        <meta name="theme-color" content="#050505">
        <meta name="apple-mobile-web-app-title" content="永義物調"> 
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover, interactive-widget=resizes-content">
        
        <link rel="apple-touch-icon" href="{icon_url}">
        <link rel="apple-touch-icon" sizes="152x152" href="{icon_url}">
        <link rel="apple-touch-icon" sizes="180x180" href="{icon_url}">
        <link rel="icon" type="image/png" href="{icon_url}">
    """, unsafe_allow_html=True)

    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');
            
            /* --- 0. 核彈級隱藏卷軸 --- */
            ::-webkit-scrollbar {
                display: none !important;
                width: 0px !important;
                height: 0px !important;
                background: transparent !important;
            }
            
            * {
                -ms-overflow-style: none !important;
                scrollbar-width: none !important;
            }

            /* --- 1. 全螢幕鎖定與防彈跳 --- */
            :root {
                --background-color: #050505;
            }

            html {
                background-color: var(--background-color) !important;
                overscroll-behavior: none !important;
                height: 100dvh !important;
                width: 100vw !important;
                overflow: hidden !important;
                position: fixed !important;
                top: 0;
                left: 0;
            }

            body {
                background-color: var(--background-color) !important;
                overscroll-behavior: none !important;
                height: 100dvh !important;
                width: 100vw !important;
                overflow: hidden !important;
                margin: 0 !important;
                padding: 0 !important;
                position: fixed !important;
                top: 0;
                left: 0;
                touch-action: none;
            }
            
            /* --- 2. 【關鍵修改】徹底移除連結按鈕與介面雜訊 --- */
            
            /* 針對標題旁的連結圖示 (Anchor Link) */
            /* 使用 display: none 讓它在佈局樹中完全消失，不佔空間 */
            .stApp a.anchor-link, 
            .stMarkdown h1 a, 
            h1 > a, 
            a[href^="#"][class*="anchor"] {
                display: none !important;
                visibility: hidden !important;
                width: 0 !important;
                height: 0 !important;
                opacity: 0 !important;
                pointer-events: none !important;
                position: absolute !important; /* 確保它不會佔位 */
                left: -9999px !important;
            }

            /* 移除上方白條 (Header) */
            header, [data-testid="stHeader"], .stAppHeader {
                display: none !important;
                height: 0px !important;
                opacity: 0 !important;
                pointer-events: none !important;
            }

            /* 移除下方 Footer */
            footer, [data-testid="stFooter"], .stFooter {
                display: none !important;
                height: 0px !important;
            }

            /* 移除選單與開發者工具 */
            #MainMenu, [data-testid="stToolbar"], 
            [data-testid="stDecoration"], [data-testid="stStatusWidget"],
            .stDeployButton, div[class*="viewerBadge"] {
                display: none !important;
            }

            /* --- 3. APP 主體容器 --- */
            .stApp {
                background-color: var(--background-color);
                background-image: radial-gradient(circle at 50% 0%, #1a1a1a 0%, #050505 80%);
                background-attachment: fixed; 
                background-size: cover;
                
                position: fixed !important;
                top: 0;
                left: 0;
                width: 100vw !important;
                height: 100dvh !important;
                
                overflow-y: auto !important;
                overflow-x: hidden !important;
                
                margin-top: 0px !important;
                padding-top: 0px !important;
                
                -webkit-overflow-scrolling: touch; 
                overscroll-behavior: none !important;
                overscroll-behavior-y: none !important;
                
                touch-action: pan-y; 
                z-index: 1;
            }

            /* --- 4. 內容區域 --- */
            .block-container { 
                padding-top: max(1rem, env(safe-area-inset-top)) !important;
                padding-left: 1rem !important;
                padding-right: 1rem !important;
                padding-bottom: max(5rem, env(safe-area-inset-bottom)) !important;
                max-width: 600px; 
                margin: 0 auto;
                width: 100%;
            }
            
            /* Standalone 模式修正 */
            @media all and (display-mode: standalone) {
                body {
                    background-color: #050505 !important;
                }
                .block-container {
                    padding-top: max(20px, env(safe-area-inset-top)) !important;
                }
            }

            /* --- 5. 元件樣式 --- */
            * {
                box-sizing: border-box;
                -webkit-tap-highlight-color: transparent;
            }

            [data-testid="stForm"] {
                background: rgba(30, 30, 30, 0.4);
                border: 1px solid rgba(197, 160, 101, 0.2);
                border-radius: 16px; 
                padding: 20px 24px;
                backdrop-filter: blur(12px);
                margin-top: 10px;
                width: 100%;
                box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
            }

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
                pointer-events: none; /* 禁止標題接收點擊事件，徹底杜絕連結 */
                padding-top: 0px !important;
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
            }
            
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
        </style>
    """, unsafe_allow_html=True)