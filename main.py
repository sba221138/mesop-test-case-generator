import mesop as me
import io
import docx
import base64
import json
import csv
import os
from dotenv import load_dotenv
from pathlib import Path
from pdf2image import convert_from_bytes
from dataclasses import field
from PIL import Image  # 用來處理圖片給 AI 看
import google.generativeai as genai # Google AI 套件
import typing_extensions as typing # 用來定義 JSON 結構
import pandas as pd
import urllib.parse

# ==============================================================================
# 🔑 設定區 (請在此填入你的 API Key)
# ==============================================================================
# 1. 載入 .env 檔案裡的設定
load_dotenv()

# 2. 從環境變數中拿取 API_KEY
API_KEY = os.getenv("API_KEY")

# 檢查有沒有拿到 Key (這行是為了保險起見)
if not API_KEY:
    print("❌ 錯誤：找不到 API_KEY！請檢查 .env 檔案。")
else:
    genai.configure(api_key=API_KEY)

# ==============================================================================
# 📋 定義資料結構 (Schema)
# 這是給 AI 看的：告訴它我们要什麼欄位
# ==============================================================================
class TestCase(typing.TypedDict):
    id: str
    title: str
    pre_condition: str
    steps: str
    expected_result: str
    remarks: str      # 備註
    test_result: str  # 測試結果 (預期 AI 留白)
    test_time: str    # 測試時間 (預期 AI 留白)
    tester: str       # 測試人員 (預期 AI 留白)

# ==============================================================================

@me.stateclass
class State:
    file_name: str = ""
    file_content: str = ""
    pdf_images: list[str] = field(default_factory=list)
    
    # 改用 List 存結構化資料，不再只是純文字
    test_cases: list[dict] = field(default_factory=list)
    
    is_loading: bool = False
    error_msg: str = ""

# --- 工具函數區 ---

def get_poppler_path():
    """動態取得專案內建的 Poppler bin 資料夾路徑"""
    base_path = Path(__file__).parent.absolute()
    poppler_bin_path = base_path / "bin" / "poppler-25.12.0" / "Library" / "bin"
    if not poppler_bin_path.exists(): return None
    return str(poppler_bin_path)

def image_to_base64(pil_image):
    """PIL Image -> Base64 (給瀏覽器看)"""
    buffered = io.BytesIO()
    pil_image.save(buffered, format="JPEG", quality=85)
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
    return f"data:image/jpeg;base64,{img_str}"

def base64_to_image(base64_str):
    """Base64 -> PIL Image (給 AI 看)"""
    # 1. 移除開頭的 "data:image/jpeg;base64," 這段宣告
    if "," in base64_str:
        base64_str = base64_str.split(",")[1]
    
    # 2. 解碼回二進位數據
    image_data = base64.b64decode(base64_str)
    
    # 3. 變回 PIL 圖片物件
    return Image.open(io.BytesIO(image_data))

def get_docx_text(file_bytes):
    doc_file = io.BytesIO(file_bytes)
    doc = docx.Document(doc_file)
    return "\n".join([para.text for para in doc.paragraphs])

# --- CSV 下載邏輯 (重點！) ---

def json_to_csv(data_list):
    """將英文 Key 的 JSON 轉換為中文 Header 的 CSV"""
    if not data_list:
        return ""
    
    output = io.StringIO()
    # 定義中文表頭對照表
    headers_map = {
        "id": "測試編號",
        "title": "測試標題",
        "pre_condition": "預置條件",
        "steps": "測試步驟",
        "expected_result": "預期結果",
        "remarks": "備註",
        "test_result": "測試結果",
        "test_time": "測試時間",
        "tester": "測試人員"
    }
    
    # 寫入 CSV (加上 BOM \ufeff 讓 Excel 開啟時不會亂碼)
    output.write('\ufeff') 
    writer = csv.DictWriter(output, fieldnames=headers_map.keys())
    
    # 1. 寫入中文表頭 (把英文 key 換成中文顯示)
    writer.writerow(headers_map) 
    
    # 2. 寫入資料
    writer.writerows(data_list)
    
    return output.getvalue()

# --- 事件處理區 (Handlers) ---

def handle_upload(event: me.UploadEvent):
    state = me.state(State)
    state.file_name = event.file.name
    state.pdf_images = []
    state.file_content = ""
    state.ai_response = "" # 清空上次的 AI 回答
    
    content_bytes = event.file.getvalue()
    filename = event.file.name.lower()
    
    try:
        if filename.endswith(".docx"):
            state.file_content = get_docx_text(content_bytes)
            
        elif filename.endswith(".pdf"):
            poppler_path = get_poppler_path()
            if poppler_path:
                images = convert_from_bytes(content_bytes, poppler_path=poppler_path, fmt="jpeg", dpi=150)
                for img in images:
                    state.pdf_images.append(image_to_base64(img))
                state.file_content = f"✅ 已載入 {len(images)} 頁 PDF 圖片。"
            else:
                state.file_content = "❌ 錯誤：找不到 Poppler。"
        else:
            state.file_content = content_bytes.decode("utf-8")
    except Exception as e:
        state.file_content = f"❌ 讀取失敗：{e}"

def generate_test_cases(event: me.ClickEvent):
    """呼叫 AI 生成測試案例的主函數"""
    state = me.state(State)
    
    # 1. 檢查有沒有東西可以送
    if not state.file_content and not state.pdf_images:
        state.ai_response = "請先上傳檔案！"
        return

    # 2. 開啟讀取中狀態 (讓按鈕轉圈圈)
    state.is_loading = True
    state.error_msg = ""
    state.test_cases = [] # 清空
    yield # 先讓畫面更新一次，顯示轉圈圈

    try:
        # 3. 準備 AI 模型 (使用 Gemini 2.5 Flash，速度快且便宜)
        model = genai.GenerativeModel('models/gemini-2.5-flash')
        
        # 4. 準備提示詞 (Prompt)
        prompt = """
        你是一個 QA 專家。請根據輸入內容生成測試案例。
        請注意：
        1. 直接回傳 JSON 格式。
        2. 'remarks', 'test_result', 'test_time', 'tester' 欄位請留空字串。
        3. 測試步驟請用條列式。
        """

        # 5. 準備要餵給 AI 的資料包
        input_data = [prompt]
        
        # 如果有文字內容，加進去
        if state.file_content:
            input_data.append(f"需求文件內容：\n{state.file_content}")
            
        # 如果有圖片 (PDF)，把它們轉回 PIL 格式並加進去
        if state.pdf_images:
            for b64_img in state.pdf_images:
                input_data.append(base64_to_image(b64_img))

        # 6. 發送給 Gemini (這步會花幾秒鐘)
        response = model.generate_content(input_data,
            generation_config=genai.GenerationConfig(
            response_mime_type="application/json", 
            response_schema=list[TestCase] # 強制 AI 符合我們的結構
            )
        )
        state.test_cases = json.loads(response.text) # 直接把 AI 回傳的 JSON 存起來
        # for chunk in response:
        #     # 7. 接收結果
        #     state.ai_response += chunk.text
        #     yield # 告訴 Mesop：「嘿，我有新文字了，快更新畫面！」

    except Exception as e:
        state.ai_response = f"❌ AI 生成失敗：{str(e)}"
    
    finally:
        # 8. 關閉讀取中狀態
        state.is_loading = False
        yield

# --- 畫面區 ---

@me.page(path="/", security_policy=me.SecurityPolicy(dangerously_disable_trusted_types=True))
def page():
    state = me.state(State)
    
    with me.box(style=me.Style(padding=me.Padding.all(20), background="#f8f9fa", min_height="100vh")):
        me.text("🤖 AI 測試案例生成器 (Excel 下載版)", type="headline-3", style=me.Style(color="#1a73e8"))
        
        # 上傳區
        with me.box(style=me.Style(background="white", padding=me.Padding.all(20), border_radius=12, margin=me.Margin(bottom=20))):
            me.uploader(
                label="📁 上傳文件",
                on_upload=handle_upload,
                accepted_file_types=[".docx", "application/pdf", "text/plain"]
            )
            if state.file_content: me.text(state.file_content[:100] + "..." if len(state.file_content)>100 else state.file_content, style=me.Style(color="green"))

        # 按鈕區
        with me.box(style=me.Style(display="flex", gap=10, justify_content="center", margin=me.Margin(bottom=20))):
            if state.is_loading:
                me.progress_spinner()
                me.text(" AI 正在分析結構中...")
            else:
                me.button("✨ 生成表格", on_click=generate_test_cases, type="flat", style=me.Style(background="#1a73e8", color="white"))
                
                # 如果有生成結果，才顯示下載按鈕
                if state.test_cases:
                    # 1. 現場把 JSON 轉成 CSV 字串
                    csv_string = json_to_csv(state.test_cases)
                    
                    # 2. 把 CSV 字串編碼成 URL 安全格式 (處理空格、換行等)
                    csv_encoded = urllib.parse.quote(csv_string)
                    
                    # 3. 組合出 Data URI
                    data_uri = f"data:text/csv;charset=utf-8,{csv_encoded}"
                    
                    # 4. 使用 me.html 畫一個原生的 HTML 連結
                    # download="filename.csv" 屬性會告訴瀏覽器這是一個下載行為
                    with me.box(style=me.Style(
                        background="white",
                        color="#1a73e8",
                        border=me.Border.all(me.BorderSide(width=1, color="#1a73e8")),
                        padding=me.Padding.symmetric(vertical=10, horizontal=16),
                        border_radius=18,
                        cursor="pointer",
                        font_weight="500",
                        font_size=14,
                        display="inline-block" # 讓它像按鈕一樣並排
                    )):
                        # 內層：只放最單純的 HTML 連結
                        # 我們把 mode="sandboxed" 拿掉，讓下載功能恢復
                        me.html(
                            f'<a href="{data_uri}" download="test_cases.csv" style="text-decoration:none; color:inherit;">📥 下載 CSV (Excel)</a>'
                        )

        if state.error_msg:
            me.text(state.error_msg, style=me.Style(color="red"))

        # 結果顯示區 (表格化)
        if state.test_cases:
            me.text(f"已生成 {len(state.test_cases)} 筆測試案例：", type="headline-6")
            
            # 使用 Table 元件顯示
            with me.box(style=me.Style(background="white", padding=me.Padding.all(15), border_radius=8, overflow_x="auto")):
                # 1. 先將資料轉成 Pandas DataFrame
                df = pd.DataFrame(state.test_cases)
                
                # 2. 【關鍵一步】直接把資料欄位改名！(Mapping)
                # 這樣表格元件就會以為原本的欄位名就是中文
                df = df.rename(columns={
                    "id": "編號",
                    "title": "標題",
                    "pre_condition": "預置條件",
                    "steps": "測試步驟",
                    "expected_result": "預期結果",
                    "remarks": "備註",
                    "test_result": "測試結果",
                    "test_time": "測試時間",
                    "tester": "測試人員"
                })

                # 3. 顯示表格
                me.table(
                    data_frame=df, # 餵入已經改名完成的 df
                    columns={
                        # 注意：這裡的 Key 必須跟上面 rename 後的中文名稱一模一樣
                        "編號": me.TableColumn(),
                        "標題": me.TableColumn(),
                        "預置條件": me.TableColumn(),
                        "測試步驟": me.TableColumn(),
                        "預期結果": me.TableColumn(),
                        "備註": me.TableColumn(),
                        "測試結果": me.TableColumn(),
                        "測試時間": me.TableColumn(),
                        "測試人員": me.TableColumn(),
                    }
                )