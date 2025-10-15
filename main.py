# -*- coding: utf-8 -*-
"""
ニュース収集 + 追記直後にAI分類（E=ポジネガ, F=カテゴリ）まで行う統合版

■ 既存仕様（維持）
- Googleニュース / Yahoo!ニュース / MSNニュース を Selenium+BS4 で取得
- スプレッドシート（タブ: Google / Yahoo / MSN）へ追記（URLで重複除外）

■ 追加仕様（本修正）
- 追記に成功した「新規行のみ」を対象に、Geminiで
    E列: ポジネガ（ポジティブ / ネガティブ / ニュートラル）
    F列: カテゴリ（定義ルールに基づく1カテゴリ）
  を即時に書き込む

■ 必要な環境変数
- GOOGLE_CREDENTIALS : サービスアカウントJSON（文字列）
  （または同階層に credentials.json を置く既存方式でも可）
- GEMINI_API_KEY      : Gemini APIキー（分類を実行する場合）

■ 追加が望ましいパッケージ（requirements.txt）
google-generativeai
"""

import os
import json
import time
import re
import random
import requests
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

# --- 追加: Gemini（インストールされていなければ分類はスキップ）
try:
    import google.generativeai as genai
except Exception:
    genai = None

# ====== 既存の設定 ======
KEYWORD = "トヨタ"
SPREADSHEET_ID = "1N7sZGtFnvICo6yBkw0WHtuYG3-lKYO0_p4-IGOcPjOo"

# ====== 共通ユーティリティ ======
def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    pub_label = pub_label.strip().lower()
    try:
        if "分前" in pub_label or "minute" in pub_label:
            m = re.search(r"(\d+)", pub_label)
            if m:
                dt = base_time - timedelta(minutes=int(m.group(1)))
                return format_datetime(dt)
        elif "時間前" in pub_label or "hour" in pub_label:
            h = re.search(r"(\d+)", pub_label)
            if h:
                dt = base_time - timedelta(hours=int(h.group(1)))
                return format_datetime(dt)
        elif "日前" in pub_label or "day" in pub_label:
            d = re.search(r"(\d+)", pub_label)
            if d:
                dt = base_time - timedelta(days=int(d.group(1)))
                return format_datetime(dt)
        elif re.match(r'\d+月\d+日', pub_label):
            dt = datetime.strptime(f"{base_time.year}年{pub_label}", "%Y年%m月%d日")
            return format_datetime(dt)
        elif re.match(r'\d{4}/\d{1,2}/\d{1,2}', pub_label):
            dt = datetime.strptime(pub_label, "%Y/%m/%d")
            return format_datetime(dt)
        elif re.match(r'\d{1,2}:\d{2}', pub_label):
            t = datetime.strptime(pub_label, "%H:%M").time()
            dt = datetime.combine(base_time.date(), t)
            if dt > base_time:
                dt -= timedelta(days=1)
            return format_datetime(dt)
    except:
        pass
    return "取得不可"

def get_last_modified_datetime(url):
    try:
        response = requests.head(url, timeout=5)
        if 'Last-Modified' in response.headers:
            dt = parsedate_to_datetime(response.headers['Last-Modified'])
            jst = dt.astimezone(tz=timedelta(hours=9))
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

# ====== スクレイピング（既存） ======
def get_google_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    articles = soup.find_all("article")
    data = []
    for article in articles:
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            source_tag = article.select_one("div.vr1PYe")
            title = a_tag.text.strip()
            href = a_tag.get("href")
            url = "https://news.google.com" + href[1:] if href.startswith("./") else href
            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)
            source = source_tag.text.strip() if source_tag else "N/A"
            data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source})
        except:
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件")
    return data

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    driver.get(search_url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()
    articles = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))
    articles_data = []

    for article in articles:
        try:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            title = title_tag.text.strip() if title_tag else ""
            link_tag = article.find("a", href=True)
            url = link_tag["href"] if link_tag else ""
            time_tag = article.find("time")
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = ""
            if date_str:
                date_str = re.sub(r'\([月火水木金土日]\)', '', date_str).strip()
                try:
                    dt_obj = datetime.strptime(date_str, "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt_obj)
                except:
                    formatted_date = date_str

            source_text = ""
            source_tag = article.find("div", class_="sc-n3vj8g-0 yoLqH")
            if source_tag:
                inner = source_tag.find("div", class_="sc-110wjhy-8 bsEjY")
                if inner and inner.span:
                    candidate = inner.span.text.strip()
                    if not candidate.isdigit():
                        source_text = candidate
            if not source_text or source_text.isdigit():
                alt_spans = article.find_all(["span", "div"], string=True)
                for s in alt_spans:
                    text = s.text.strip()
                    if 2 <= len(text) <= 20 and not text.isdigit() and re.search(r'[ぁ-んァ-ン一-龥A-Za-z]', text):
                        source_text = text
                        break

            if title and url:
                articles_data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": formatted_date if formatted_date else "取得不可",
                    "引用元": source_text
                })
        except:
            continue

    print(f"✅ Yahoo!ニュース件数: {len(articles_data)} 件")
    return articles_data

def get_msn_news_with_selenium(keyword: str) -> list[dict]:
    now = datetime.utcnow() + timedelta(hours=9)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d'1'&form=YFNR"
    driver.get(url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()
    cards = soup.select("div.news-card")
    data = []

    for card in cards:
        try:
            title = card.get("data-title", "").strip()
            url = card.get("data-url", "").strip()
            source = card.get("data-author", "").strip()
            pub_label = ""

            pub_tag = card.find("span", attrs={"aria-label": True})
            if pub_tag and pub_tag.has_attr("aria-label"):
                pub_label = pub_tag["aria-label"].strip().lower()

            pub_date = parse_relative_time(pub_label, now)
            if pub_date == "取得不可" and url:
                pub_date = get_last_modified_datetime(url)

            if title and url:
                data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": pub_date,
                    "引用元": source if source else "MSN"
                })
        except Exception as e:
            print(f"⚠️ MSN記事処理エラー: {e}")
            continue

    print(f"✅ MSNニュース件数: {len(data)} 件")
    return data

# ====== 追記＋新規行インデックスの取得（修正） ======
def ensure_headers_and_get_existing(gc_sheet, worksheet_name: str):
    """
    ワークシートを開き、なければ新規作成。
    ヘッダー（A:タイトル, B:URL, C:投稿日, D:引用元, E:ポジネガ, F:カテゴリ）を保証。
    既存データを取得して返す。
    """
    try:
        worksheet = gc_sheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = gc_sheet.add_worksheet(title=worksheet_name, rows="1", cols="6")
        worksheet.append_row(['タイトル', 'URL', '投稿日', '引用元', 'ポジネガ', 'カテゴリ'])
        return worksheet, [['タイトル', 'URL', '投稿日', '引用元', 'ポジネガ', 'カテゴリ']]

    # 既存ヘッダ確認＆E/Fが無ければ補完
    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(['タイトル', 'URL', '投稿日', '引用元', 'ポジネガ', 'カテゴリ'])
        values = [['タイトル', 'URL', '投稿日', '引用元', 'ポジネガ', 'カテゴリ']]
    else:
        header = values[0] if values else []
        changed = False
        # ヘッダの長さを6列以上に拡張
        if len(header) < 6:
            header = (header + [''] * (6 - len(header)))[:6]
            changed = True
        # E/F 名称を設定（空なら埋める）
        if header[4] != 'ポジネガ':
            header[4] = 'ポジネガ'; changed = True
        if header[5] != 'カテゴリ':
            header[5] = 'カテゴリ'; changed = True
        if changed:
            worksheet.update('A1:F1', [header])
            # 最新の全量は再取得しない（行数のみで十分）
    return worksheet, values

def write_to_spreadsheet(articles: list[dict], spreadsheet_id: str, worksheet_name: str):
    """
    追記を行い、今回新規に追加できた行番号のリストを返す（1始まり）。
    """
    credentials_json_str = os.environ.get('GCP_SERVICE_ACCOUNT_KEY') or os.environ.get('GOOGLE_CREDENTIALS')
    credentials = json.loads(credentials_json_str) if credentials_json_str else json.load(open('credentials.json'))
    gc = gspread.service_account_from_dict(credentials)

    # ワークシート確保 + 既存データ取得
    sh = gc.open_by_key(spreadsheet_id)
    worksheet, existing_data = ensure_headers_and_get_existing(sh, worksheet_name)

    existing_urls = set(row[1] for row in existing_data[1:] if len(row) > 1)
    new_rows_payload = [
        [a['タイトル'], a['URL'], a['投稿日'], a['引用元']]
        for a in articles if a['URL'] not in existing_urls
    ]

    if not new_rows_payload:
        print("⚠️ 追記すべき新しいデータはありません。")
        return gc, worksheet, []  # 新規なし

    # 追記前の総行数
    prev_rows = len(existing_data)  # ヘッダ含む行数
    # A〜D列のみ書き込み（E/FはAIで後から埋める）
    worksheet.append_rows(new_rows_payload, value_input_option='USER_ENTERED')
    print(f"✅ {len(new_rows_payload)}件をスプレッドシートに追記しました。")

    # 追記された行のインデックス（1始まり）
    added_row_indices = list(range(prev_rows + 1, prev_rows + len(new_rows_payload) + 1))
    return gc, worksheet, added_row_indices

# ====== 追記直後の新規行だけをGeminiで分類して E/F に反映（追加） ======
GEMINI_PROMPT = """
あなたは敏腕雑誌記者です。Webニュースの「タイトル」だけを見て、次を厳密に分類してください。

【1】ポジネガ判定（必ず次のいずれか一語のみ）：
- ポジティブ
- ネガティブ
- ニュートラル

【2】記事のカテゴリー判定（最も関連が高い1つだけを選ぶ。並記禁止）：
- 会社：企業の施策や生産、販売台数など。ニッサン、トヨタ、ホンダ、スバル、マツダ、スズキ、ミツビシ、ダイハツの記事の場合は () 付きで企業名を記載。それ以外は「その他」。
- 車：クルマの名称が含まれているもの（会社名だけの場合は「車」に分類しない）。新型/現行/旧型 + 名称 を () 付きで記載（例：新型リーフ、現行セレナ、旧型スカイライン）。日産以外の車の場合は「車（競合）」と記載。
- 技術（EV）：電気自動車の技術に関わるもの（ただしバッテリー工場建設や企業の施策は含まない）。
- 技術（e-POWER）：e-POWERに関わるもの。
- 技術（e-4ORCE）：4WD/2WD/AWDに関わるもの。
- 技術（AD/ADAS）：自動運転・先進運転支援に関わるもの。
- 技術：上記以外の技術。
- モータースポーツ：F1やラリー、フォーミュラEなど自動車レース。
- 株式：株式発行や株価の値動き、投資に関わるもの。
- 政治・経済：政治家や選挙、税金、経済に関わるもの。
- スポーツ：自動車以外のスポーツ。
- その他：上記に含まれないもの。

【出力要件】
- **JSON配列**のみを返してください（余計な文章や注釈を含めない）。
- 各要素は次の形式：
  {"row": 行番号, "sentiment": "ポジティブ|ネガティブ|ニュートラル", "category": "カテゴリ名"}
- 入力の「タイトル」文字列は変更しないこと（出力には含めなくて良い）。
""".strip()

def classify_rows_with_gemini(worksheet: gspread.Worksheet, row_indices: list[int]):
    """
    指定された行番号（1始まり。ヘッダ=1）のうち、データ行のみを対象に
    A列（タイトル）からE/F列をGeminiで埋める。
    """
    if not row_indices:
        return

    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key or genai is None:
        print("ℹ Gemini分類はスキップ（APIキー未設定 or ライブラリ未インストール）。")
        return

    genai.configure(api_key=api_key)

    # ヘッダ行(1)は除外
    target_rows = [r for r in row_indices if r > 1]
    if not target_rows:
        return

    # タイトル取得（A列）
    titles = worksheet.batch_get([f"A{r}" for r in target_rows])
    # batch_getの戻りは [[["タイトル"]], [["タイトル"]], ...] の形になることが多いので整形
    flat_titles = []
    for cell in titles:
        val = ""
        if cell and len(cell) > 0 and len(cell[0]) > 0:
            val = cell[0][0]
        flat_titles.append(val)

    # Geminiへ投げるペイロード
    items = [{"row": r, "title": t} for r, t in zip(target_rows, flat_titles) if t]

    if not items:
        print("ℹ 分類対象にタイトルがありません。")
        return

    # まとめて送る（40件程度ずつ）
    BATCH = 40
    updates = []
    for i in range(0, len(items), BATCH):
        batch = items[i:i+BATCH]
        prompt = GEMINI_PROMPT + "\n\n" + json.dumps(batch, ensure_ascii=False, indent=2)

        try:
            model = genai.GenerativeModel("gemini-1.5-flash-latest")
            resp = model.generate_content(prompt)
            text = (resp.text or "").strip()

            # JSON抽出（安全策）
            import re as _re
            m = _re.search(r"\[.*\]", text, flags=_re.DOTALL)
            json_text = m.group(0) if m else text
            result = json.loads(json_text)

            for obj in result:
                try:
                    row_idx = int(obj.get("row"))
                except Exception:
                    continue
                sentiment = str(obj.get("sentiment", "")).strip()
                category  = str(obj.get("category", "")).strip()

                # 語彙のゆらぎ補正
                if sentiment not in ("ポジティブ", "ネガティブ", "ニュートラル"):
                    if "ポジ" in sentiment:
                        sentiment = "ポジティブ"
                    elif "ネガ" in sentiment:
                        sentiment = "ネガティブ"
                    else:
                        sentiment = "ニュートラル"

                updates.append({
                    "range": f"E{row_idx}:F{row_idx}",  # E=ポジネガ, F=カテゴリ
                    "values": [[sentiment, category]]
                })
        except Exception as e:
            print(f"⚠ Gemini応答の解析に失敗: {e}")

    if updates:
        worksheet.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"✨ Gemini分類を {len(updates)} 行に反映しました。")
    else:
        print("ℹ Gemini分類の更新はありませんでした。")

# ====== メイン処理（既存 + 追記後分類を呼び出し） ======
def process_one_source(articles: list[dict], sheet_id: str, tab_name: str):
    """
    1) スプレッドシートへ追記（重複除外）
    2) 今回追加された行だけ Gemini で E/F を埋める
    """
    if not articles:
        print(f"（{tab_name}）新規記事なし")
        return

    # 追記 & 追加行インデックス取得
    gc, worksheet, added_rows = write_to_spreadsheet(articles, sheet_id, tab_name)
    # 追加行のみ分類
    classify_rows_with_gemini(worksheet, added_rows)

if __name__ == "__main__":
    print("\n--- Google News ---")
    google_news_articles = get_google_news_with_selenium(KEYWORD)
    process_one_source(google_news_articles, SPREADSHEET_ID, "Google")

    print("\n--- Yahoo! News ---")
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    process_one_source(yahoo_news_articles, SPREADSHEET_ID, "Yahoo")

    print("\n--- MSN News ---")
    msn_news_articles = get_msn_news_with_selenium(KEYWORD)
    process_one_source(msn_news_articles, SPREADSHEET_ID, "MSN")
