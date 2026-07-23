# 歌詞拼音工具 — 規範與限制

本文件整理這個工具的處理規範、已知限制與部署需求,供操作與維護參考。

---

## 一、處理規範(工具的行為規則)

- **只處理主歌詞**:僅對「字級 ≥ `min_pt`(預設 40pt)且含中文字元」的段落加拼音。藉此自動略過標題(約 14pt)、footer(約 10pt)、底部重複字幕(約 28pt)。
- **拼音位置**:放在中文行的**正下方**,每個音節對齊到對應中文字的中心(類似 ruby 注音)。
- **對齊方式**:用字型度量計算位置、以空格墊出來,**不用 tab stops**(LibreOffice 等播放軟體會忽略 tab stop)。因為 Arial 與 Liberation Sans 字寬相同,PowerPoint 與 LibreOffice 渲染結果一致。
- **拼音字級**:絕對 pt 值(預設 20pt),在 Advanced settings 可調。
- **純函式設計**:`add_pinyin(pptx) → pptx`,不改動輸入檔,同一份檔可重複處理。
- **內建讀音修正**:`祢 → nǐ`、`尊主為大 → wéi`。程式裡 `CHAR_OVERRIDES` / `PHRASE_OVERRIDES` 兩個 dict 可自行擴充。
- **原檔完整保留**:只「新增」拼音段落,原有段落的文字、字級、字型、粗體、顏色一字不動。已用逐段格式指紋比對驗證(96 段全部一致),背景圖與 theme 主題檔逐 byte 相同。

---

## 二、已知限制

- **多音字需人工把關**:pypinyin 對部分字會判錯(例:`祢` 曾判成 mí、`為` 判成 wèi)。常見要注意的字:行、樂、降、還、為、祢。發現錯的就加一條 override。
- **字級門檻是啟發式**:歌詞若低於 40pt 不會被處理,要在 Advanced settings 調低 `min_pt`;但門檻調太低(低於約 29pt),底部那條 28pt 重複字幕也會被加上拼音。安全區間是「最大的非歌詞文字」與「最小的歌詞」之間。
- **相鄰寬音節無法完全置中**:像 `shēng mìng` 這種兩個寬音節相鄰時,物理上無法同時精準置中,演算法會讓它們對稱地微幅錯開,最差偏差約 1/4 字寬,視覺上仍明確落在對應字下方。這也是拼音字級不宜設太大的原因。
- **絕對 pt 不隨行縮放**:整份檔所有拼音同一大小,不會依各行中文字級自動縮放。這份檔主歌詞都是 48pt,所以沒差;若某份檔各行中文字級不同,拼音不會跟著變。
- **python-pptx 會重寫 XML**:存檔時會改寫 XML 格式(namespace 順序、空白、標籤寫法),功能上無害,但用檔案層級 diff 會看到很多 slide/layout 顯示「changed」。若要求真正最小 diff,需直接操作 XML(程式會複雜不少,目前不採用)。
- **UI 樣式依賴 Streamlit 內部選擇器**:美化是靠注入 CSS 命中 Streamlit 的 `data-testid` 選擇器。這些選擇器穩定,但 Streamlit 若大改版動到它們,拖曳區或按鈕樣式可能需微調(通常幾行就能修)。

---

## 三、部署規範(Streamlit Community Cloud)

**Repo 根目錄需放 4 個檔案(檔名與內容都不要改):**

| 檔案 | 用途 | 內容 |
|------|------|------|
| `app.py` | Streamlit UI(**Main file path 設這個**) | 介面 + 呼叫處理邏輯 |
| `pinyin_pptx.py` | 核心邏輯模組(被 app.py 匯入,不單獨啟動) | `add_pinyin` 純函式 |
| `requirements.txt` | Python 套件(由 pip 安裝) | `streamlit`、`python-pptx`、`pypinyin`、`fonttools` |
| `packages.txt` | 系統套件(由 apt 安裝) | `fonts-liberation` |

- **`packages.txt` 必要**:拼音對齊要靠 Liberation Sans 的字寬計算,那是系統字型、不是 pip 套件,少了它對齊會失準。
- **部署設定**:Repository → 你的 repo;Branch → main;Main file path → `app.py`。按 Deploy 後 Streamlit 會先讀 `requirements.txt` 與 `packages.txt` 裝好環境,再跑 `app.py`。
- **顯示字型**:Sora / Inter 由使用者瀏覽器向 Google Fonts 載入(client-side),部署環境不需額外安裝。

**免費層須知:**
- 記憶體約 1GB(處理幾 MB 的 PPTX 綽綽有餘)。
- App 閒置一段時間會休眠,下次開啟需等十幾秒喚醒。
- 預設是**公開網址**,任何知道連結的人都能用。歌詞檔不敏感通常無所謂;要限制可在 app 加密碼,或用 Streamlit 的 email 白名單。

---

## 四、每次處理新歌的人工檢查

1. 掃一眼多音字(行、樂、降、還、為、祢 等),讀音錯的加 override。
2. 確認歌詞字級;若不是 40pt 以上,在 Advanced settings 調 `min_pt`。
3. 打開處理後的檔案,確認拼音沒有溢出背景、對齊正常。