# Excel 關鍵字替換工具｜網頁版

這是一個可部署給同事使用的 Streamlit 網頁工具。

## 功能

- 上傳 `.xlsx` Excel 檔案
- 批次替換儲存格文字
- 批次替換 Excel 超連結網址
- 可刪除指定關鍵字，例如：蝦皮、聊聊、shopee
- 可移除常見追蹤參數，例如：utm_source、fbclid、gclid
- 處理後可直接下載新 Excel

## 本機測試

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署成真正網頁版：Streamlit Community Cloud

1. 註冊或登入 GitHub
2. 建立一個新的 GitHub Repository
3. 把這三個檔案上傳到 Repository：
   - `app.py`
   - `requirements.txt`
   - `README.md`
4. 到 Streamlit Community Cloud 建立 App
5. 選擇剛剛的 GitHub Repository
6. Main file path 填：

```text
app.py
```

7. Deploy 後會得到一組網址，可直接給同事使用。

## 替換規則格式

每行一組：

```text
原文字=>新文字
```

刪除關鍵字：

```text
蝦皮=>
聊聊=>
shopee=>
```

替換成其他文字：

```text
蝦皮=>平台
聊聊=>私訊
shopee=>shop
```

## 注意事項

- 目前支援 `.xlsx`，不支援舊版 `.xls`
- 建議先用備份檔測試
- 若部署到免費雲端平台，請避免上傳敏感或機密檔案
