[README.md](https://github.com/user-attachments/files/28398017/README.md)
# Excel 關鍵字替換工具

這是可部署到 Streamlit Community Cloud 的 Excel 線上處理工具。

## 功能

- 上傳 `.xlsx`
- 批次替換儲存格文字
- 批次替換 Excel 超連結網址
- 英文大小寫不區分
- 移除常見追蹤參數，例如 `utm_source`、`fbclid`、`gclid`
- 自動修復常見髒 Excel 檔案，例如內部 XML 含非法控制字元
- 下載檔名維持與上傳檔名相同，不加 processed 或日期

## 部署方式

1. 把 `app.py`、`requirements.txt`、`README.md` 上傳到 GitHub repository。
2. 到 Streamlit Community Cloud。
3. 選擇 Deploy a public app from GitHub。
4. Repository 選你的專案。
5. Branch 選 `main`。
6. Main file path 填：

```text
app.py
```

7. 按 Deploy。

## 注意

- 此工具只支援真正的 `.xlsx`。
- 若檔案內部壞得太嚴重，仍可能需要先用 Excel 開啟後另存。
- 下載檔名會維持上傳檔名；若使用者電腦下載資料夾已有同名檔，瀏覽器可能自動加上 `(1)`，這是瀏覽器行為，程式無法完全禁止。
