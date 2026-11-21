# 建案查詢 LINE 機器人

- `Code.gs`：Google Apps Script 程式（綁在「建案查詢機器人」試算表），員工輸入案名 → 回傳 Flex 卡片。
- `build_bot_data.py`：從 `index.html` 產生 `index.json` 與 `p/*.json`。**地圖資料更新後要重跑並一起推上來**，機器人才會是最新資料。
- 地圖網址支援 `?n=案名&d=行政區`，直接打開該建案。
