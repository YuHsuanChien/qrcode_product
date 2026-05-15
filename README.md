# QR Code 產生工具

從 Excel 員工名單自動產生 QR Code 圖片，並將 QR Code 嵌入回 Excel 檔案。

## 環境需求

- Node.js >= 14
- npm

## 安裝

```bash
npm install
```

## Excel 檔案格式

將 Excel 檔案命名為 `record.xlsx` 放在專案根目錄，並確保包含一個名為 **「夥伴名單」** 的工作表。

欄位格式（第 1 行為標題列）：

| 欄位 | 內容       | 說明                                  |
| ---- | ---------- | ------------------------------------- |
| A    | id         | 員工 ID（用於產生 QR Code，不可為空） |
| B    | code       | 代碼                                  |
| C    | acount     | 帳號                                  |
| D    | staff_name | 姓名                                  |
| E    | family     | 家族                                  |
| F    | team       | 組別                                  |

## 執行

```bash
npm start
```

## 產出結果

| 產出                      | 說明                                       |
| ------------------------- | ------------------------------------------ |
| `member_qrcode/`          | 每位員工的 QR Code 圖片（`{id}.png`）      |
| `record_with_qrcode.xlsx` | 新的 Excel 檔案，G 欄插入對應 QR Code 圖片 |

原始 `record.xlsx` 不會被修改。

## 注意事項

- Excel 檔案必須命名為 `record.xlsx`
- 工作表名稱必須為「夥伴名單」
- A 欄 (id) 為空的行會被自動過濾
- 每次執行會清空 `member_qrcode/` 資料夾後重新產生

## 專案結構

```
├── record.xlsx              # 輸入的 Excel 檔案
├── record_with_qrcode.xlsx  # 產出的 Excel（含 QR Code 圖片）
├── member_qrcode/           # 產出的 QR Code 圖片資料夾
├── src/
│   ├── index.ts             # 主程式進入點
│   └── service/
│       ├── qrcode.ts            # QR Code 產生服務
│       ├── read_excel_file.ts   # Excel 讀取服務
│       └── write_excel_file.ts  # Excel 寫入服務
├── package.json
└── tsconfig.json
```
