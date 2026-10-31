# 脂寶烤烤樂｜脂本燒肉候位遊戲

網址：https://gthp0222.github.io/storesystem/zhiben-game/

| 分店 | 掃碼網址 |
| --- | --- |
| 員林店 | https://gthp0222.github.io/storesystem/zhiben-game/?s=yl |
| 彰化店 | https://gthp0222.github.io/storesystem/zhiben-game/?s=ch |
| 北屯店 | https://gthp0222.github.io/storesystem/zhiben-game/?s=bt |

客人掃哪一店的 QR Code，遊戲就自動鎖定那一店（排行榜、核銷都算那一店）。

## 檔案

- `index.html`：整個遊戲（單一檔案）
- `qr/`：三店 QR Code、A4 候位海報、A6 桌立牌（PNG 與可直接印刷的 PDF）
- `backend/Code.gs`：排行榜與核銷後端（Google Apps Script）
- `og-image.png`、`icon-*.png`、`manifest.webmanifest`：分享預覽圖與手機桌面圖示

## 第一次上線（GitHub Pages）

1. 檔案推上 `gthp0222/zhiben-game` 的 `main` 分支。
2. GitHub → repo → Settings → Pages → Source 選「Deploy from a branch」→ Branch 選 `main`、資料夾 `/ (root)` → Save。
3. 約 1 分鐘後打開網址確認。

## 接上真正的排行榜與核銷（上線前一定要做）

沒有後端時，排行榜只是示意資料，核銷碼也寫在網頁裡。正式營運請：

1. 建一個 Google 試算表 →「擴充功能 → Apps Script」→ 貼上 `backend/Code.gs`。
2. 執行一次 `setup()`，到「專案設定 → 指令碼屬性」把 `PINS` 改成三店真正的核銷碼，例如 `{"yl":"4826","ch":"7193","bt":"3057"}`。
3. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我／存取權：所有人）→ 複製網址。
4. 在 `index.html` 搜尋 `API_URL: ''`，把網址貼進引號裡；同一段的 `STAFF_PIN` 改成 `{}`（核銷碼改由後端驗證）。
5. 觸發條件 → 新增 → `weeklyWinners` → 時間驅動 → 每週 → 星期一 00:00–01:00，自動產生得獎名單。

## 常調整的設定（在 index.html 搜尋 `const CFG`）

| 設定 | 預設 | 說明 |
| --- | --- | --- |
| `LEVEL_TIME` | 25 | 每關秒數（＝1 局） |
| `CHEF_GOAL` | 15 | 今日玩滿幾局送主廚招待 |
| `TABLE_LIMIT` | 2 | 每桌招待上限（POS 管控） |
| `PIN_MAX_FAIL` / `PIN_LOCK_MIN` | 5 / 5 | 核銷碼錯幾次鎖幾分鐘 |

關卡門檻在 `const TARGETS`，預設 `[250, 400, 560, 740, 940, 1160, 1400, 1660]`。

## 注意

- QR Code 指向上面的固定網址，換網址就要重印。
- 核銷碼請每月更換一次，只告訴店長與服務人員。
