# VC Academy 報名表單 — 後端設定

網站表單（Apply 頁）會把資料送到 Google Apps Script，由它寫進 Google Sheet。
做法與 SELL GLOBAL（training.mvl.biz）相同。

---

## 步驟 1 — 建立 Google Sheet

新建一份 Google Sheet，命名例如 `VC Academy Applications`。
**不需要**手動建立標題列，程式會自動建立。

## 步驟 2 — 貼上 Apps Script

在該 Sheet 中：**擴充功能 → Apps Script**，把預設內容全部刪除，貼上以下程式碼：

```javascript
const SHEET_NAME = 'Applications';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // 首次執行時自動建立工作表與標題列
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Submitted at', 'Name', 'Email', 'Role & company',
        'LinkedIn', 'Background', 'Why', 'Source'
      ]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.name       || '',
      data.email      || '',
      data.role       || '',
      data.linkedin   || '',
      data.background || '',
      data.why        || '',
      data.source     || ''
    ]);

    // 每筆申請寄一封通知信（不需要就把這行連同下面 3 行刪掉）
    MailApp.sendEmail({
      to: 'vh@mvl.biz',
      subject: 'VC Academy — new application: ' + (data.name || 'unknown'),
      body: [
        'Name:       ' + (data.name || ''),
        'Email:      ' + (data.email || ''),
        'Role:       ' + (data.role || ''),
        'LinkedIn:   ' + (data.linkedin || ''),
        'Background: ' + (data.background || ''),
        'Source:     ' + (data.source || ''),
        '',
        'Why they want to join:',
        (data.why || '')
      ].join('\n')
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

> 通知信收件者目前設為 `vh@mvl.biz`。要改成別的信箱請直接改那一行；
> 要同時寄給多人，用逗號分隔即可：`'vh@mvl.biz, kh@mvl.biz, jy@mvl.biz'`。

## 步驟 3 — 部署

1. 右上角 **部署 → 新增部署作業**
2. 齒輪圖示 → 選 **網頁應用程式**
3. 設定：
   - **執行身分**：我（你的 Google 帳號）
   - **誰可以存取**：**任何人** ← 這項很重要，否則訪客送不出來
4. 按 **部署**，第一次會要求授權，照畫面允許即可
5. 複製產生的網址（結尾是 `/exec`）

## 步驟 4 — 填進網站

把網址給我，或自己編輯 `index.html`，找到這一行填入：

```javascript
const APPLY_URL = "";   // ← 貼在這裡
```

填好之後送出鈕就會啟用。填之前送出鈕是停用狀態，並顯示提醒，不會靜默丟失資料。

---

## 表單欄位對照

網站表單送出的 JSON 欄位名稱如下（與 Apps Script 的 `data.xxx` 對應）：

| JSON 欄位 | 表單標籤 | 必填 | 型態 |
|---|---|---|---|
| `name` | Full name | ✅ | 短文字 |
| `email` | Email | ✅ | 短文字（含格式驗證） |
| `role` | Current role & company | | 短文字 |
| `linkedin` | LinkedIn | | 短文字 |
| `background` | Which best describes you? | ✅ | 下拉選單 |
| `why` | Why do you want to join the founding cohort? | ✅ | 長文字 |
| `source` | How did you hear about VC Academy? | | 下拉選單 |
| `timestamp` | （自動帶入送出時間） | | ISO 字串 |

### 下拉選單選項

**Which best describes you?**
- Student / recent graduate
- Analyst or associate
- Operator moving into investing
- Founder
- Angel / family office
- Working investor
- Other

**How did you hear about VC Academy?**
- The Aug 14 launch event
- Mosaic Venture Lab
- A friend or colleague
- LinkedIn
- Luma
- Other

---

## 備案：獨立腳本（Sheet 開不了 Apps Script 時用）

若「擴充功能 → Apps Script」出現 **「很抱歉，目前無法開啟這個檔案」**，
多半是多個 Google 帳號同時登入導致帳號索引錯誤。可改用獨立腳本：

**1. 先取得 Sheet ID** —— 從 Sheet 網址中間那段：

```
https://docs.google.com/spreadsheets/d/【這一段就是 SHEET_ID】/edit#gid=0
```

**2. 到 [script.google.com/create](https://script.google.com/create) 建立新專案**，貼上：

```javascript
const SHEET_ID   = '把上面複製的 SHEET_ID 貼在這裡';
const SHEET_NAME = 'Applications';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);   // ← 與容器綁定版的唯一差別
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Submitted at', 'Name', 'Email', 'Role & company',
        'LinkedIn', 'Background', 'Why', 'Source'
      ]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.name       || '',
      data.email      || '',
      data.role       || '',
      data.linkedin   || '',
      data.background || '',
      data.why        || '',
      data.source     || ''
    ]);

    MailApp.sendEmail({
      to: 'vh@mvl.biz',
      subject: 'VC Academy — new application: ' + (data.name || 'unknown'),
      body: [
        'Name:       ' + (data.name || ''),
        'Email:      ' + (data.email || ''),
        'Role:       ' + (data.role || ''),
        'LinkedIn:   ' + (data.linkedin || ''),
        'Background: ' + (data.background || ''),
        'Source:     ' + (data.source || ''),
        '',
        'Why they want to join:',
        (data.why || '')
      ].join('\n')
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

**3. 部署方式與上面完全相同**（網頁應用程式 → 執行身分「我」→ 存取權「任何人」）。

> 第一次授權時會多要一項 Google Sheets 的存取權限，這是正常的 ——
> 因為獨立腳本需要被授權去開啟那份試算表。

## 常見踩雷點

**① 存取權沒設成「任何人」← 最常見**
部署時若選成「只有我」或「機構內的任何人」，你自己測試會成功（因為你已登入），
但訪客送出會全部失敗，而且網頁端看不出來（no-cors 讀不到錯誤）。
→ 必須是 **「任何人」/ "Anyone"**。

**② 複製到 `/dev` 網址**
部署畫面可能同時出現兩個網址。`/dev` 只有你自己能用，
**一定要用 `/exec` 結尾那個**。

**③ 授權時出現「Google 尚未驗證這個應用程式」**
這是正常的 —— 因為這支腳本是你自己寫的，沒有經過 Google 審查。
點 **「進階」→「前往〔專案名稱〕（不安全）」** 繼續即可。

**④ 改了程式碼卻沒生效**
Apps Script 改完要重新部署才會更新：
**部署 → 管理部署作業 → 右上鉛筆圖示 → 版本選「新版本」→ 部署**。
（直接按「新增部署作業」會產生新網址，那樣就要回來換 `APPLY_URL`。）

**⑤ 通知信沒收到**
第一次執行 `MailApp` 需要授權；另外免費 Google 帳號每天寄信有上限（約 100 封），
以每梯約 10 人的申請量不會碰到。

## 測試

部署並填好 `APPLY_URL` 後，在網站 Apply 頁送出一筆測試資料，確認：

1. Google Sheet 出現新的一列
2. 通知信有收到（若有保留寄信功能）
3. 網頁顯示 "Application received." 成功畫面

> ⚠️ 注意：瀏覽器基於安全限制無法讀取 Apps Script 的回應（no-cors），
> 所以網頁只要送出成功就會顯示成功畫面。**務必實際測一筆**，
> 確認資料真的有進到 Sheet。
