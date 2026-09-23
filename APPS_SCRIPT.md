# VC Academy 報名表單 — 後端設定

網站表單（Apply 頁）會把資料送到 Google Apps Script，由它寫進 Google Sheet。
每筆申請也會寄一封通知信給團隊，並自動回一封中英雙語的確認信給申請人。
做法與 SELL GLOBAL（training.mvl.biz）相同。

---

## 步驟 1 — 建立 Google Sheet

新建一份 Google Sheet，命名例如 `VC Academy Applications`。
**不需要**手動建立標題列，程式會自動建立。

## 步驟 2 — 貼上 Apps Script

在該 Sheet 中：**擴充功能 → Apps Script**，把預設內容全部刪除，貼上以下程式碼：

```javascript
/**
 * VC Academy 報名表單後端（Google Apps Script）
 *
 * 做三件事：
 *   ① 把申請寫進 Google Sheet
 *   ② 寄一封通知信給團隊
 *   ③ 自動回一封中英雙語的確認信給申請人（新增）
 *
 * 部署：見 APPS_SCRIPT.md。已經部署過的話，不要「新增部署作業」（會換網址），
 * 請用「部署 → 管理部署作業 → 右上鉛筆 → 版本選『新版本』→ 部署」，網址不變，網站不用改。
 */

// ===== 設定（要改的只有這裡）=====
const SHEET_NAME  = 'Applications';
const NOTIFY_TO   = 'vh@mvl.biz';   // 每筆申請的內部通知信收件者；多人用逗號分隔：'a@mvl.biz, b@mvl.biz'
const REPLY_TO    = 'vh@mvl.biz';   // 申請人在確認信按「回覆」，會寄到這裡
const SENDER_NAME = 'VC Academy';   // 確認信的寄件人名稱（寄件信箱是部署這支腳本的 Google 帳號）
const FIRST_SESSION_EN = 'Wednesday, November 11, 2026';
const FIRST_SESSION_ZH = '2026 年 11 月 11 日（週三）';
// VC Starter Kit 做好之後，把下載連結貼在下面的引號裡，重新部署，之後的確認信就會附上連結。
// 留空的話，信裡會寫「開課前會用 Email 寄給你」。
const STARTER_KIT_URL = '';

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

    // 兩封信都是「盡力而為」：寄信出錯，不能讓已經寫進 Sheet 的申請被當成失敗
    // （網站會顯示錯誤，訪客就會以為沒送成而重複送出）。錯誤會記在「執行作業」的紀錄裡。
    try { notifyTeam_(data); } catch (err) { console.error('notifyTeam_ failed: ' + err); }
    try { autoReply_(data);  } catch (err) { console.error('autoReply_ failed: ' + err); }

    return json_({ result: 'success' });

  } catch (err) {
    return json_({ result: 'error', message: err.message });
  }
}

// 給團隊的通知信（內容與原本相同）
function notifyTeam_(data) {
  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject: 'VC Academy — new application: ' + clean_(data.name, 80, 'unknown'),
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
}

// 給申請人的確認信（中英雙語：上英下中）
function autoReply_(data) {
  const email = String(data.email || '').trim();
  if (!isEmail_(email)) return;                        // 信箱格式不對就不寄
  if (MailApp.getRemainingDailyQuota() < 5) return;    // 每日寄信額度快用完時，留給團隊通知信

  const name = clean_(data.name, 80, '');
  const kitEn = STARTER_KIT_URL
    ? 'Your VC Starter Kit: ' + STARTER_KIT_URL
    : "We'll email you the VC Starter Kit before the first session.";
  const kitZh = STARTER_KIT_URL
    ? '你的 VC Starter Kit：' + STARTER_KIT_URL
    : '開課前，我們會用 Email 把 VC Starter Kit 寄給你。';

  const body = [
    (name ? 'Hi ' + name + ',' : 'Hi,'),
    '',
    "Thank you for applying to the VC Academy founding cohort. We've received your application.",
    '',
    'What happens next',
    "- Our team reviews every application and replies personally. If you're accepted, we'll send you payment details.",
    '- ' + kitEn,
    '- The first session is on ' + FIRST_SESSION_EN + ', in person in Taipei.',
    '',
    'Questions? Just reply to this email.',
    '',
    '— VC Academy · Mosaic Venture Lab',
    '',
    '――――――――――――――――',
    '',
    (name ? name + ' 你好，' : '你好，'),
    '',
    '謝謝你申請 VC Academy 創始班，我們已收到你的申請。',
    '',
    '接下來',
    '- 我們的團隊會逐一審閱每份申請，並親自回覆。若你獲得錄取，我們會寄付款資訊給你。',
    '- ' + kitZh,
    '- 首堂課是 ' + FIRST_SESSION_ZH + '，在台北實體授課。',
    '',
    '有任何問題，直接回覆這封信即可。',
    '',
    '— VC Academy · Mosaic Venture Lab'
  ].join('\n');

  MailApp.sendEmail({
    to: email,
    replyTo: REPLY_TO,
    name: SENDER_NAME,
    subject: 'VC Academy — we received your application · 已收到你的申請',
    body: body
  });
}

// 手動測試：在編輯器上方的函式選單選 testAutoReply → 執行。
// 會寄一封範例確認信到 NOTIFY_TO 的第一個信箱（第一次執行會要求授權寄信）。
function testAutoReply() {
  autoReply_({ name: 'Test Applicant', email: NOTIFY_TO.split(',')[0].trim() });
}

// ===== 小工具 =====
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// 去掉換行（避免有人在姓名欄塞換行去影響信件標題／內容）、限制長度；'—' 是網站對空欄位的預設值
function clean_(s, max, fallback) {
  const t = String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  return (t && t !== '—') ? t : (fallback || '');
}
// 只接受「一個」正常的信箱（不含逗號、分號、空白），避免一次寄給多人
function isEmail_(s) {
  return /^[^\s@,;<>()]+@[^\s@,;<>()]+\.[^\s@,;<>()]{2,}$/.test(s);
}
```

> 要改的只有程式最上面的「設定」區：通知信收件者（`NOTIFY_TO`，多人用逗號分隔，例如 `'vh@mvl.biz, kh@mvl.biz, jy@mvl.biz'`）、
> 申請人回信的信箱（`REPLY_TO`）、首堂課日期，以及 Starter Kit 連結（`STARTER_KIT_URL`）。

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

## 自動回信（申請人收到的確認信）

每筆申請送出後，除了通知團隊，申請人也會立刻收到一封中英雙語的確認信（上英下中）：我們已收到申請、團隊會逐一審閱並親自回覆、錄取後會寄付款資訊、VC Starter Kit 開課前寄出（或附上連結）、首堂課日期。申請人按「回覆」會寄到 `REPLY_TO`。

- **Starter Kit 做好之後**：把下載連結貼進 `STARTER_KIT_URL`，重新部署（步驟見下方「改了程式碼卻沒生效」），之後的確認信就會附上連結。
- **寄件人**是部署這支腳本的 Google 帳號。建議用 @mvl.biz 帳號部署，比較像正式信件，也比較不會被歸到垃圾信匣。
- **先自己測一次**：在 Apps Script 編輯器上方的函式選單選 `testAutoReply` → 執行，會寄一封範例確認信到 `NOTIFY_TO` 的第一個信箱。第一次執行會要求授權寄信。
- 兩封信都是「盡力而為」：寄信出錯（例如當天額度用盡）不會影響申請寫進 Sheet，網站也不會顯示錯誤讓訪客重複送出；錯誤會記在「執行作業」的紀錄裡。
- 首堂課日期改了，要同時改 `FIRST_SESSION_EN` 與 `FIRST_SESSION_ZH`。

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

**2. 到 [script.google.com/create](https://script.google.com/create) 建立新專案**，把上面「步驟 2」的完整程式貼上，再改兩處：

1. 在最上面加一行：`const SHEET_ID = '把上面複製的 SHEET_ID 貼在這裡';`
2. 把 `SpreadsheetApp.getActiveSpreadsheet()` 改成 `SpreadsheetApp.openById(SHEET_ID)`（這是與容器綁定版的唯一差別）

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

**⑤ 通知信或確認信沒收到**
第一次執行 `MailApp` 需要授權；另外免費 Google 帳號每天寄信有上限（約 100 封），
每筆申請會寄 2 封（團隊通知＋申請人確認信），以每梯約 10 人的申請量不會碰到。確認信寄不出去不會影響申請寫進 Sheet；也請提醒申請人查一下垃圾信匣。

## 測試

部署並填好 `APPLY_URL` 後，在網站 Apply 頁送出一筆測試資料，確認：

1. Google Sheet 出現新的一列
2. 通知信有收到（若有保留寄信功能）
3. 網頁顯示 "Application received." 成功畫面
4. 申請人填的信箱收到中英雙語確認信（也檢查垃圾信匣）

> ⚠️ 注意：瀏覽器基於安全限制無法讀取 Apps Script 的回應（no-cors），
> 所以網頁只要送出成功就會顯示成功畫面。**務必實際測一筆**，
> 確認資料真的有進到 Sheet。
