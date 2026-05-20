/**
 * Web App 入口 — 用 ?action= 路由
 *   home      預設首頁（LINE 登入按鈕）
 *   login     轉跳到 LINE OAuth 授權頁
 *   callback  LINE 授權後的 callback handler
 *   bind      綁定連結著陸頁（帶 ?token=xxx）
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  try {
    // LINE OAuth callback：只要帶 code+state 就視為 callback（與 action 無關）
    if (params.code && params.state) return pageCallback_(e);
    if (params.error) return pageError_('LINE 授權失敗：' + (params.error_description || params.error));

    const action = params.action || 'home';
    if (action === 'home')     return pageHome_();
    if (action === 'bind')     return pageBindLanding_(e);
    if (action === 'schedule') return pageSchedule_(e);
    return pageError_('未知頁面：' + action);
  } catch (err) {
    return pageError_('系統錯誤：' + (err && err.message ? err.message : err));
  }
}

function pageHome_() {
  const tpl = HtmlService.createTemplateFromFile('home');
  tpl.loginUrl = buildLineOAuthUrl_(null);
  return tpl.evaluate()
    .setTitle('彩券行員工系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** 建構 LINE OAuth 授權 URL — 給 home 和 bind 頁直接連到 LINE，繞過 Apps Script iframe */
function buildLineOAuthUrl_(bindToken) {
  const cfg = getLineLoginCfg_();
  const nonce = Utilities.getUuid();
  const stateObj = { nonce: nonce };
  if (bindToken) stateObj.bind_token = bindToken;
  const state = Utilities.base64EncodeWebSafe(JSON.stringify(stateObj));
  CacheService.getScriptCache().put('oauth_' + nonce, '1', 600);

  return 'https://access.line.me/oauth2/v2.1/authorize' +
    '?response_type=code' +
    '&client_id=' + encodeURIComponent(cfg.channelId) +
    '&redirect_uri=' + encodeURIComponent(getCallbackUrl_()) +
    '&state=' + encodeURIComponent(state) +
    '&scope=' + encodeURIComponent('profile openid');
}

function pageCallback_(e) {
  if (e.parameter.error) {
    return pageError_('LINE 授權失敗：' + (e.parameter.error_description || e.parameter.error));
  }
  if (!e.parameter.code || !e.parameter.state) {
    return pageError_('缺少必要參數');
  }

  let stateObj;
  try {
    stateObj = JSON.parse(
      Utilities.newBlob(Utilities.base64DecodeWebSafe(e.parameter.state)).getDataAsString()
    );
  } catch (err) {
    return pageError_('state 格式錯誤');
  }
  if (!CacheService.getScriptCache().get('oauth_' + stateObj.nonce)) {
    return pageError_('登入工作階段已過期，請重新開始');
  }
  CacheService.getScriptCache().remove('oauth_' + stateObj.nonce);

  const token = lineExchangeCode_(e.parameter.code);
  const profile = lineFetchProfile_(token.access_token);

  if (stateObj.bind_token) {
    return doBindWithToken_(stateObj.bind_token, profile);
  }
  return doLogin_(profile);
}

function pageSchedule_(e) {
  // 身分驗證：要求 LINE 登入後的 session
  const emp = getEmployeeFromSession_(e);
  if (!emp) {
    // Apps Script 沙盒禁用程式自動跳轉，改顯示登入按鈕讓使用者主動點
    return pageHtml_(
      '需要登入',
      '<h2>需要登入</h2>' +
      '<p>請先用 LINE 登入後再進入排班系統。</p>' +
      '<p><a href="' + getWebAppUrl_() + '" target="_top" style="display:inline-block;background:#06c755;color:#fff;padding:12px 20px;border-radius:6px;text-decoration:none;font-weight:600">前往登入頁</a></p>'
    );
  }
  return renderSchedulePage_(e.parameter.sid || '', emp, e);
}

function pageBindLanding_(e) {
  const token = e.parameter.token;
  if (!token) return pageError_('缺少綁定 token');
  const info = lookupBindToken_(token);
  if (!info) return pageError_('綁定連結無效');
  if (info.used_at) return pageError_('此綁定連結已使用過，請聯絡管理者重發');
  if (new Date() > new Date(info.expires_at)) return pageError_('此綁定連結已過期，請聯絡管理者重發');

  const tpl = HtmlService.createTemplateFromFile('bind');
  tpl.empId = info.employee_id;
  tpl.empName = info.employee_name;
  tpl.expiresAt = Utilities.formatDate(new Date(info.expires_at), TIMEZONE, 'yyyy/MM/dd HH:mm');
  // 直接連到 LINE OAuth URL（含 bind_token 在 state 內）— 繞過 Apps Script iframe 跨域問題
  tpl.confirmUrl = buildLineOAuthUrl_(token);
  return tpl.evaluate()
    .setTitle('員工身分綁定')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function doLogin_(profile) {
  const emp = findEmployeeByLineId_(profile.userId);
  if (!emp) {
    return pageHtml_(
      '尚未綁定',
      '<h2>尚未綁定員工身分</h2>' +
      '<p>哈囉 ' + escapeHtml_(profile.displayName) + '，您的 LINE 尚未綁定到任何員工身分。</p>' +
      '<p>請聯絡管理者取得綁定連結。</p>' +
      '<p style="color:#888;font-size:11px;word-break:break-all">LINE userId: ' + escapeHtml_(profile.userId) + '</p>'
    );
  }
  // 建立 session，直接渲染排班頁（不重定向 — Apps Script iframe 禁用 programmatic top-level navigation）
  const sid = createSession_(profile.userId);
  return renderSchedulePage_(sid, emp, null);
}

function doBindWithToken_(bindToken, profile) {
  const info = lookupBindToken_(bindToken);
  if (!info) return pageError_('綁定連結無效');
  if (info.used_at) return pageError_('此綁定連結已使用過');
  if (new Date() > new Date(info.expires_at)) return pageError_('此綁定連結已過期');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  const data = empSheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(COL.EMPLOYEES.ID);
  const lineCol = headers.indexOf(COL.EMPLOYEES.LINE_USER_ID);
  const updatedCol = headers.indexOf(COL.EMPLOYEES.UPDATED_AT);

  for (let i = 1; i < data.length; i++) {
    if (data[i][lineCol] === profile.userId && data[i][idCol] !== info.employee_id) {
      return pageError_('您的 LINE 已綁定到另一位員工：' + data[i][idCol]);
    }
  }

  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === info.employee_id) {
      if (data[i][lineCol] && data[i][lineCol] !== profile.userId) {
        return pageError_('此員工已綁定不同的 LINE 帳號，請聯絡管理者');
      }
      empSheet.getRange(i + 1, lineCol + 1).setValue(profile.userId);
      empSheet.getRange(i + 1, updatedCol + 1).setValue(new Date());
      updated = true;
      break;
    }
  }
  if (!updated) return pageError_('找不到員工：' + info.employee_id);

  markBindTokenUsed_(bindToken, profile.userId);
  // 綁定成功後直接渲染排班頁
  const sid = createSession_(profile.userId);
  // 重新從 Sheet 取資料（剛綁完 line_user_id 才有值）
  const empFresh = findEmployeeByLineId_(profile.userId);
  return renderSchedulePage_(sid, empFresh || { [COL.EMPLOYEES.ID]: info.employee_id, [COL.EMPLOYEES.NAME]: info.employee_name, [COL.EMPLOYEES.ROLE]: 'staff', [COL.EMPLOYEES.HOME_STORE]: '' }, null);
}

function lineExchangeCode_(code) {
  const cfg = getLineLoginCfg_();
  const resp = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/token', {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      code: code,
      redirect_uri: getCallbackUrl_(),
      client_id: cfg.channelId,
      client_secret: cfg.channelSecret
    },
    muteHttpExceptions: true
  });
  const body = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200) {
    throw new Error('LINE token 交換失敗：' + JSON.stringify(body));
  }
  return body;
}

function lineFetchProfile_(accessToken) {
  const resp = UrlFetchApp.fetch('https://api.line.me/v2/profile', {
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('LINE profile 取得失敗：' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

function findEmployeeByLineId_(lineUserId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const lineCol = headers.indexOf(COL.EMPLOYEES.LINE_USER_ID);
  for (let i = 1; i < data.length; i++) {
    if (data[i][lineCol] === lineUserId) {
      const obj = {};
      headers.forEach(function (h, j) { obj[h] = data[i][j]; });
      return obj;
    }
  }
  return null;
}

function getLineLoginCfg_() {
  const id = getSetting_('line_login_channel_id');
  const secret = getSetting_('line_login_channel_secret');
  if (!id || !secret) {
    throw new Error('LINE Login channel ID / secret 未設定，請至 Settings 分頁填入');
  }
  return { channelId: String(id), channelSecret: String(secret) };
}

/** 建立登入 session（1 小時 TTL），回傳 session id */
function createSession_(lineUserId) {
  const sid = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('sess_' + sid, lineUserId, 3600);
  return sid;
}

/** 由 URL ?sid=xxx 解析回員工資料 */
function getEmployeeFromSession_(e) {
  const sid = e && e.parameter && e.parameter.sid;
  if (!sid) return null;
  const lineUserId = CacheService.getScriptCache().get('sess_' + sid);
  if (!lineUserId) return null;
  return findEmployeeByLineId_(lineUserId);
}

/** 渲染排班頁 — 由 pageSchedule_ 與 doLogin_/doBindWithToken_ 共用 */
function renderSchedulePage_(sid, emp, e) {
  const now = new Date();
  const year = parseInt((e && e.parameter && e.parameter.year), 10) || now.getFullYear();
  const month = parseInt((e && e.parameter && e.parameter.month), 10) || (now.getMonth() + 1);

  const tpl = HtmlService.createTemplateFromFile('schedule');
  tpl.data = getScheduleViewData(year, month);
  tpl.webAppUrl = getWebAppUrl_();
  tpl.sid = sid || '';
  tpl.currentUser = {
    id: emp[COL.EMPLOYEES.ID],
    name: emp[COL.EMPLOYEES.NAME],
    role: emp[COL.EMPLOYEES.ROLE],
    home_store: emp[COL.EMPLOYEES.HOME_STORE]
  };
  return tpl.evaluate()
    .setTitle('排班總覽 · ' + year + '/' + month)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getWebAppUrl_() {
  // 優先：當前執行環境的真實 deployment URL（Web App context 會回 /exec）
  // 若是從編輯器執行（會回 /dev），fallback 到 Settings 內手動設定的 /exec URL
  try {
    const live = ScriptApp.getService().getUrl();
    if (live && live.indexOf('/exec') > 0) return live;
  } catch (e) {}
  const stored = getSetting_('web_app_url');
  if (stored) return String(stored).trim();
  return '';
}

function getCallbackUrl_() {
  // LINE Callback URL 設定為 bare /exec，code/state 帶回時由 doGet 自動偵測
  return getWebAppUrl_();
}

function pageError_(msg) {
  return pageHtml_('錯誤', '<h2>⚠️ 錯誤</h2><p>' + escapeHtml_(msg) + '</p>');
}

function pageHtml_(title, bodyHtml) {
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + escapeHtml_(title) + '</title>' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;max-width:480px;margin:40px auto;padding:0 20px;color:#333;line-height:1.6}' +
    'h2{color:#06c755}' +
    'ul{padding-left:20px}' +
    '</style>' +
    '</head><body>' + bodyHtml + '</body></html>'
  ).setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function escapeHtml_(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/** 偵錯：印出當前 Web App URL，比對部署 URL 是否一致 */
function DebugUrls() {
  const url = ScriptApp.getService().getUrl();
  const html =
    '<div style="font-family:monospace;font-size:12px;line-height:1.6">' +
    '<p><b>ScriptApp.getService().getUrl() 回傳：</b></p>' +
    '<textarea readonly style="width:100%;height:70px;font-size:11px" onclick="this.select()">' + (url || '(null)') + '</textarea>' +
    '<p style="margin-top:16px">請對照三處 URL 是否完全一致：</p>' +
    '<ol>' +
    '<li>上面這個（程式碼用的）</li>' +
    '<li>「管理部署」對話框中顯示的網址</li>' +
    '<li>LINE Developers Console 的 Callback URL 欄位</li>' +
    '</ol>' +
    '<p>有任何一個不一樣 → OAuth 就會壞。</p>' +
    '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(720).setHeight(360),
    'URL 偵錯'
  );
}

/** 偵錯：測試 UrlFetchApp 權限有沒有給對 */
function TestUrlFetch() {
  try {
    const r = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=dummy', {
      muteHttpExceptions: true
    });
    SpreadsheetApp.getUi().alert(
      'UrlFetchApp 可用 ✓\n\n' +
      '回應碼: ' + r.getResponseCode() + '（400 = LINE 回應正常，代表網路通）\n\n' +
      '回應內容: ' + r.getContentText().substring(0, 200)
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'UrlFetchApp 失敗:\n\n' + e.message + '\n\n' +
      '若是權限問題，請執行此函式時授權新的 scope，再試一次綁定。'
    );
  }
}

/** 給「呼叫此 Web App URL」用 — 部署後在 Apps Script 編輯器執行可印出 URL */
function ShowWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const callback = url + '?action=callback';
  SpreadsheetApp.getUi().alert(
    'Web App URL:\n' + url + '\n\n' +
    'LINE Login Callback URL（貼回 LINE Developers Console）:\n' + callback
  );
}
