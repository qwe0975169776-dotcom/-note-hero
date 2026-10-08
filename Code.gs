/**
 * 音符勇者 — Google 試算表存檔程式
 * 用法：在試算表選「擴充功能 → Apps Script」，把這整段貼上後部署成「網頁應用程式」。
 * 第一次執行時會自動建立 Users、成績、錯題統計、排行榜 四個分頁。
 */

// ====== 只有第一次手動貼上時才需要填（之後自動更新的版本這裡是空的，數值已存在「指令碼屬性」） ======
const SETUP_SHEET_ID = '';      // 試算表網址中 /d/ 和 /edit 之間那串字
const SETUP_TEACHER_PIN = '';   // 老師密碼（要和遊戲老師面板一樣）

// 自動更新：從 GitHub 抓最新的 Code.gs，換掉自己並部署新版本
const GITHUB_RAW = 'https://raw.githubusercontent.com/qwe0975169776-dotcom/-note-hero/main/Code.gs';

// 試算表 ID 與老師密碼存在「專案設定 → 指令碼屬性」，不寫在程式碼裡（GitHub 是公開的）
function props_() { return PropertiesService.getScriptProperties(); }
function saveSetup_() {
  const p = props_();
  if (SETUP_SHEET_ID && p.getProperty('SHEET_ID') !== SETUP_SHEET_ID) p.setProperty('SHEET_ID', SETUP_SHEET_ID);
  if (SETUP_TEACHER_PIN && !p.getProperty('TEACHER_PIN')) p.setProperty('TEACHER_PIN', SETUP_TEACHER_PIN);
}
function pin_() { return String(props_().getProperty('TEACHER_PIN') || '1234'); }
function sheetId_() { return props_().getProperty('SHEET_ID') || SETUP_SHEET_ID; }

function ss_() {
  const id = sheetId_();
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表：請在「專案設定 → 指令碼屬性」新增 SHEET_ID');
  return ss;
}

/** 第一次貼上後，在編輯器選這個函式按「執行」，同意授權 */
function 一次設定() {
  saveSetup_();
  setup();
  Logger.log('完成！試算表：' + ss_().getName() + '；腳本 ID：' + ScriptApp.getScriptId());
}

const NOTE_KEYS  = ['do', 're', 'mi', 'fa', 'sol', 'la', 'si', 'hdo'];
const NOTE_NAMES = { do: 'ㄉㄡ', re: 'ㄖㄨㄟ', mi: 'ㄇㄧ', fa: 'ㄈㄚ', sol: 'ㄙㄛ', la: 'ㄌㄚ', si: 'ㄒㄧ', hdo: '高音ㄉㄡ' };

// ---------- 分頁 ----------
function sheet_(name, headers) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.setFrozenRows(1);
  }
  // 表頭和程式不一樣時（例如更新版本後多了欄位）自動補上
  const head = s.getRange(1, 1, 1, headers.length);
  if (head.getValues()[0].join('|') !== headers.join('|')) {
    head.setValues([headers]).setFontWeight('bold').setBackground('#fff4cf');
  }
  return s;
}
function usersSheet_()  { return sheet_('Users',    ['姓名', '目前積分', '累計獲得', '獎章', '最後更新', '存檔', '段位', '段位積分', '已過關卡', '最常錯的音']); }
function scoresSheet_() { return sheet_('成績',     ['時間', '姓名', '模式', '關卡', '答對', '答錯', '正確率', '最高連擊', '星數', '獲得金幣', '答錯的音']); }
function feedSheet_()   { return sheet_('廣播',     ['時間', '姓名', '內容', 'ts']); }
function raidSheet_()   { return sheet_('世界魔王', ['狀態', '名稱', '地區', '總血量', '剩餘血量', '開始ts', '最後一擊', '結束時間']); }
function raidDmgSheet_(){ return sheet_('世界魔王貢獻', ['開始ts', '姓名', '傷害']); }
function configSheet_() { return sheet_('設定', ['項目', '值']); }
function boardSheet_()  { return sheet_('排行榜',   ['名次', '姓名', '段位', '段位積分', '已過關卡', '累計獲得']); }
function tuneSheet_()   {
  const s = sheet_('修改器', ['代號', '類別', '名稱', '項目', '值（空白＝預設）', '預設值', '說明']);
  s.getRange('E:F').setNumberFormat('@');   // 純文字，避免 3,4 被變成數字
  return s;
}
function logSheet_()    { return sheet_('修改紀錄', ['時間', '對象', '內容']); }
function reportSheet_() { return sheet_('學習報告', ['姓名', '總答題數', '正確率', '最近7天答題', '最近7天正確率', '進步情形', '最常錯的音', '最常搞混', '建議練習', '完成的練習課', '最後更新']); }
function notesSheet_()  {
  const s = sheet_('錯題統計', ['音', '出題次數', '答錯次數', '錯誤率']);
  if (s.getLastRow() < 2) {
    s.getRange(2, 1, NOTE_KEYS.length, 4).setValues(NOTE_KEYS.map(k => [NOTE_NAMES[k], 0, 0, 0]));
    s.getRange(2, 4, NOTE_KEYS.length, 1).setNumberFormat('0%');
  }
  return s;
}

/** 在編輯器裡手動執行一次，可以先把分頁都建好 */
function setup() { usersSheet_(); scoresSheet_(); notesSheet_(); boardSheet_(); feedSheet_(); raidSheet_(); raidDmgSheet_(); configSheet_(); tuneSheet_(); logSheet_(); reportSheet_(); }

/** 檢查用：在編輯器執行，Users 分頁應該會出現「測試小朋友」 */
function testSave() {
  saveUser_('測試小朋友', { coins: 1, total: 1, ts: Date.now(), tierName: '測試', rankScore: 0, progress: '0/25 關', weak: '' });
  rebuildBoard_();
}

const CODE_VERSION = 10;   // 檢查用：網址回傳的版本號

// ---------- 網頁應用程式入口 ----------
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    saveSetup_();
    if (p.action === 'ping') {
      // 真的打開試算表檢查一次，設定有問題時「測試連線」就會顯示錯誤
      usersSheet_(); scoresSheet_(); notesSheet_(); boardSheet_(); feedSheet_(); tuneSheet_(); logSheet_(); reportSheet_();
      return out_({ ok: true, version: CODE_VERSION, sheet: ss_().getName() });
    }
    if (p.action === 'testwrite') {
      // 檢查用：從網址直接寫一筆「網址測試」到 Users
      saveUser_('網址測試', { coins: 2, total: 2, ts: Date.now(), tierName: '測試', rankScore: 0, progress: '0/25 關', weak: '' });
      rebuildBoard_();
      return out_({ ok: true, version: CODE_VERSION, wrote: '網址測試' });
    }
    if (p.action === 'load')  return out_({ ok: true, user: loadUser_(p.name) });
    if (p.action === 'class') return out_({ ok: true, stats: noteStats_(), board: board_() });
    if (p.action === 'raid')  return out_({ ok: true, raid: raid_() });
    if (p.action === 'config') return out_({ ok: true, config: config_(), tune: tune_() });
    if (p.action === 'feed')  return out_({ ok: true, events: feed_(Number(p.since) || 0) });
    return out_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    saveSetup_();
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'save') {
      const res = saveUser_(d.name, d.save) || {};
      if (d.record) addRecord_(d.record);
      if (d.notes)  addNotes_(d.notes);
      rebuildBoard_();
      return out_({ ok: true, ops: res.ops || [], tSeen: res.tSeen || 0 });
    }
    // ---------- 修改器（都要老師密碼） ----------
    if (['tuneSync', 'tuneSet', 'users', 'userEdit'].includes(d.action) && String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
    if (d.action === 'tuneSync') { tuneSync_(d.rows); return out_({ ok: true, tune: tune_() }); }
    if (d.action === 'tuneSet')  { tuneSync_(d.rows); tuneSet_(d.changes || {}); return out_({ ok: true, tune: tune_() }); }
    if (d.action === 'users')    return out_({ ok: true, users: users_() });
    if (d.action === 'userEdit') { const r = userEdit_(d.name, d.op, d.label); rebuildBoard_(); return out_(Object.assign({ ok: true }, r)); }
    if (d.action === 'record') {
      if (d.record) addRecord_(d.record);
      if (d.notes)  addNotes_(d.notes);
      return out_({ ok: true });
    }
    if (d.action === 'event') {
      addEvent_(d.name, d.text, d.ts);
      return out_({ ok: true });
    }
    if (d.action === 'setConfig') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
      setConfig_(d.key, d.value);
      return out_({ ok: true, config: config_() });
    }
    if (d.action === 'raidHit')   return out_({ ok: true, raid: raidHit_(d) });
    if (d.action === 'raidStart') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
      return out_({ ok: true, raid: raidStart_(d) });
    }
    if (d.action === 'raidStop') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
      return out_({ ok: true, raid: raidStop_() });
    }
    if (d.action === 'setPin') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '舊的老師密碼不對' });
      const np = String(d.newPin || '').trim();
      if (!np) return out_({ ok: false, error: '新密碼不能空白' });
      props_().setProperty('TEACHER_PIN', np);
      return out_({ ok: true });
    }
    if (d.action === 'selfUpdate') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
      return out_(Object.assign({ ok: true }, selfUpdate_(d.deploymentId)));
    }
    if (d.action === 'resetCoins') {
      if (String(d.pin) !== pin_()) return out_({ ok: false, error: '老師密碼不對' });
      resetCoins_();
      rebuildBoard_();
      return out_({ ok: true });
    }
    return out_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- 玩家 ----------
function findRow_(s, name) {
  const n = s.getLastRow();
  if (n < 2) return -1;
  const names = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) if (String(names[i][0]).trim() === name) return i + 2;
  return -1;
}

function loadUser_(name) {
  name = String(name || '').trim();
  if (!name) return null;
  const s = usersSheet_();
  const r = findRow_(s, name);
  if (r < 0) return null;
  return storedSave_(s, r);
}

// 讀出存檔；如果老師直接在試算表改了「目前積分」那一格，就記成一次調整
function storedSave_(s, r) {
  const row = s.getRange(r, 1, 1, 6).getValues()[0];
  let sv;
  try { sv = JSON.parse(row[5]); } catch (e) { return null; }
  if (!sv) return null;
  const cell = row[1];
  if (cell !== '' && cell !== null && isFinite(Number(cell)) && Number(cell) !== Number(sv.coins || 0)) {
    addOp_(sv, { op: 'setCoins', v: Math.max(0, Math.round(Number(cell))) });
    applyOp_(sv, sv.tOps[sv.tOps.length - 1]);
    sv.tSeen = sv.tOps[sv.tOps.length - 1].id;
    sv.ts = Date.now();
    s.getRange(r, 6).setValue(JSON.stringify(sv));
    log_(row[0], '在試算表直接把金幣改成 ' + sv.coins);
  }
  return sv;
}

function saveUser_(name, save) {
  name = String(name || '').trim();
  if (!name || !save) return;
  const s = usersSheet_();
  let r = findRow_(s, name);
  if (r < 0) { s.appendRow([name]); r = s.getLastRow(); }
  // 老師做過的調整：這次上傳的存檔還沒包含的，就補套上去，並回傳給遊戲
  const old = storedSave_(s, r);
  const ops = (old && old.tOps) || [];
  const seen = Number(save.tSeen || 0);
  const miss = ops.filter(o => o.id > seen);
  miss.forEach(o => applyOp_(save, o));
  save.tOps = ops;
  if (ops.length) save.tSeen = Math.max(seen, ops[ops.length - 1].id);
  const badges = (save.badgeNames || []).join('、');
  const extra = [save.tierName || '', save.rankScore || 0, save.progress || '', save.weak || ''];
  const report = save.report;
  ['badgeNames', 'tierName', 'rankScore', 'progress', 'weak', 'report'].forEach(k => delete save[k]);
  if (report) saveReport_(name, report);
  s.getRange(r, 1, 1, 10).setValues([[name, save.coins || 0, save.total || 0, badges, new Date(), JSON.stringify(save)].concat(extra)]);
  return { ops: miss, tSeen: save.tSeen || 0 };
}

function resetCoins_() {
  const s = usersSheet_();
  const n = s.getLastRow();
  if (n < 2) return;
  const rows = s.getRange(2, 1, n - 1, 6).getValues();
  rows.forEach(row => {
    row[1] = 0;
    try { const sv = JSON.parse(row[5]); sv.coins = 0; row[5] = JSON.stringify(sv); } catch (e) {}
  });
  s.getRange(2, 1, n - 1, 6).setValues(rows);
}

// ---------- 成績與錯題 ----------
function addRecord_(r) {
  const total = (r.right || 0) + (r.wrong || 0);
  scoresSheet_().appendRow([
    new Date(), r.name || '', r.mode || '單人', r.level || '',
    r.right || 0, r.wrong || 0, total ? Math.round((r.right || 0) / total * 100) + '%' : '',
    r.best || 0, r.stars || '', r.coins || 0, r.missed || ''
  ]);
}

function addNotes_(notes) {
  const s = notesSheet_();
  const rng = s.getRange(2, 1, NOTE_KEYS.length, 4);
  const v = rng.getValues();
  NOTE_KEYS.forEach((k, i) => {
    const d = notes[k];
    if (!d) return;
    v[i][1] = Number(v[i][1] || 0) + Number(d.asked || 0);
    v[i][2] = Number(v[i][2] || 0) + Number(d.wrong || 0);
    v[i][3] = v[i][1] ? v[i][2] / v[i][1] : 0;
  });
  rng.setValues(v);
}

function noteStats_() {
  const v = notesSheet_().getRange(2, 1, NOTE_KEYS.length, 3).getValues();
  const out = {};
  NOTE_KEYS.forEach((k, i) => { out[k] = { asked: Number(v[i][1] || 0), wrong: Number(v[i][2] || 0) }; });
  return out;
}

// ---------- 全班設定（關閉冒險地圖、鎖商店） ----------
const CONFIG_KEYS = ['lockAdventure', 'lockShop', 'bigKeys', 'earQ'];
function config_() {
  const s = configSheet_(), n = s.getLastRow(), out = { lockAdventure: false, lockShop: false, bigKeys: false, earQ: false };
  if (n >= 2) s.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (CONFIG_KEYS.includes(r[0])) out[r[0]] = r[1] === true || r[1] === 'TRUE' || r[1] === 'true'; });
  return out;
}
function setConfig_(key, value) {
  if (!CONFIG_KEYS.includes(key)) return;
  const s = configSheet_(), n = s.getLastRow();
  if (n >= 2) { const v = s.getRange(2, 1, n - 1, 1).getValues(); for (let i = 0; i < v.length; i++) if (v[i][0] === key) { s.getRange(i + 2, 2).setValue(!!value); return; } }
  s.appendRow([key, !!value]);
}

// ---------- 世界魔王（全班合作） ----------
function raidRow_() {
  const s = raidSheet_();
  if (s.getLastRow() < 2) return null;
  return s.getRange(2, 1, 1, 8).getValues()[0];
}
function raid_() {
  const r = raidRow_();
  if (!r) return { active: false };
  const id = Number(r[5]) || 0;
  return { active: r[0] === '進行中', status: r[0], name: r[1], region: Number(r[2]) || 0, max: Number(r[3]) || 0, hp: Number(r[4]) || 0, id, last: r[6] || '', top: raidTop_(id) };
}
function raidTop_(id) {
  const s = raidDmgSheet_(), n = s.getLastRow();
  if (n < 2 || !id) return [];
  return s.getRange(2, 1, n - 1, 3).getValues().filter(r => Number(r[0]) === id)
    .sort((a, b) => Number(b[2]) - Number(a[2])).slice(0, 5).map(r => ({ name: r[1], dmg: Number(r[2]) }));
}
function raidStart_(d) {
  const hp = Math.max(10, Number(d.hp) || 300), name = String(d.name || '世界魔王');
  raidSheet_().getRange(2, 1, 1, 8).setValues([['進行中', name, Number(d.region) || 0, hp, hp, Date.now(), '', '']]);
  addEvent_('老師', `召喚了「${name}」！全班一起來打！`, Date.now());
  return raid_();
}
function raidStop_() {
  const r = raidRow_();
  if (r && r[0] === '進行中') raidSheet_().getRange(2, 1, 1, 8).setValues([['已結束', r[1], r[2], r[3], r[4], r[5], r[6], new Date()]]);
  return raid_();
}
function raidHit_(d) {
  const r = raidRow_();
  if (!r || r[0] !== '進行中' || Number(r[5]) !== Number(d.id)) return raid_();
  const dmg = Math.max(0, Math.min(50, Number(d.dmg) || 0)), name = String(d.name || '訪客');
  const hp = Math.max(0, Number(r[4]) - dmg);
  const s = raidDmgSheet_(), n = s.getLastRow();
  let row = -1;
  if (n >= 2) { const v = s.getRange(2, 1, n - 1, 2).getValues(); for (let i = 0; i < v.length; i++) if (Number(v[i][0]) === Number(r[5]) && v[i][1] === name) { row = i + 2; break; } }
  if (row < 0) s.appendRow([Number(r[5]), name, dmg]); else s.getRange(row, 3).setValue(Number(s.getRange(row, 3).getValue()) + dmg);
  if (hp === 0) {
    raidSheet_().getRange(2, 1, 1, 8).setValues([['已打倒', r[1], r[2], r[3], 0, r[5], name, new Date()]]);
    addEvent_(name, `給了「${r[1]}」最後一擊！全班一起打倒了魔王！`, Date.now());
  } else {
    raidSheet_().getRange(2, 5).setValue(hp);
  }
  return raid_();
}

// ---------- 全頻廣播 ----------
function addEvent_(name, text, ts) {
  if (!name || !text) return;
  const s = feedSheet_();
  s.appendRow([new Date(), String(name), String(text), Number(ts) || Date.now()]);
  if (s.getLastRow() > 600) s.deleteRows(2, 200);   // 只保留最近的廣播
}
function feed_(since) {
  const s = feedSheet_();
  const n = s.getLastRow();
  if (n < 2) return [];
  const start = Math.max(2, n - 59);
  return s.getRange(start, 1, n - start + 1, 4).getValues()
    .filter(r => Number(r[3]) > since)
    .map(r => ({ name: r[1], text: r[2], ts: Number(r[3]) }));
}

// ---------- 排行榜（依段位積分排名） ----------
function rebuildBoard_() {
  const s = usersSheet_();
  const n = s.getLastRow();
  const b = boardSheet_();
  if (b.getLastRow() > 1) b.getRange(2, 1, b.getLastRow() - 1, 6).clearContent();
  if (n < 2) return;
  const rows = s.getRange(2, 1, n - 1, 10).getValues()
    .filter(r => String(r[0]).trim())
    .sort((a, c) => (Number(c[7] || 0) - Number(a[7] || 0)) || (Number(c[2] || 0) - Number(a[2] || 0)))
    .map((r, i) => [i + 1, r[0], r[6], r[7] || 0, r[8], r[2] || 0]);
  if (rows.length) b.getRange(2, 1, rows.length, 6).setValues(rows);
}

function board_() {
  const b = boardSheet_();
  const n = b.getLastRow();
  if (n < 2) return [];
  return b.getRange(2, 1, Math.min(n - 1, 60), 6).getValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({ rank: r[0], name: r[1], tier: r[2], score: r[3], progress: r[4], total: r[5] }));
}

// ---------- 遊戲修改器 ----------
// 「修改器」分頁：每一列是一個可以調整的數值。「值」那一欄空白＝用預設值。
function tune_() {
  const s = tuneSheet_(), n = s.getLastRow(), out = {};
  if (n < 2) return out;
  s.getRange(2, 1, n - 1, 5).getValues().forEach(r => {
    const k = String(r[0]).trim(), v = r[4];
    if (k && v !== '' && v !== null) out[k] = String(v);
  });
  return out;
}
// 遊戲把全部項目（名稱、預設值、說明）送過來，缺的就補上，已經填的值不動
function tuneSync_(rows) {
  if (!Array.isArray(rows) || !rows.length) return;
  const s = tuneSheet_(), n = s.getLastRow();
  const cur = n >= 2 ? s.getRange(2, 1, n - 1, 7).getValues() : [];
  const idx = {};
  cur.forEach((r, i) => { idx[String(r[0]).trim()] = i; });
  rows.forEach(x => {
    const key = String(x[0]);
    if (key in idx) { const r = cur[idx[key]]; r[1] = x[1]; r[2] = x[2]; r[3] = x[3]; r[5] = String(x[4]); r[6] = x[5]; }
    else { idx[key] = cur.length; cur.push([key, x[1], x[2], x[3], '', String(x[4]), x[5]]); }
  });
  if (cur.length) s.getRange(2, 1, cur.length, 7).setValues(cur);
}
function tuneSet_(changes) {
  const s = tuneSheet_(), n = s.getLastRow();
  if (n < 2) return;
  const v = s.getRange(2, 1, n - 1, 6).getValues();
  const col = v.map(r => [r[4]]);
  const notes = [];
  Object.keys(changes).forEach(k => {
    const i = v.findIndex(r => String(r[0]).trim() === k);
    if (i < 0) return;
    const nv = String(changes[k] === null ? '' : changes[k]);
    col[i][0] = nv;
    notes.push(`${v[i][1]} ${v[i][2]} ${v[i][3]}：${nv === '' ? '恢復預設（' + v[i][5] + '）' : nv}`);
  });
  s.getRange(2, 5, col.length, 1).setValues(col);
  if (notes.length) log_('全班', notes.join('\n'));
}
function log_(who, text) { logSheet_().appendRow([new Date(), String(who || ''), String(text || '')]); }

// 玩家名單
function users_() {
  const s = usersSheet_(), n = s.getLastRow();
  if (n < 2) return [];
  const out = [];
  s.getRange(2, 1, n - 1, 10).getValues().forEach((r, i) => {
    const name = String(r[0]).trim();
    if (!name) return;
    let coins = Number(r[1] || 0);
    // 試算表上直接改過的金幣也要算進來
    try { const sv = JSON.parse(r[5]); if (r[1] === '' && sv) coins = sv.coins || 0; } catch (e) {}
    const t = r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'M/d HH:mm') : String(r[4] || '');
    out.push({ name, coins, total: Number(r[2] || 0), tier: r[6], score: Number(r[7] || 0), progress: r[8], updated: t });
  });
  return out.sort((a, b) => String(a.name).localeCompare(String(b.name), 'zh-Hant'));
}

// 調整的紀錄存在玩家存檔的 tOps（最多 30 筆），小朋友的裝置下次存檔時會補套
function addOp_(sv, op) {
  sv.tOps = Array.isArray(sv.tOps) ? sv.tOps : [];
  const last = sv.tOps.length ? sv.tOps[sv.tOps.length - 1].id : 0;
  op.id = Math.max(Date.now(), last + 1);
  sv.tOps.push(op);
  if (sv.tOps.length > 30) sv.tOps = sv.tOps.slice(-30);
  return op;
}
function applyOp_(sv, op) {
  ['owned', 'wpnOwned', 'fxOwned', 'cleared'].forEach(k => { if (!Array.isArray(sv[k])) sv[k] = []; });
  ['prog', 'items'].forEach(k => { if (!sv[k] || typeof sv[k] !== 'object') sv[k] = {}; });
  if (op.op === 'setCoins') sv.coins = Math.max(0, Math.round(Number(op.v) || 0));
  else if (op.op === 'addCoins') { const v = Math.round(Number(op.v) || 0); sv.coins = Math.max(0, (sv.coins || 0) + v); if (v > 0) sv.total = (sv.total || 0) + v; }
  else if (op.op === 'give') {
    const n = Math.max(1, Number(op.n) || 1);
    if (op.kind === 'skin' && sv.owned.indexOf(op.item) < 0) sv.owned.push(op.item);
    if (op.kind === 'wpn' && sv.wpnOwned.indexOf(op.item) < 0) sv.wpnOwned.push(op.item);
    if (op.kind === 'fx' && sv.fxOwned.indexOf(op.item) < 0) sv.fxOwned.push(op.item);
    if (op.kind === 'item') sv.items[op.item] = (sv.items[op.item] || 0) + n;
  }
  else if (op.op === 'unlockAll') {
    (op.stages || []).forEach(id => { sv.prog[id] = Math.max(sv.prog[id] || 0, 1); });
    (op.regions || []).forEach(id => { if (sv.cleared.indexOf(id) < 0) sv.cleared.push(id); });
  }
  else if (op.op === 'resetProg') { sv.prog = {}; sv.cleared = []; }
}
function userEdit_(name, op, label) {
  name = String(name || '').trim();
  const s = usersSheet_(), r = findRow_(s, name);
  if (r < 0) throw new Error('找不到這位玩家：' + name);
  if (!op || !op.op) throw new Error('不知道要改什麼');
  if (op.op === 'delete') { s.deleteRow(r); const rs = reportSheet_(), rr = findRow_(rs, name); if (rr > 0) rs.deleteRow(rr); log_(name, '刪除帳號'); return {}; }
  const sv = storedSave_(s, r) || {};
  const clean = { op: op.op, v: op.v, kind: op.kind, item: op.item, n: op.n, stages: op.stages, regions: op.regions };
  Object.keys(clean).forEach(k => clean[k] === undefined && delete clean[k]);
  addOp_(sv, clean);
  applyOp_(sv, clean);
  sv.tSeen = clean.id;
  sv.ts = Date.now();
  const row = s.getRange(r, 1, 1, 10).getValues()[0];
  row[1] = sv.coins || 0; row[2] = sv.total || 0; row[4] = new Date(); row[5] = JSON.stringify(sv);
  if (op.op === 'unlockAll' || op.op === 'resetProg') { row[8] = Object.keys(sv.prog).length + ' 關（老師調整）'; }
  s.getRange(r, 1, 1, 10).setValues([row]);
  log_(name, String(label || op.op));
  return { op: clean, user: users_().find(u => u.name === name) };
}

// ---------- 學習報告（每位小朋友一列，存檔時更新） ----------
function saveReport_(name, p) {
  const s = reportSheet_();
  let r = findRow_(s, name);
  if (r < 0) { s.appendRow([name]); r = s.getLastRow(); }
  s.getRange(r, 1, 1, 11).setValues([[name, p.total || 0, p.acc || '', p.total7 || 0, p.acc7 || '', p.trend || '', p.weak || '', p.confuse || '', p.suggest || '', p.practice || 0, new Date()]]);
}

// ---------- 自動更新（需要：Apps Script API 開關打開、appsscript.json 有 script.projects 和 script.deployments 權限） ----------
function selfUpdate_(deploymentId) {
  const res = UrlFetchApp.fetch(GITHUB_RAW + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('抓不到 GitHub 上的 Code.gs（' + res.getResponseCode() + '）');
  const src = res.getContentText();
  const m = src.match(/const CODE_VERSION = (\d+);/);
  const newVer = m ? Number(m[1]) : 0;
  if (!newVer || src.indexOf('function doPost') < 0) throw new Error('GitHub 上的 Code.gs 看起來不完整，沒有更新');
  if (newVer <= CODE_VERSION) return { updated: false, version: CODE_VERSION, msg: '已經是最新版（第 ' + CODE_VERSION + ' 版）' };
  const id = ScriptApp.getScriptId();
  const api = 'https://script.googleapis.com/v1/projects/' + id;
  const call = (method, path, body) => {
    const r = UrlFetchApp.fetch(api + path, { method, contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: body ? JSON.stringify(body) : undefined, muteHttpExceptions: true });
    const t = r.getContentText();
    if (r.getResponseCode() >= 300) {
      if (/has not been used|SERVICE_DISABLED|User has not enabled the Apps Script API/i.test(t)) throw new Error('Apps Script API 還沒打開：到 script.google.com/home/usersettings 打開');
      if (/insufficient|PERMISSION_DENIED|scope/i.test(t)) throw new Error('權限不夠：appsscript.json 要加上 script.projects 和 script.deployments，並重新執行「一次設定」授權');
      throw new Error('更新失敗（' + r.getResponseCode() + '）：' + t.slice(0, 200));
    }
    return t ? JSON.parse(t) : {};
  };
  const content = call('get', '/content');
  const files = content.files || [];
  let idx = files.findIndex(f => f.type === 'SERVER_JS' && /const CODE_VERSION = \d+;/.test(f.source || ''));
  if (idx < 0) idx = files.findIndex(f => f.type === 'SERVER_JS');
  if (idx < 0) throw new Error('找不到要替換的程式檔');
  files[idx] = { name: files[idx].name, type: 'SERVER_JS', source: src };
  call('put', '/content', { files: files.map(f => ({ name: f.name, type: f.type, source: f.source })) });
  const ver = call('post', '/versions', { description: '自動更新到第 ' + newVer + ' 版' });
  if (deploymentId) {
    call('put', '/deployments/' + deploymentId, { deploymentConfig: { scriptId: id, versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: '音符勇者 第 ' + newVer + ' 版' } });
  }
  log_('伺服器', '自動更新：第 ' + CODE_VERSION + ' 版 → 第 ' + newVer + ' 版');
  return { updated: true, from: CODE_VERSION, version: newVer, msg: '已更新到第 ' + newVer + ' 版' };
}
