/**
 * 群組待辦板 — 後端
 *
 * 這支程式住在你的 Google 試算表裡，負責讀寫資料。
 * 資料在同一份試算表的「任務」「紀錄」「重複設定」三個分頁，你隨時打開就看得到。
 * 照片存到你雲端硬碟的「群組待辦板照片」資料夾。
 *
 * 編輯密碼不寫在這裡（這個檔案在 GitHub 上是公開的）。
 * 它存在 Apps Script 的「專案設定 → 指令碼屬性」，名稱是 EDIT_PASSWORD。
 * 沒有設定密碼時，所有人都只能看、不能改。
 */

var APP_TZ       = 'Asia/Taipei';   // 「今天」以台灣時間為準，不受專案時區設定影響
var TRASH_DAYS   = 30;
var TASK_SHEET   = '任務';
var LOG_SHEET    = '紀錄';
var REPEAT_SHEET = '重複設定';
var PHOTO_FOLDER = '群組待辦板照片';

var TASK_HEADERS = ['編號','標題','分類','說明','狀態','建立者','建立時間',
                    '完工者','完工時間','完工備註','刪除者','刪除時間','照片','系統編號'];
var LOG_HEADERS  = ['時間','任務編號','任務標題','誰','動作','內容'];
var REPEAT_HEADERS = ['編號','標題','分類','說明','頻率','設定','參數','下次出現',
                      '建立者','建立時間','狀態','系統編號','上次產生','最後修改者','最後修改時間'];

var COL = { no:0, title:1, site:2, desc:3, status:4, createdBy:5, createdAt:6,
            doneBy:7, doneAt:8, doneNote:9, delBy:10, delAt:11, photos:12, id:13 };

var RCOL = { no:0, title:1, site:2, desc:3, freq:4, label:5, param:6, next:7,
             by:8, at:9, status:10, id:11, lastGen:12, editBy:13, editAt:14 };

var FREQS = ['每天','每週','每月','每年'];
var OPEN = '待辦中', DONE = '已完工', TRASH = '刪除區';
var ACTIVE = '啟用', PAUSED = '停用';

/* ---------------- 入口 ---------------- */

function doGet(e) {
  return json_({ ok: true, hint: '待辦板後端正常運作中。請從網頁開啟。' });
}

var WRITES = {
  add:        function (q, w) { return addTask_(q, w); },
  edit:       function (q, w) { return editTask_(q, w); },
  done:       function (q, w) { return finishTask_(q, w); },
  photos:     function (q, w) { return addPhotos_(q, w); },
  note:       function (q, w) { return addNote_(q, w); },
  reopen:     function (q, w) { return reopenTask_(q, w); },
  trash:      function (q, w) { return trashTask_(q, w); },
  restore:    function (q, w) { return restoreTask_(q, w); },
  addRepeat:  function (q, w) { return addRepeat_(q, w); },
  editRepeat: function (q, w) { return editRepeat_(q, w); },
  delRepeat:  function (q, w) { return delRepeat_(q, w); }
};

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var who = String(req.who || '').trim().slice(0, 20) || '（未署名）';

    if (req.action === 'list') {
      var auth0 = checkEdit_(req.pw);
      var board0 = listTasks_(false);
      board0.canEdit = auth0.ok;
      board0.passwordSet = auth0.passwordSet;
      return json_(board0);
    }

    if (req.action === 'auth') {
      var a = checkEdit_(req.pw);
      if (!a.passwordSet) return json_({ ok: false, error: '還沒有設定編輯密碼，請找板子的建立者' });
      if (!a.ok) return json_({ ok: false, wrongPassword: true, error: a.error || '密碼不對' });
      return json_({ ok: true, canEdit: true });
    }

    var fn = WRITES[req.action];
    if (!fn) throw new Error('不認得的動作：' + req.action);

    var auth = checkEdit_(req.pw);
    if (!auth.ok) {
      return json_({ ok: false, readOnly: true,
                     error: auth.error || (auth.passwordSet
                       ? '你目前是唯讀，要先輸入編輯密碼才能修改'
                       : '還沒有設定編輯密碼，目前所有人都只能看') });
    }

    // 寫入完直接把最新的板子一起回傳，前端就不必再跑一趟
    var res = fn(req, who) || {};
    var board = listTasks_(req.action === 'addRepeat' || req.action === 'editRepeat');
    board.canEdit = true;
    board.passwordSet = true;
    if (res.id) board.id = res.id;
    return json_(board);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 編輯密碼 ---------------- */

var FAIL_LIMIT = 30;          // 10 分鐘內錯這麼多次，就先暫停驗證
var FAIL_WINDOW_SEC = 600;

/**
 * 密碼放在「指令碼屬性」，不在程式碼裡。
 * 猜錯會慢一秒回應，而且 10 分鐘內錯太多次會暫停驗證，擋住一直猜密碼的人。
 */
function checkEdit_(pw) {
  var saved = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  if (!saved) return { ok: false, passwordSet: false };
  pw = String(pw || '');
  if (!pw) return { ok: false, passwordSet: true };

  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pwFails') || 0);
  if (fails >= FAIL_LIMIT) {
    return { ok: false, passwordSet: true, error: '密碼錯太多次，請 10 分鐘後再試' };
  }
  if (pw === String(saved)) return { ok: true, passwordSet: true };

  cache.put('pwFails', String(fails + 1), FAIL_WINDOW_SEC);
  Utilities.sleep(1000);
  return { ok: false, passwordSet: true, error: '密碼不對' };
}

/* ---------------- 日期（一律用台灣的年月日） ---------------- */

function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function ymdStr_(o) { return o.y + '-' + pad2_(o.m) + '-' + pad2_(o.d); }
function parseYmd_(s) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(s || ''));
  return m ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) } : null;
}
function todayYmd_() {
  return parseYmd_(Utilities.formatDate(new Date(), APP_TZ, 'yyyy-MM-dd'));
}
function addDaysYmd_(o, n) {
  var d = new Date(Date.UTC(o.y, o.m - 1, o.d + n));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}
function cmpYmd_(a, b) {
  return (a.y * 10000 + a.m * 100 + a.d) - (b.y * 10000 + b.m * 100 + b.d);
}
function daysInMonth_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

/**
 * 試算表裡的日期欄位可能是真的日期（舊資料），也可能是「2026-12-09」這種文字（新資料）。
 * 真的日期要用試算表自己的時區解讀，才不會差一天。
 */
function cellYmd_(v) {
  if (!v) return null;
  if (v instanceof Date) {
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || APP_TZ;
    return parseYmd_(Utilities.formatDate(v, tz, 'yyyy-MM-dd'));
  }
  return parseYmd_(v);
}

/* ---------------- 試算表 ---------------- */

// Apps Script 每次請求都是全新的執行環境，所以這個暫存只活在單次請求裡，
// 用來避免同一份分頁在一次請求中被重覆抓好幾次（每次都是一趟遠端呼叫）。
var SHEETS_ = {};

function sheet_(name, headers) {
  if (SHEETS_[name]) return SHEETS_[name];
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    // 新版本多了欄位，把缺的標題補上
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  SHEETS_[name] = sh;
  return sh;
}
function tasksSheet_()  { return sheet_(TASK_SHEET, TASK_HEADERS); }
function logsSheet_()   { return sheet_(LOG_SHEET, LOG_HEADERS); }
function repeatSheet_() { return sheet_(REPEAT_SHEET, REPEAT_HEADERS); }

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function findRow_(sh, idCol, id, missing) {
  var data = rows_(sh);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return { row: i + 2, values: data[i] };
  }
  throw new Error(missing || '找不到這件待辦，可能已經被別人刪掉了');
}

/**
 * 使用者打的字一律當文字存。
 * 不這樣做的話，標題「=1+1」會變成公式、「8/18」會變成日期、「-帶鑰匙」會變成 #ERROR!。
 * 前面加的單引號是試算表的「這是文字」記號，存進去之後不會顯示、讀出來也不會帶著。
 */
function txt_(s) {
  s = String(s == null ? '' : s);
  return s ? "'" + s : '';
}

/** 舊資料裡已經被試算表轉成日期或數字的標題，讀出來時轉回人看得懂的文字。 */
function cellText_(v) {
  if (v instanceof Date) {
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || APP_TZ;
    return Utilities.formatDate(v, tz, 'M/d');
  }
  return String(v == null ? '' : v);
}

function stamp_(v) {
  if (!v) return 0;
  if (v instanceof Date) return v.getTime();
  var d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function lock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/**
 * 編號永遠往上加、不回收。
 * 以前是「目前最大的編號 + 1」，最大的那件被永久刪掉之後，新的任務會拿到同一個編號，
 * 而紀錄是靠編號對應的——直接在試算表刪掉一列的話，新任務會繼承到舊任務的紀錄。
 */
function nextNo_(key, sh, noCol, logPrefix) {
  var props = PropertiesService.getScriptProperties();
  var n = Number(props.getProperty(key) || 0);
  if (!n) {
    var max = rows_(sh).reduce(function (m, r) { return Math.max(m, Number(r[noCol]) || 0); }, 0);
    rows_(logsSheet_()).forEach(function (r) {
      var raw = String(r[1]);
      if (logPrefix ? raw.indexOf(logPrefix) !== 0 : !/^\d+$/.test(raw)) return;
      max = Math.max(max, Number(raw.slice(logPrefix ? logPrefix.length : 0)) || 0);
    });
    n = max + 1;
  }
  props.setProperty(key, String(n + 1));
  return n;
}

/* ---------------- 照片 ---------------- */

function photoFolder_() {
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty('photoFolderId');
  if (saved) {
    try { return DriveApp.getFolderById(saved); } catch (e) { /* 被刪了就重建 */ }
  }
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('photoFolderId', folder.getId());
  return folder;
}

function savePhoto_(dataUrl) {
  var m = /^data:(image\/(?:jpeg|png));base64,([\s\S]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('照片格式不支援');
  var name = 'photo_' + new Date().getTime() + '_' +
             Math.random().toString(36).slice(2, 7) + (m[1] === 'image/png' ? '.png' : '.jpg');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  var file = photoFolder_().createFile(blob);
  // 網頁要能顯示，檔案必須是「知道連結的人可以看」
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* 公司帳號可能禁止，照片就只有你看得到 */ }
  return file.getId();
}

function photoObj_(id) {
  return {
    id: id,
    url:   'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600',
    thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400',
    open:  'https://drive.google.com/file/d/' + id + '/view'
  };
}

function photoIds_(cell) {
  return String(cell || '').split(',').map(function (s) { return s.trim(); })
                           .filter(function (s) { return s; });
}

/* ---------------- 紀錄 ---------------- */

/** 任務的紀錄用編號（例如 12），重複設定的紀錄用 R 加編號（例如 R3），兩邊不會撞號。 */
function log_(no, title, who, action, text) {
  logsSheet_().appendRow([new Date(), txt_(no), txt_(title), txt_(who), action, txt_(text)]);
}

function logsByNo_(logRows) {
  var out = {};
  logRows.forEach(function (r) {
    var no = String(r[1]);
    (out[no] = out[no] || []).push({ at: stamp_(r[0]), who: String(r[3] || ''),
                                     action: String(r[4] || ''), text: String(r[5] || '') });
  });
  return out;
}

/**
 * 這塊板子上出現過的人，最近用過的排前面。
 * 手機的內建瀏覽器常常把本機記憶清掉，有這份名單就能一鍵選回自己，不必重打。
 */
function knownPeople_(taskRows, logRows) {
  var all = [];
  logRows.forEach(function (r) { all.push({ who: String(r[3] || ''), at: stamp_(r[0]) }); });
  taskRows.forEach(function (r) {
    if (r[COL.createdBy]) all.push({ who: String(r[COL.createdBy]), at: stamp_(r[COL.createdAt]) });
    if (r[COL.doneBy]) all.push({ who: String(r[COL.doneBy]), at: stamp_(r[COL.doneAt]) });
  });
  all.sort(function (a, b) { return b.at - a.at; });

  var seen = {}, out = [];
  for (var i = 0; i < all.length && out.length < 12; i++) {
    var n = String(all[i].who || '').trim();
    if (!n || n === '（未署名）' || n === '系統' || seen[n]) continue;
    seen[n] = true;
    out.push(n);
  }
  return out;
}

function short_(s, n) {
  s = String(s == null ? '' : s);
  if (!s) return '（空白）';
  return s.length > n ? s.slice(0, n) + '…' : s;
}

/* ---------------- 每日維護 ---------------- */

/**
 * 清理刪除區、產生到期的重複任務。
 * 一天做一次就夠（以台灣日期為準）；新增或修改重複設定時會強制再跑一次。
 */
function maybeMaintain_(force) {
  var props = PropertiesService.getScriptProperties();
  var today = ymdStr_(todayYmd_());
  if (!force && props.getProperty('lastMaintainDay') === today) return;

  lock_(function () {
    // 拿到鎖之後再確認一次，避免兩個人同時開板子時重覆跑
    if (!force && props.getProperty('lastMaintainDay') === today) return;

    // 早期版本這欄叫「案場」，順手改成現在的名稱
    var sh = tasksSheet_();
    if (sh.getLastRow() >= 1 && sh.getRange(1, COL.site + 1).getValue() === '案場') {
      sh.getRange(1, COL.site + 1).setValue('分類');
    }
    purge_();
    repairLegacyRepeats_();
    generateDue_();
    props.setProperty('lastMaintainDay', today);
  });
}

/* ---------------- 讀取 ---------------- */

function listTasks_(force) {
  maybeMaintain_(force);
  var logRows = rows_(logsSheet_());
  var taskRows = rows_(tasksSheet_());
  var logs = logsByNo_(logRows);
  var tasks = taskRows.map(function (r) {
    var status = r[COL.status] === DONE ? 'done' : (r[COL.status] === TRASH ? 'trash' : 'open');
    return {
      id: String(r[COL.id]),
      no: Number(r[COL.no]) || 0,
      title: cellText_(r[COL.title]),
      site: cellText_(r[COL.site]),
      desc: cellText_(r[COL.desc]),
      status: status,
      createdBy: String(r[COL.createdBy] || ''),
      createdAt: stamp_(r[COL.createdAt]),
      doneBy: String(r[COL.doneBy] || ''),
      doneAt: stamp_(r[COL.doneAt]),
      doneNote: String(r[COL.doneNote] || ''),
      deletedBy: String(r[COL.delBy] || ''),
      deletedAt: stamp_(r[COL.delAt]),
      photos: photoIds_(r[COL.photos]).map(photoObj_),
      logs: logs[String(r[COL.no])] || []
    };
  });

  var repeats = rows_(repeatSheet_())
    .filter(function (r) { return r[RCOL.title]; })
    .map(function (r) {
      var rule = parseRule_(r[RCOL.freq], r[RCOL.param]);
      var next = cellYmd_(r[RCOL.next]);
      var nextLunar = next ? solarToLunar(next.y, next.m, next.d) : null;
      return {
        id: String(r[RCOL.id]),
        no: Number(r[RCOL.no]) || 0,
        title: cellText_(r[RCOL.title]),
        site: cellText_(r[RCOL.site]),
        desc: cellText_(r[RCOL.desc]),
        freq: rule.freq,
        cal: rule.cal,
        weekday: rule.weekday,
        month: rule.month,
        day: rule.day,
        label: ruleLabel_(rule),
        next: next ? ymdStr_(next) : '',
        nextLunar: nextLunar ? lunarText(nextLunar) : '',
        by: String(r[RCOL.by] || ''),
        at: stamp_(r[RCOL.at]),
        paused: r[RCOL.status] === PAUSED,
        editBy: String(r[RCOL.editBy] || ''),
        editAt: stamp_(r[RCOL.editAt]),
        logs: logs['R' + (Number(r[RCOL.no]) || 0)] || []
      };
    });

  return { ok: true, tasks: tasks, repeats: repeats, today: ymdStr_(todayYmd_()),
           people: knownPeople_(taskRows, logRows), trashDays: TRASH_DAYS };
}

/* ---------------- 任務 ---------------- */

function addTask_(req, who) {
  return lock_(function () {
    var title = String(req.title || '').trim().slice(0, 120);
    if (!title) throw new Error('先寫要做什麼');
    var site = String(req.site || '').trim().slice(0, 60);
    var sh = tasksSheet_();
    var no = nextNo_('nextTaskNo', sh, COL.no, '');
    var id = 't' + new Date().getTime() + Math.random().toString(36).slice(2, 6);

    sh.appendRow([no, txt_(title), txt_(site), txt_(String(req.desc || '').trim().slice(0, 2000)),
                  OPEN, txt_(who), new Date(), '', '', '', '', '', '', id]);
    log_(no, title, who, '建立', site ? '分類：' + site : '');
    return { ok: true, id: id };
  });
}

function finishTask_(req, who) {
  var ids = (req.photos || []).slice(0, 6).map(savePhoto_);
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);
    var note = String(req.note || '').trim().slice(0, 1000);
    var kept = photoIds_(hit.values[COL.photos]).concat(ids);

    sh.getRange(hit.row, COL.status + 1).setValue(DONE);
    sh.getRange(hit.row, COL.doneBy + 1, 1, 3).setValues([[txt_(who), new Date(), txt_(note)]]);
    if (ids.length) sh.getRange(hit.row, COL.photos + 1).setValue(kept.join(','));

    var no = hit.values[COL.no], title = hit.values[COL.title];
    if (ids.length) log_(no, title, who, '加照片', '上傳 ' + ids.length + ' 張完工照');
    log_(no, title, who, '完工', note);
    return { ok: true };
  });
}

function addPhotos_(req, who) {
  var ids = (req.photos || []).slice(0, 6).map(savePhoto_);
  if (!ids.length) throw new Error('沒有收到照片');
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);
    var kept = photoIds_(hit.values[COL.photos]).concat(ids);
    sh.getRange(hit.row, COL.photos + 1).setValue(kept.join(','));
    log_(hit.values[COL.no], hit.values[COL.title], who, '加照片', '上傳 ' + ids.length + ' 張照片');
    return { ok: true };
  });
}

function addNote_(req, who) {
  var text = String(req.text || '').trim().slice(0, 1000);
  if (!text) throw new Error('先寫點東西');
  return lock_(function () {
    var hit = findRow_(tasksSheet_(), COL.id, req.id);
    log_(hit.values[COL.no], hit.values[COL.title], who, '留言', text);
    return { ok: true };
  });
}

/**
 * 修改待辦內容。改了什麼一定寫進紀錄，包含被覆蓋掉的舊值，
 * 這樣事後翻紀錄才知道原本寫的是什麼。
 */
function editTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);

    var title = String(req.title || '').trim().slice(0, 120);
    if (!title) throw new Error('標題不能空白');
    var site = String(req.site || '').trim().slice(0, 60);
    var desc = String(req.desc || '').trim().slice(0, 2000);

    var oldTitle = cellText_(hit.values[COL.title]);
    var oldSite = cellText_(hit.values[COL.site]);
    var oldDesc = cellText_(hit.values[COL.desc]);

    var changes = [];
    if (title !== oldTitle) changes.push('標題：' + short_(oldTitle, 20) + ' → ' + short_(title, 20));
    if (site !== oldSite) changes.push('分類：' + short_(oldSite, 15) + ' → ' + short_(site, 15));
    if (desc !== oldDesc) changes.push('說明改過（原本：' + short_(oldDesc, 30) + '）');
    if (!changes.length) return { ok: true };

    // 標題、分類、說明剛好是相鄰三欄，一次寫完
    sh.getRange(hit.row, COL.title + 1, 1, 3).setValues([[txt_(title), txt_(site), txt_(desc)]]);
    log_(hit.values[COL.no], title, who, '修改', changes.join('；'));
    return { ok: true };
  });
}

function reopenTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);
    sh.getRange(hit.row, COL.status + 1).setValue(OPEN);
    sh.getRange(hit.row, COL.doneBy + 1, 1, 3).setValues([['', '', '']]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '退回待辦', '');
    return { ok: true };
  });
}

function trashTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);
    sh.getRange(hit.row, COL.status + 1).setValue(TRASH);
    sh.getRange(hit.row, COL.delBy + 1, 1, 2).setValues([[txt_(who), new Date()]]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '移到刪除區', '');
    return { ok: true };
  });
}

function restoreTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, COL.id, req.id);
    var back = hit.values[COL.doneAt] ? DONE : OPEN;
    sh.getRange(hit.row, COL.status + 1).setValue(back);
    sh.getRange(hit.row, COL.delBy + 1, 1, 2).setValues([['', '']]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '還原', '');
    return { ok: true };
  });
}

/* ---------------- 重複設定 ---------------- */

/**
 * 規則存在「頻率」「參數」兩欄。參數格式：
 *   每週 → 3（星期三；7 是星期日）
 *   每月 → 15（國曆 15 號）　　　　 L:15（農曆十五）
 *   每年 → 9-1（國曆 9 月 1 日）　　L:8-15（農曆八月十五）
 * 舊資料沒有 L: 開頭，一律當國曆，照常運作。
 */
function parseRule_(freq, param) {
  freq = String(freq || '');
  // 舊版把「12-1」這種參數直接寫進試算表，被自動轉成了 12 月 1 日這個日期。取回原本的月、日。
  if (param instanceof Date) {
    var pd = cellYmd_(param);
    param = pd ? (freq === '每年' ? pd.m + '-' + pd.d : String(pd.d)) : '';
  }
  param = String(param == null ? '' : param);
  var cal = 'G';
  if (param.indexOf('L:') === 0) { cal = 'L'; param = param.slice(2); }
  var rule = { freq: freq, cal: cal, weekday: 0, month: 0, day: 0 };
  if (freq === '每週') rule.weekday = Math.min(7, Math.max(1, Number(param) || 1));
  if (freq === '每月') rule.day = Number(param) || 1;
  if (freq === '每年') {
    var p = param.split('-');
    rule.month = Number(p[0]) || 1;
    rule.day = Number(p[1]) || 1;
  }
  if (freq === '每天' || freq === '每週') rule.cal = 'G';
  if (rule.day) rule.day = Math.min(rule.cal === 'L' ? 30 : 31, Math.max(1, rule.day));
  if (rule.month) rule.month = Math.min(12, Math.max(1, rule.month));
  return rule;
}

function ruleParam_(rule) {
  var pre = rule.cal === 'L' ? 'L:' : '';
  if (rule.freq === '每週') return String(rule.weekday);
  if (rule.freq === '每月') return pre + rule.day;
  if (rule.freq === '每年') return pre + rule.month + '-' + rule.day;
  return '';
}

/** 從前端送來的欄位組出規則，順便擋掉不合理的值。 */
function ruleFromReq_(req) {
  var freq = String(req.freq || '');
  if (FREQS.indexOf(freq) < 0) throw new Error('頻率設定不正確');
  var cal = (req.cal === 'L' && (freq === '每月' || freq === '每年')) ? 'L' : 'G';
  var rule = { freq: freq, cal: cal, weekday: 0, month: 0, day: 0 };
  var maxDay = cal === 'L' ? 30 : 31;
  if (freq === '每週') {
    rule.weekday = Number(req.weekday);
    if (!(rule.weekday >= 1 && rule.weekday <= 7)) throw new Error('星期設定不正確');
  }
  if (freq === '每月' || freq === '每年') {
    rule.day = Number(req.day);
    if (!(rule.day >= 1 && rule.day <= maxDay)) throw new Error('日期設定不正確');
  }
  if (freq === '每年') {
    rule.month = Number(req.month);
    if (!(rule.month >= 1 && rule.month <= 12)) throw new Error('月份設定不正確');
  }
  return rule;
}

function ruleLabel_(rule) {
  var W = ['', '週一', '週二', '週三', '週四', '週五', '週六', '週日'];
  if (rule.freq === '每天') return '每天';
  if (rule.freq === '每週') return '每' + W[rule.weekday];
  if (rule.freq === '每月') {
    return rule.cal === 'L' ? '每月 農曆' + lunarDayName(rule.day) : '每月 ' + rule.day + ' 號';
  }
  if (rule.freq === '每年') {
    return rule.cal === 'L'
      ? '每年 農曆' + LUNAR_MONTH_NAMES[rule.month] + lunarDayName(rule.day)
      : '每年 ' + rule.month + ' 月 ' + rule.day + ' 日';
  }
  return rule.freq;
}

/** 從 from 當天算起（含當天），下一次符合規則的日子。from、回傳值都是 {y,m,d}。 */
function nextOccurrence_(rule, from) {
  if (rule.freq === '每天') return from;

  if (rule.freq === '每週') {
    var target = rule.weekday % 7;          // 1=週一 … 7=週日 → JS 的 1…6,0
    var d = from;
    for (var i = 0; i < 7; i++) {
      if (new Date(Date.UTC(d.y, d.m - 1, d.d)).getUTCDay() === target) return d;
      d = addDaysYmd_(d, 1);
    }
    return from;
  }

  if (rule.freq === '每月') {
    if (rule.cal === 'L') return nextLunarMonthly(rule.day, from);
    var y = from.y, m = from.m;
    for (var k = 0; k < 24; k++) {
      // 31 號遇到只有 30 天的月份，就落在該月最後一天
      var cand = { y: y, m: m, d: Math.min(rule.day, daysInMonth_(y, m)) };
      if (cmpYmd_(cand, from) >= 0) return cand;
      m++;
      if (m > 12) { m = 1; y++; }
    }
    return from;
  }

  if (rule.freq === '每年') {
    if (rule.cal === 'L') return nextLunarYearly(rule.month, rule.day, from);
    for (var j = 0; j < 4; j++) {
      var c = { y: from.y + j, m: rule.month,
                d: Math.min(rule.day, daysInMonth_(from.y + j, rule.month)) };
      if (cmpYmd_(c, from) >= 0) return c;
    }
  }
  return from;
}

/** 今天已經產生過的，就從明天開始算，避免改完設定當天又多長一件。 */
function scheduleFrom_(lastGen) {
  var today = todayYmd_();
  var last = cellYmd_(lastGen);
  return (last && cmpYmd_(last, today) >= 0) ? addDaysYmd_(today, 1) : today;
}

function addRepeat_(req, who) {
  return lock_(function () {
    var title = String(req.title || '').trim().slice(0, 120);
    if (!title) throw new Error('先寫要做什麼');
    var rule = ruleFromReq_(req);
    var sh = repeatSheet_();
    var no = nextNo_('nextRepeatNo', sh, RCOL.no, 'R');
    var id = 'r' + new Date().getTime() + Math.random().toString(36).slice(2, 6);
    var next = nextOccurrence_(rule, todayYmd_());
    var label = ruleLabel_(rule);

    sh.appendRow([no, txt_(title), txt_(String(req.site || '').trim().slice(0, 60)),
                  txt_(String(req.desc || '').trim().slice(0, 2000)),
                  rule.freq, label, txt_(ruleParam_(rule)), next ? txt_(ymdStr_(next)) : '',
                  txt_(who), new Date(), ACTIVE, id, '', '', '']);
    log_('R' + no, title, who, '建立重複', label);
    return { ok: true, id: id };
  });
}

/**
 * 修改重複設定：內容、週期、暫停都能改。
 * 和待辦一樣，改了什麼、原本是什麼都寫進紀錄，並記下最後修改的人和時間。
 */
function editRepeat_(req, who) {
  return lock_(function () {
    var sh = repeatSheet_();
    var hit = findRow_(sh, RCOL.id, req.id, '找不到這項重複設定，可能已經被別人刪掉了');
    var v = hit.values;

    var title = String(req.title || '').trim().slice(0, 120);
    if (!title) throw new Error('標題不能空白');
    var site = String(req.site || '').trim().slice(0, 60);
    var desc = String(req.desc || '').trim().slice(0, 2000);
    var rule = ruleFromReq_(req);
    var paused = !!req.paused;

    var oldRule = parseRule_(v[RCOL.freq], v[RCOL.param]);
    var oldLabel = ruleLabel_(oldRule), newLabel = ruleLabel_(rule);
    var wasPaused = v[RCOL.status] === PAUSED;

    var changes = [];
    if (title !== cellText_(v[RCOL.title])) {
      changes.push('標題：' + short_(v[RCOL.title], 20) + ' → ' + short_(title, 20));
    }
    if (site !== cellText_(v[RCOL.site])) {
      changes.push('分類：' + short_(v[RCOL.site], 15) + ' → ' + short_(site, 15));
    }
    if (desc !== cellText_(v[RCOL.desc])) {
      changes.push('說明改過（原本：' + short_(v[RCOL.desc], 30) + '）');
    }
    var ruleChanged = ruleParam_(rule) !== ruleParam_(oldRule) || rule.freq !== oldRule.freq;
    if (ruleChanged) changes.push('週期：' + oldLabel + ' → ' + newLabel);
    if (paused !== wasPaused) changes.push(paused ? '暫停' : '恢復');
    if (!changes.length) return { ok: true };

    var next = cellYmd_(v[RCOL.next]);
    if (ruleChanged || (wasPaused && !paused) || !next) {
      next = nextOccurrence_(rule, scheduleFrom_(v[RCOL.lastGen]));
    }

    // 標題到下次出現是相鄰七欄，一次寫完
    sh.getRange(hit.row, RCOL.title + 1, 1, 7).setValues([[
      txt_(title), txt_(site), txt_(desc), rule.freq, newLabel, txt_(ruleParam_(rule)),
      next ? txt_(ymdStr_(next)) : ''
    ]]);
    sh.getRange(hit.row, RCOL.status + 1).setValue(paused ? PAUSED : ACTIVE);
    sh.getRange(hit.row, RCOL.editBy + 1, 1, 2).setValues([[txt_(who), new Date()]]);
    log_('R' + (Number(v[RCOL.no]) || 0), title, who, '修改', changes.join('；'));
    return { ok: true };
  });
}

function delRepeat_(req, who) {
  return lock_(function () {
    var sh = repeatSheet_();
    var hit = findRow_(sh, RCOL.id, req.id, '找不到這項重複設定');
    sh.deleteRow(hit.row);
    log_('R' + (Number(hit.values[RCOL.no]) || 0), hit.values[RCOL.title], who, '刪除重複',
         ruleLabel_(parseRule_(hit.values[RCOL.freq], hit.values[RCOL.param])));
    return { ok: true };
  });
}

/**
 * 舊版的每年重複設定，參數「月-日」被試算表存成了日期。
 * 舊程式讀不懂，觸發一次之後就把下次出現改成 1 月 1 日。
 * 這裡把參數改回文字；下次出現如果在未來，就照正確的月日重算（今天之前或今天的照常觸發）。
 * 每一筆修正都記在紀錄裡。
 */
function repairLegacyRepeats_() {
  var sh = repeatSheet_();
  var data = rows_(sh);
  var today = todayYmd_();
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!(r[RCOL.param] instanceof Date) || !r[RCOL.title]) continue;
    var rule = parseRule_(r[RCOL.freq], r[RCOL.param]);
    var label = ruleLabel_(rule);
    var oldNext = cellYmd_(r[RCOL.next]);
    var next = oldNext;
    if (!oldNext || cmpYmd_(oldNext, today) > 0) {
      next = nextOccurrence_(rule, addDaysYmd_(today, 1));
    }
    sh.getRange(i + 2, RCOL.label + 1, 1, 3).setValues([[
      label, txt_(ruleParam_(rule)), next ? txt_(ymdStr_(next)) : ''
    ]]);
    var moved = oldNext && next && cmpYmd_(oldNext, next) !== 0;
    log_('R' + (Number(r[RCOL.no]) || 0), r[RCOL.title], '系統', '自動修正',
         '週期設定被試算表誤存成日期，已修正為「' + label + '」' +
         (moved ? '；下次出現 ' + ymdStr_(oldNext) + ' → ' + ymdStr_(next) : ''));
  }
}

/**
 * 到期的重複設定會自動長出一件待辦。
 * 若中間隔了好幾期沒人開板子，只補最近的一件，不會一次灌一堆進來。
 */
function generateDue_() {
  var sh = repeatSheet_();
  var data = rows_(sh);
  if (!data.length) return;
  var today = todayYmd_();

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[RCOL.title] || r[RCOL.status] === PAUSED) continue;
    var rule = parseRule_(r[RCOL.freq], r[RCOL.param]);

    var next = cellYmd_(r[RCOL.next]);
    if (!next) {
      var first = nextOccurrence_(rule, scheduleFrom_(r[RCOL.lastGen]));
      sh.getRange(i + 2, RCOL.next + 1).setValue(first ? txt_(ymdStr_(first)) : '');
      continue;
    }
    if (cmpYmd_(next, today) > 0) continue;

    createFromRepeat_(r, rule);
    var after = nextOccurrence_(rule, addDaysYmd_(today, 1));
    sh.getRange(i + 2, RCOL.next + 1).setValue(after ? txt_(ymdStr_(after)) : '');
    sh.getRange(i + 2, RCOL.lastGen + 1).setValue(txt_(ymdStr_(today)));
  }
}

function createFromRepeat_(r, rule) {
  var sh = tasksSheet_();
  var no = nextNo_('nextTaskNo', sh, COL.no, '');
  var id = 't' + new Date().getTime() + Math.random().toString(36).slice(2, 6);
  sh.appendRow([no, txt_(r[RCOL.title]), txt_(r[RCOL.site]), txt_(r[RCOL.desc]), OPEN,
                txt_(r[RCOL.by]), new Date(), '', '', '', '', '', '', id]);
  log_(no, r[RCOL.title], r[RCOL.by], '自動建立', '來自重複設定：' + ruleLabel_(rule));
}

/**
 * 刪除區裡超過 30 天的，連照片帶紀錄一起永久刪除。
 * 每天第一次有人開板子時順手檢查，不需要另外設定排程。
 */
function purge_() {
  var cutoff = new Date().getTime() - TRASH_DAYS * 86400000;
  var sh = tasksSheet_();
  var data = rows_(sh);
  var goneNos = [];

  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    if (r[COL.status] !== TRASH) continue;
    var at = stamp_(r[COL.delAt]);
    if (!at || at >= cutoff) continue;

    photoIds_(r[COL.photos]).forEach(function (pid) {
      try { DriveApp.getFileById(pid).setTrashed(true); } catch (e) { /* 已經不在了 */ }
    });
    sh.deleteRow(i + 2);
    goneNos.push(String(r[COL.no]));
  }

  if (goneNos.length) {
    var ls = logsSheet_();
    var lrows = rows_(ls);
    for (var j = lrows.length - 1; j >= 0; j--) {
      if (goneNos.indexOf(String(lrows[j][1])) >= 0) ls.deleteRow(j + 2);
    }
  }
}

/* ================= 農曆 =================
 * 表格涵蓋 1900–2100 年，以東經 120 度（台灣、中國共用的 UTC+8）計算。
 * 每個數字編碼一年：低 4 位元是閏月是第幾月（0 表示沒有閏月），
 * 第 5–16 位元依序是 1–12 月是大月（30 天）還是小月（29 天），第 17 位元是閏月大小。
 * 1933、2097 兩年依天文計算修正過；已逐日比對 1901–2100 共 73,049 天，與天文演算法一致。
 * 這段程式在後端和網頁裡一字不差地各放一份，兩邊算出來的日子才會一致。
 */
var LUNAR_INFO = [
  0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2, // 1900
  0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977, // 1910
  0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970, // 1920
  0x06566,0x0d4a0,0x0ea50,0x16a95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950, // 1930
  0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557, // 1940
  0x06ca0,0x0b550,0x15355,0x04da0,0x0a5b0,0x14573,0x052b0,0x0a9a8,0x0e950,0x06aa0, // 1950
  0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0, // 1960
  0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b6a0,0x195a6, // 1970
  0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570, // 1980
  0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06b58,0x05ac0,0x0ab60,0x096d5,0x092e0, // 1990
  0x0c960,0x0d954,0x0d4a0,0x0da50,0x07552,0x056a0,0x0abb7,0x025d0,0x092d0,0x0cab5, // 2000
  0x0a950,0x0b4a0,0x0baa4,0x0ad50,0x055d9,0x04ba0,0x0a5b0,0x15176,0x052b0,0x0a930, // 2010
  0x07954,0x06aa0,0x0ad50,0x05b52,0x04b60,0x0a6e6,0x0a4e0,0x0d260,0x0ea65,0x0d530, // 2020
  0x05aa0,0x076a3,0x096d0,0x04afb,0x04ad0,0x0a4d0,0x1d0b6,0x0d250,0x0d520,0x0dd45, // 2030
  0x0b5a0,0x056d0,0x055b2,0x049b0,0x0a577,0x0a4b0,0x0aa50,0x1b255,0x06d20,0x0ada0, // 2040
  0x14b63,0x09370,0x049f8,0x04970,0x064b0,0x168a6,0x0ea50,0x06b20,0x1a6c4,0x0aae0, // 2050
  0x092e0,0x0d2e3,0x0c960,0x0d557,0x0d4a0,0x0da50,0x05d55,0x056a0,0x0a6d0,0x055d4, // 2060
  0x052d0,0x0a9b8,0x0a950,0x0b4a0,0x0b6a6,0x0ad50,0x055a0,0x0aba4,0x0a5b0,0x052b0, // 2070
  0x0b273,0x06930,0x07337,0x06aa0,0x0ad50,0x14b55,0x04b60,0x0a570,0x054e4,0x0d160, // 2080
  0x0e968,0x0d520,0x0daa0,0x16aa6,0x056d0,0x04ae0,0x0a9d4,0x0a2d0,0x0d150,0x0f252, // 2090
  0x0d520 // 2100
];
var LUNAR_MIN = 1900, LUNAR_MAX = 2100;
var LUNAR_BASE_UTC = Date.UTC(1900, 0, 31);   // 農曆 1900 年正月初一

function lunarLeapMonth(y) { return LUNAR_INFO[y - LUNAR_MIN] & 0xf; }
function lunarLeapDays(y) {
  return lunarLeapMonth(y) ? ((LUNAR_INFO[y - LUNAR_MIN] & 0x10000) ? 30 : 29) : 0;
}
function lunarMonthDays(y, m) {
  return (LUNAR_INFO[y - LUNAR_MIN] & (0x10000 >> m)) ? 30 : 29;
}
function lunarYearDays(y) {
  var sum = 348, info = LUNAR_INFO[y - LUNAR_MIN];
  for (var bit = 0x8000; bit > 0x8; bit >>= 1) { if (info & bit) sum += 1; }
  return sum + lunarLeapDays(y);
}

var LUNAR_YEAR_START_ = null;   // 每個農曆年的正月初一，距離基準日幾天
function lunarYearStarts_() {
  if (LUNAR_YEAR_START_) return LUNAR_YEAR_START_;
  var out = [], acc = 0;
  for (var y = LUNAR_MIN; y <= LUNAR_MAX; y++) { out.push(acc); acc += lunarYearDays(y); }
  LUNAR_YEAR_START_ = out;
  return out;
}

/** 一年裡的農曆月份依序排出來，閏月排在同名月份後面。 */
function lunarMonthsOf(y) {
  var leap = lunarLeapMonth(y), out = [];
  for (var m = 1; m <= 12; m++) {
    out.push({ m: m, leap: false, days: lunarMonthDays(y, m) });
    if (m === leap) out.push({ m: m, leap: true, days: lunarLeapDays(y) });
  }
  return out;
}

function ymdFromOffset_(off) {
  var d = new Date(LUNAR_BASE_UTC + off * 86400000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

/** 國曆 → 農曆。回傳 {y, m, d, leap}；超出 1900–2100 回傳 null。 */
function solarToLunar(Y, M, D) {
  var off = Math.round((Date.UTC(Y, M - 1, D) - LUNAR_BASE_UTC) / 86400000);
  if (off < 0) return null;
  var starts = lunarYearStarts_();
  var y = LUNAR_MIN;
  while (y < LUNAR_MAX && starts[y - LUNAR_MIN + 1] <= off) y++;
  var rest = off - starts[y - LUNAR_MIN];
  if (rest >= lunarYearDays(y)) return null;
  var months = lunarMonthsOf(y);
  for (var i = 0; i < months.length; i++) {
    if (rest < months[i].days) return { y: y, m: months[i].m, d: rest + 1, leap: months[i].leap };
    rest -= months[i].days;
  }
  return null;
}

/** 農曆 → 國曆。日子超過該月天數時，落在該月最後一天（例如三十遇到小月就是廿九）。 */
function lunarToSolar(y, m, d, leap) {
  if (y < LUNAR_MIN || y > LUNAR_MAX) return null;
  if (leap && lunarLeapMonth(y) !== m) return null;
  var off = lunarYearStarts_()[y - LUNAR_MIN];
  var months = lunarMonthsOf(y);
  for (var i = 0; i < months.length; i++) {
    var mo = months[i];
    if (mo.m === m && mo.leap === !!leap) {
      return ymdFromOffset_(off + Math.min(Math.max(1, d), mo.days) - 1);
    }
    off += mo.days;
  }
  return null;
}

var LUNAR_MONTH_NAMES = ['', '正月', '二月', '三月', '四月', '五月', '六月',
                         '七月', '八月', '九月', '十月', '十一月', '十二月'];
function lunarDayName(d) {
  var n = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
  if (d <= 10) return '初' + n[d];
  if (d < 20) return '十' + n[d - 10];
  if (d === 20) return '二十';
  if (d < 30) return '廿' + n[d - 20];
  return '三十';
}
function lunarText(l) {
  return l ? (l.leap ? '閏' : '') + LUNAR_MONTH_NAMES[l.m] + lunarDayName(l.d) : '';
}

/** 以 {y,m,d} 比大小 */
function ymdKey_(o) { return o.y * 10000 + o.m * 100 + o.d; }

/**
 * 從某天（含當天）起，下一個農曆每月 D 日。閏月也算一個月——初一十五拜拜，閏月照拜。
 * 回傳國曆 {y,m,d}。
 */
function nextLunarMonthly(D, from) {
  var cur = solarToLunar(from.y, from.m, from.d);
  if (!cur) return null;
  var fromKey = ymdKey_(from);
  for (var y = cur.y; y <= Math.min(cur.y + 2, LUNAR_MAX); y++) {
    var months = lunarMonthsOf(y);
    for (var i = 0; i < months.length; i++) {
      var s = lunarToSolar(y, months[i].m, D, months[i].leap);
      if (s && ymdKey_(s) >= fromKey) return s;
    }
  }
  return null;
}

/** 從某天（含當天）起，下一個農曆每年 M 月 D 日。只認正常月份，不在閏月重複一次。 */
function nextLunarYearly(M, D, from) {
  var cur = solarToLunar(from.y, from.m, from.d);
  if (!cur) return null;
  var fromKey = ymdKey_(from);
  for (var y = cur.y - 1; y <= Math.min(cur.y + 2, LUNAR_MAX); y++) {
    if (y < LUNAR_MIN) continue;
    var s = lunarToSolar(y, M, D, false);
    if (s && ymdKey_(s) >= fromKey) return s;
  }
  return null;
}
/* ================= 農曆結束 ================= */

