/**
 * 群組待辦板 — 後端
 *
 * 這支程式住在你的 Google 試算表裡，負責讀寫資料。
 * 資料就在同一份試算表的「任務」和「紀錄」兩個分頁，你隨時打開就看得到。
 * 照片存到你雲端硬碟的「群組待辦板照片」資料夾。
 *
 * ↓↓↓ 只有這一格需要你動 ↓↓↓
 */
var PASSCODE = '';   // 想加通行碼就填在引號中間，例如 '1234'。留空就是任何人有連結都能用。

/* ↑↑↑ 以下都不用改 ↑↑↑ */

var TRASH_DAYS   = 30;
var TASK_SHEET   = '任務';
var LOG_SHEET    = '紀錄';
var PHOTO_FOLDER = '群組待辦板照片';

var REPEAT_SHEET = '重複設定';

var TASK_HEADERS = ['編號','標題','分類','說明','狀態','建立者','建立時間',
                    '完工者','完工時間','完工備註','刪除者','刪除時間','照片','系統編號'];
var LOG_HEADERS  = ['時間','任務編號','任務標題','誰','動作','內容'];
var REPEAT_HEADERS = ['編號','標題','分類','說明','頻率','設定','參數','下次出現',
                      '建立者','建立時間','狀態','系統編號'];

var COL = { no:0, title:1, site:2, desc:3, status:4, createdBy:5, createdAt:6,
            doneBy:7, doneAt:8, doneNote:9, delBy:10, delAt:11, photos:12, id:13 };

var RCOL = { no:0, title:1, site:2, desc:3, freq:4, label:5, param:6, next:7,
             by:8, at:9, status:10, id:11 };

var FREQS = ['每天','每週','每月','每年'];

var OPEN = '待辦中', DONE = '已完工', TRASH = '刪除區';

/* ---------------- 入口 ---------------- */

function doGet(e) {
  return json_({ ok: true, hint: '待辦板後端正常運作中。請從網頁開啟。' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (PASSCODE && String(req.pass || '') !== String(PASSCODE)) {
      return json_({ ok: false, error: '通行碼不對', needPass: true });
    }
    var who = String(req.who || '').slice(0, 20) || '（未署名）';
    var out;

    switch (req.action) {
      case 'list':    out = listTasks_(); break;
      case 'add':     out = addTask_(req, who); break;
      case 'done':    out = finishTask_(req, who); break;
      case 'photos':  out = addPhotos_(req, who); break;
      case 'note':    out = addNote_(req, who); break;
      case 'reopen':  out = reopenTask_(req, who); break;
      case 'trash':   out = trashTask_(req, who); break;
      case 'restore': out = restoreTask_(req, who); break;
      case 'addRepeat': out = addRepeat_(req, who); break;
      case 'delRepeat': out = delRepeat_(req, who); break;
      default: throw new Error('不認得的動作：' + req.action);
    }
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 試算表 ---------------- */

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}
function tasksSheet_() {
  var sh = sheet_(TASK_SHEET, TASK_HEADERS);
  // 早期版本這欄叫「案場」，順手改成現在的名稱
  if (sh.getLastRow() >= 1 && sh.getRange(1, COL.site + 1).getValue() === '案場') {
    sh.getRange(1, COL.site + 1).setValue('分類');
  }
  return sh;
}
function logsSheet_()   { return sheet_(LOG_SHEET, LOG_HEADERS); }
function repeatSheet_() { return sheet_(REPEAT_SHEET, REPEAT_HEADERS); }

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function findRow_(sh, id) {
  var data = rows_(sh);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][COL.id]) === String(id)) return { row: i + 2, values: data[i] };
  }
  throw new Error('找不到這件待辦，可能已經被別人刪掉了');
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

function log_(no, title, who, action, text) {
  logsSheet_().appendRow([new Date(), no, title, who, action, text || '']);
}

function logsByNo_() {
  var out = {};
  rows_(logsSheet_()).forEach(function (r) {
    var no = String(r[1]);
    (out[no] = out[no] || []).push({ at: stamp_(r[0]), who: r[3], action: r[4], text: r[5] });
  });
  return out;
}

/* ---------------- 動作 ---------------- */

function listTasks_() {
  // 兩個人同時開板子時，不能重複產生或重複刪除
  lock_(function () { purge_(); generateDue_(); });
  var logs = logsByNo_();
  var tasks = rows_(tasksSheet_()).map(function (r) {
    var status = r[COL.status] === DONE ? 'done' : (r[COL.status] === TRASH ? 'trash' : 'open');
    return {
      id: String(r[COL.id]),
      no: Number(r[COL.no]) || 0,
      title: String(r[COL.title]),
      site: String(r[COL.site] || ''),
      desc: String(r[COL.desc] || ''),
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
      return {
        id: String(r[RCOL.id]),
        no: Number(r[RCOL.no]) || 0,
        title: String(r[RCOL.title]),
        site: String(r[RCOL.site] || ''),
        desc: String(r[RCOL.desc] || ''),
        freq: String(r[RCOL.freq]),
        label: String(r[RCOL.label] || repeatLabel_(r[RCOL.freq], r[RCOL.param])),
        next: stamp_(r[RCOL.next]),
        by: String(r[RCOL.by] || ''),
        status: String(r[RCOL.status] || '啟用')
      };
    });

  return { ok: true, tasks: tasks, repeats: repeats, trashDays: TRASH_DAYS };
}

function addTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var data = rows_(sh);
    var maxNo = data.reduce(function (m, r) { return Math.max(m, Number(r[COL.no]) || 0); }, 0);
    var no = maxNo + 1;
    var id = 't' + new Date().getTime() + Math.random().toString(36).slice(2, 6);
    var title = String(req.title || '').slice(0, 120);
    if (!title) throw new Error('先寫要做什麼');
    var site = String(req.site || '').slice(0, 60);

    sh.appendRow([no, title, site, String(req.desc || '').slice(0, 2000), OPEN,
                  who, new Date(), '', '', '', '', '', '', id]);
    log_(no, title, who, '建立', site ? '案場：' + site : '');
    return { ok: true, id: id };
  });
}

function finishTask_(req, who) {
  var ids = (req.photos || []).slice(0, 6).map(savePhoto_);
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, req.id);
    var note = String(req.note || '').slice(0, 1000);
    var kept = photoIds_(hit.values[COL.photos]).concat(ids);

    sh.getRange(hit.row, COL.status + 1).setValue(DONE);
    sh.getRange(hit.row, COL.doneBy + 1, 1, 3).setValues([[who, new Date(), note]]);
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
    var hit = findRow_(sh, req.id);
    var kept = photoIds_(hit.values[COL.photos]).concat(ids);
    sh.getRange(hit.row, COL.photos + 1).setValue(kept.join(','));
    log_(hit.values[COL.no], hit.values[COL.title], who, '加照片', '上傳 ' + ids.length + ' 張照片');
    return { ok: true };
  });
}

function addNote_(req, who) {
  var text = String(req.text || '').slice(0, 1000);
  if (!text) throw new Error('先寫點東西');
  return lock_(function () {
    var hit = findRow_(tasksSheet_(), req.id);
    log_(hit.values[COL.no], hit.values[COL.title], who, '留言', text);
    return { ok: true };
  });
}

function reopenTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, req.id);
    sh.getRange(hit.row, COL.status + 1).setValue(OPEN);
    sh.getRange(hit.row, COL.doneBy + 1, 1, 3).setValues([['', '', '']]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '退回待辦', '');
    return { ok: true };
  });
}

function trashTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, req.id);
    sh.getRange(hit.row, COL.status + 1).setValue(TRASH);
    sh.getRange(hit.row, COL.delBy + 1, 1, 2).setValues([[who, new Date()]]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '移到刪除區', '');
    return { ok: true };
  });
}

function restoreTask_(req, who) {
  return lock_(function () {
    var sh = tasksSheet_();
    var hit = findRow_(sh, req.id);
    var back = hit.values[COL.doneAt] ? DONE : OPEN;
    sh.getRange(hit.row, COL.status + 1).setValue(back);
    sh.getRange(hit.row, COL.delBy + 1, 1, 2).setValues([['', '']]);
    log_(hit.values[COL.no], hit.values[COL.title], who, '還原', '');
    return { ok: true };
  });
}

/* ---------------- 重複設定 ---------------- */

function dayStart_(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays_(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
function daysInMonth_(y, m) { return new Date(y, m + 1, 0).getDate(); }

/** 從 from 當天算起（含當天），下一次符合這個規則的日期。 */
function nextOccurrence_(freq, param, from) {
  var start = dayStart_(from);

  if (freq === '每天') return start;

  if (freq === '每週') {
    var target = Number(param) % 7;          // 1=週一 … 7=週日 → JS 的 1…6,0
    var d = new Date(start);
    for (var i = 0; i < 7; i++) {
      if (d.getDay() === target) return d;
      d = addDays_(d, 1);
    }
    return start;
  }

  if (freq === '每月') {
    var want = Math.min(31, Math.max(1, Number(param) || 1));
    var y = start.getFullYear(), m = start.getMonth();
    for (var k = 0; k < 24; k++) {
      // 31 號遇到只有 30 天的月份，就落在該月最後一天
      var cand = new Date(y, m, Math.min(want, daysInMonth_(y, m)));
      if (cand.getTime() >= start.getTime()) return cand;
      m++;
      if (m > 11) { m = 0; y++; }
    }
    return start;
  }

  if (freq === '每年') {
    var p = String(param).split('-');
    var mm = Math.min(12, Math.max(1, Number(p[0]) || 1)) - 1;
    var dd = Math.max(1, Number(p[1]) || 1);
    var yy = start.getFullYear();
    for (var j = 0; j < 4; j++) {
      var c = new Date(yy + j, mm, Math.min(dd, daysInMonth_(yy + j, mm)));
      if (c.getTime() >= start.getTime()) return c;
    }
  }

  return start;
}

function repeatLabel_(freq, param) {
  var W = ['', '週一', '週二', '週三', '週四', '週五', '週六', '週日'];
  if (freq === '每天') return '每天';
  if (freq === '每週') return '每' + (W[Number(param)] || '週一');
  if (freq === '每月') return '每月 ' + (Number(param) || 1) + ' 號';
  if (freq === '每年') {
    var p = String(param).split('-');
    return '每年 ' + (Number(p[0]) || 1) + ' 月 ' + (Number(p[1]) || 1) + ' 日';
  }
  return String(freq);
}

function addRepeat_(req, who) {
  return lock_(function () {
    var title = String(req.title || '').slice(0, 120);
    if (!title) throw new Error('先寫要做什麼');
    var freq = String(req.freq || '');
    if (FREQS.indexOf(freq) < 0) throw new Error('頻率設定不正確');

    var param = String(req.param || '');
    var sh = repeatSheet_();
    var no = rows_(sh).reduce(function (m, r) { return Math.max(m, Number(r[RCOL.no]) || 0); }, 0) + 1;
    var id = 'r' + new Date().getTime() + Math.random().toString(36).slice(2, 6);

    sh.appendRow([no, title, String(req.site || '').slice(0, 60),
                  String(req.desc || '').slice(0, 2000),
                  freq, repeatLabel_(freq, param), param,
                  nextOccurrence_(freq, param, new Date()),
                  who, new Date(), '啟用', id]);
    return { ok: true, id: id };
  });
}

function delRepeat_(req, who) {
  return lock_(function () {
    var sh = repeatSheet_();
    var data = rows_(sh);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][RCOL.id]) === String(req.id)) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    throw new Error('找不到這項重複設定');
  });
}

/**
 * 到期的重複設定會自動長出一件待辦。
 * 若中間隔了好幾期沒人開板子，只補最近的一件，不會一次灌一堆進來。
 */
function generateDue_() {
  var sh = repeatSheet_();
  var data = rows_(sh);
  if (!data.length) return;
  var today = dayStart_(new Date());

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[RCOL.title] || r[RCOL.status] === '停用') continue;

    var next = r[RCOL.next] ? dayStart_(r[RCOL.next]) : null;
    if (!next) {
      sh.getRange(i + 2, RCOL.next + 1).setValue(nextOccurrence_(r[RCOL.freq], r[RCOL.param], today));
      continue;
    }
    if (next.getTime() > today.getTime()) continue;

    createFromRepeat_(r);
    sh.getRange(i + 2, RCOL.next + 1)
      .setValue(nextOccurrence_(r[RCOL.freq], r[RCOL.param], addDays_(today, 1)));
  }
}

function createFromRepeat_(r) {
  var sh = tasksSheet_();
  var no = rows_(sh).reduce(function (m, x) { return Math.max(m, Number(x[COL.no]) || 0); }, 0) + 1;
  var id = 't' + new Date().getTime() + Math.random().toString(36).slice(2, 6);
  sh.appendRow([no, r[RCOL.title], r[RCOL.site], r[RCOL.desc], OPEN,
                r[RCOL.by], new Date(), '', '', '', '', '', '', id]);
  log_(no, r[RCOL.title], r[RCOL.by], '自動建立', '來自重複設定：' + r[RCOL.label]);
}

/**
 * 刪除區裡超過 30 天的，連照片帶紀錄一起永久刪除。
 * 每次載入板子時順手檢查，不需要另外設定排程。
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
