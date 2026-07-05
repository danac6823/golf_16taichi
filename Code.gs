/**
 * 高爾夫球隊例賽報名 — Google Apps Script 後端
 * 搭配 LIFF 前端網頁使用。所有資料存在這份試算表。
 *
 * 部署方式:部署 → 新增部署作業 → 類型「網頁應用程式」
 *   執行身分:我
 *   存取權:任何人
 * 取得的 /exec 網址,貼到前端 index.html 的 GAS_URL。
 *
 * 用 JSONP 回傳,前端用 <script> 載入,完全避開 CORS。
 */

// ===== LINE 身分驗證(只驗寫入動作)=====
// 填入 LINE Login 頻道密鑰後即「啟用寫入需驗證」;留空 = 不啟用(維持原行為)。
// 需先到 LINE Developers 後台,為此 LIFF 所屬頻道開啟 OpenID Connect(openid scope)。
var LINE_CHANNEL_ID     = '';   // LINE Login 頻道 ID(數字),選填;留空則不檢查 aud
var LINE_CHANNEL_SECRET = '';   // LINE Login 頻道密鑰;填入後啟用驗證

// ---- 工作表名稱 ----
var SHEET_ROSTER  = 'Roster';     // userId | name | hcp | cart | note | ts
var SHEET_MEMBERS = 'Members';    // userId | name | role(會員/來賓) | gender(男/女) | birthYear | phone  (LINE 綁定)
var SHEET_MASTER  = '會員名單';   // 姓名 | 性別 | 年初差點 | 出生年  (主檔,管理員維護)
var SHEET_CONFIG  = 'Config';     // key | value
var SHEET_HCP     = 'Handicaps';  // name | hcp         (跟著姓名走的差點)
var SHEET_SCORES  = 'Scores';     // date | name | out | in | gross | hcp | net | rankType | rank | hcpAfter | ts
var SHEET_PAY     = 'Payments';   // name | paid        (本年度會費是否已收)
var SHEET_SCHED   = '行事曆';      // 年|月|日期|星期|隊伍|類型|球場|開球時間|預訂狀態|可預約日|預約提醒|提前天數|提前月數|組數|擊球價
var SHEET_TECH    = '技術獎';      // 日期 | 獎項 | 得獎者 | 條數  (送球以條計,1盒=4條)
var SHEET_ROSTER_ARCHIVE = '報名紀錄'; // 場次 | 姓名 | 狀態 | 用餐 | 差點 | 備註 | userId | 報名時間  (每月報名封存)

// ---- 活動資訊預設值 ----
var DEFAULT_CONFIG = {
  title:    '意象太極',
  titleZh:  '○○高爾夫球隊 · 月例賽',
  date:     '6/28(六)',
  course:   '○○高爾夫俱樂部',
  tee:      '07:30(報到 07:00)',
  fee:      'NT$2,800',
  deadline: '6/25(三)18:00',
  cap:      '32',
  par:      '72',        // 標準桿
  hcpK:     '0.2',       // 差點更新平滑係數
  newHcpFactor: '0.9',   // 新會員第二場建立差點的倍率:((第一場+第二場)/2 − 標準桿)× 此值
  maxHcp:   '36',        // 差點上限(所有計算結果不超過此值)
  feeFemale: '10000',    // 女會員年費
  feeMale:   '11000',    // 男會員年費
  noticeEvent: '',       // 本場注意事項(報名卡下方;換月自動清空)
  announce:  '',         // 最新消息 / 外地賽(首頁上方常駐)
  bankInfo:  '',         // 匯款資訊(我的會費區,只對未繳者顯示)
  eventGraceDays: '3',   // 賽後仍把「本場」當作目前場次顯示的天數
  adminUserId: ''
};

// LINE ID Token 本地驗章(HS256,用頻道密鑰);驗過回傳 payload(payload.sub = userId),否則 null
function verifyIdToken_(idToken) {
  if (!idToken || !LINE_CHANNEL_SECRET) return null;
  var parts = String(idToken).split('.');
  if (parts.length !== 3) return null;
  try {
    var signingInput = parts[0] + '.' + parts[1];
    var sigBytes = Utilities.computeHmacSha256Signature(signingInput, LINE_CHANNEL_SECRET);
    var expected = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
    if (expected !== parts[2]) return null;                                    // 簽章不符
    var json = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString('UTF-8');
    var payload = JSON.parse(json);
    if (payload.exp && (Number(payload.exp) * 1000) < Date.now()) return null;  // 已過期
    if (payload.iss && payload.iss !== 'https://access.line.me') return null;   // 來源不符
    if (LINE_CHANNEL_ID && payload.aud && String(payload.aud) !== String(LINE_CHANNEL_ID)) return null; // aud 不符
    return payload;
  } catch (err) {
    return null;
  }
}

// ===== 自動備份(複製整份試算表到雲端硬碟)=====
var BACKUP_FOLDER = '例賽系統備份';
var BACKUP_KEEP = 30;          // 最多保留份數,超過自動清最舊

function backupSpreadsheet_(tag) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var folder = backupFolder_();
    var name = ss.getName() + '｜備份 ' + backupStamp_(new Date()) + (tag ? ('｜' + tag) : '');
    file.makeCopy(name, folder);
    pruneBackups_(folder);
    return true;
  } catch (e) { return false; }   // 備份失敗絕不影響主流程
}
function backupFolder_() {
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
}
function backupStamp_(d) {
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
}
function pruneBackups_(folder) {
  var files = [], it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  if (files.length <= BACKUP_KEEP) return;
  files.sort(function (a, b) { return b.getDateCreated().getTime() - a.getDateCreated().getTime(); });
  for (var i = BACKUP_KEEP; i < files.length; i++) { try { files[i].setTrashed(true); } catch (e) {} }
}
// 立即備份(可在 Apps Script 編輯器手動執行)
function backupNow() { return backupSpreadsheet_('手動') ? '已備份一份到「' + BACKUP_FOLDER + '」' : '備份失敗'; }
// 在 Apps Script 編輯器執行一次,建立「每天 09:00 檢查」觸發器(月例賽隔天自動備份)
// 保溫:每隔幾分鐘輕觸一次 試算表,降低 GAS 冷啟動機率(到 Apps Script 執行一次 setupKeepWarm 即可)
function keepWarm() { try { SpreadsheetApp.getActiveSpreadsheet().getName(); } catch (e) {} }
function setupKeepWarm() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'keepWarm') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(5).create();
  return '已設定:每 5 分鐘保溫一次(降低冷啟動)';
}
function setupBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyBackupCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackupCheck').timeBased().everyDays(1).atHour(9).create();
  return '已設定:每天 09:00 檢查,月例賽隔天自動備份';
}
function dailyBackupCheck() {
  var y = new Date(); y.setDate(y.getDate() - 1);                // 昨天
  var key = y.getFullYear() * 10000 + (y.getMonth() + 1) * 100 + y.getDate();
  var TYPES = { '月例賽': 1, '外地賽': 1, '國外賽': 1, '交接': 1 };   // 這些類型的隔天要備份
  var sched = getSchedule(), ev = null;
  sched.forEach(function (e) {
    var t = String(e.type || '');
    var hit = (t.indexOf('賽') >= 0) || (t.indexOf('交接') >= 0);   // 例賽/月例賽/外地賽/國外賽/交接
    if (!ev && e.sortKey === key && hit) ev = e;
  });
  if (!ev) return;
  var now = new Date();
  var thisMonth = now.getFullYear() + '/' + (now.getMonth() + 1);
  if (getConfig().lastBackupMonth === thisMonth) return;        // 本月已備份過 → 不重複(一個月一次)
  if (backupSpreadsheet_(ev.type + '後')) setConfig('lastBackupMonth', thisMonth);
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback || 'callback';
  var out;
  try {
    out = handle(p);
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handle(p) {
  __cfgCache = null;   // 每次請求開始先清快取,避免暖啟動殘留舊設定
  __memInfoCache = __hcpMapCache = __masterListCache = __masterMapCache = null;
  __schedCache = __rosterCache = __membersDataCache = null;
  var action = p.action || 'bootstrap';
  var userId = p.userId || '';

  // 只驗「寫入動作」與「管理員資料」:啟用驗證(已填頻道密鑰)時,
  // 用 LINE ID Token 本地驗章取得真實 userId,忽略前端帶來的 userId,防止假冒。
  var VERIFY = {
    setName:1, register:1, requestLeave:1, addGuest:1, removeGuest:1, cancel:1, claimAdmin:1,
    proxyRegister:1, proxyRemove:1,
    updateEvent:1, clearRoster:1, makeGroups:1, clearGroups:1, adjustGroups:1, importHcp:1, submitScores:1,
    submitMatch:1, seasonReset:1, buildSeason:1, buildSeasonMatrix:1, setPaid:1, ensurePayList:1, buildSchedule:1,
    importMembers:1, recordTech:1, buildBalls:1, clearTechDate:1, exportPdf:1,
    seedTestData:1, clearTestData:1, adminData:1, lockScores:1, unlockScores:1, setNote:1, clearMatch:1, flushCache:1, syncRosterNames:1
  };
  if (LINE_CHANNEL_SECRET && VERIFY[action]) {
    var vp = verifyIdToken_(p.idToken);
    if (!vp || !vp.sub) return { ok: false, error: '身分驗證失敗,請重新開啟頁面再試' };
    userId = String(vp.sub);   // 用驗證過的真實 userId
  }

  switch (action) {
    case 'bootstrap':   return bootstrap(userId, p);
    case 'setName':     return setName(userId, p);
    case 'register':    return register(userId, p);
    case 'requestLeave': return requestLeave(userId, p);
    case 'addGuest':    return addGuest(userId, p);
    case 'removeGuest': return removeGuest(userId, p);
    case 'proxyRegister': return proxyRegister(userId, p);
    case 'proxyRemove':   return proxyRemove(userId, p);
    case 'cancel':      return cancel(userId);
    case 'updateEvent': return updateEvent(userId, p);
    case 'clearRoster': return clearRoster(userId);
    case 'claimAdmin':  return claimAdmin(userId);
    case 'makeGroups':  return makeGroups(userId, p);
    case 'clearGroups': return clearGroups(userId);
    case 'adjustGroups': return adjustGroups(userId, p);
    case 'importHcp':   return importHcp(userId, p);
    case 'submitScores':return submitScores(userId, p);
    case 'submitMatch': return submitMatch(userId, p);
    case 'lockScores': return lockScoresAction(userId, p);
    case 'unlockScores': return unlockScoresAction(userId, p);
    case 'seasonReset': return seasonReset(userId, p);
    case 'seasonResults': return seasonData(userId, p);
    case 'myStats':     return myStatsAction(userId, p);
    case 'adminData': return adminData(userId, p);
    case 'buildSeason': return buildSeason(userId);
    case 'buildSeasonMatrix': return buildSeasonMatrix(userId);
    case 'scoreMatrix': return { ok: true, matrix: buildScoreMatrix_(getConfig()) };
    case 'setPaid':     return setPaidAction(userId, p);
    case 'ensurePayList': return ensurePayList(userId);
    case 'buildSchedule': return buildSchedule(userId);
    case 'importMembers': return importMembers(userId, p);
    case 'recordTech':  return recordTech(userId, p);
    case 'buildBalls':  return buildBalls(userId);
    case 'clearTechDate': return clearTechDate(userId, p);
    case 'exportPdf':   return exportPdf(userId, p);
    case 'clearMatch':  return clearMatch(userId, p);
    case 'flushCache':  return flushCacheAction(userId, p);
    case 'syncRosterNames': return syncRosterNames(userId, p);
    case 'setNote':     return setNote(userId, p);
    case 'seedTestData': return seedTestData(userId);
    case 'clearTestData': return clearTestData(userId);
    default:            return { ok: false, error: 'unknown action' };
  }
}

// ---------- 主要動作 ----------

// 報名清單套用「管理中的差點」(與 bootstrap 顯示一致),供 bootstrap 與寫入動作共用
function rosterWithHcp_() {
  var hmap = getHcpMap();
  return getRoster().map(function (r) {
    if (hmap.hasOwnProperty(r.name)) r.hcp = hmap[r.name];
    return r;
  });
}

function bootstrap(userId, p) {
  var cfg = getConfig();
  rolloverRosterIfSwitched_(cfg);                 // 換場自動封存上一場 + 清空(常態只比對不寫入)
  var hmap = getHcpMap();
  var roster = rosterWithHcp_();
  var member = getMember(userId);
  var myHcp = (member && hmap.hasOwnProperty(member.name)) ? hmap[member.name] : '';
  var realAdmin = !!userId && userId === cfg.adminUserId;
  var preview = realAdmin && p && String(p.asMember) === '1';   // 管理員預覽會員檢視
  var isAdmin = realAdmin && !preview;
  var groups = null, ranking = null;
  if (cfg.groups) { try { groups = JSON.parse(cfg.groups); } catch (e) {} }
  if (cfg.lastRanking) { try { ranking = JSON.parse(cfg.lastRanking); } catch (e) {} }
  var latestDate = ranking ? ranking.date : null;   // 最近一場日期(時間鎖會把 ranking 設為 null,先記下)

  // 時間鎖:分組當天 08:00、成績當天 18:00 才開放給隊員(管理員不受限)
  var rev = revealState(cfg);
  var groupsLocked = null, scoreLocked = null;
  if (!isAdmin) {
    if (!rev.groupsOpen && groups) { groups = null; groupsLocked = rev.groupLabel; }
    if (!rev.scoreOpen && ranking) { ranking = null; scoreLocked = rev.scoreLabel; }
  }

  var res = {
    ok: true,
    event: cfg,
    roster: roster,
    member: member,                              // {name, role, gender, birthYear} 或 null
    myHcp: String(myHcp),                        // 我目前的差點(系統管理)
    myFee: String(feeOf(member)),                // 我的會費(來賓為空)
    groups: groups,                              // 分組結果或 null
    groupsLocked: groupsLocked,                  // 尚未到開放時間時的提示文字
    ranking: ranking,                            // 最近一場排名或 null
    scoreLocked: scoreLocked,                    // 尚未到開放時間時的提示文字
    isAdmin: isAdmin,
    realAdmin: realAdmin,                        // 真實身份是否為管理員(供前端顯示切換鈕)
    preview: preview,                            // 目前是否處於會員預覽模式
    hasAdmin: !!cfg.adminUserId
  };
  // 管理員資料(差點/會費/綁定/球數/技術獎)改為「展開管理員專區時才抓」(adminData action),不在首屏 bootstrap 一次讀完,加快登入
  res.schedule = getSchedule();                   // 賽程(所有人可看)
  res.eventView = effectiveEvent(cfg, res.schedule); // 報名頁顯示用(自動帶下一場)
  res.notices = { announce: cfg.announce || '', event: cfg.noticeEvent || '', bank: cfg.bankInfo || '' };
  res.myPaid = true;                              // 匯款區只給未繳會員 → 有匯款資訊且為會員時才查繳費狀態
  if (res.notices.bank && member && member.role === '會員') res.myPaid = !!getPaidMap()[member.name];
  var __mInfo = getMemberInfoMap();               // 一次讀會員資料,同時產生 性別/身分 對照
  res.genderMap = (function(){ var g={}; for(var k in __mInfo){ var gv=__mInfo[k].gender; if(gv==='男'||gv==='女') g[k]=gv; } return g; })(); // 姓名→性別(前端上色用)
  res.roleMap = (function(){ var r={}; for(var k in __mInfo){ r[k]=__mInfo[k].role; } return r; })();   // 姓名→身分(會員/來賓)
  // 我的本年度出席/獎金改為「首屏畫好後再補抓」(myStats action),不在 bootstrap 掃整年 Scores,加快登入
  res.progress = getReportProgress();             // 回報進度(年度會員/已登記/未登記;用快取,無額外讀取)
  if (isAdmin && latestDate) {                      // 管理員:最近一場是否已鎖定(供鎖定/解鎖鈕)
    res.matchLockDate = latestDate;
    res.matchLocked = isDateLocked_(cfg, latestDate);
  }
  // 年度各場成績改為「進年度成績時才抓」(seasonResults action),不在首屏 bootstrap 掃整年,加快登入
  var nowK = new Date();
  res.todayKey = nowK.getFullYear() * 10000 + (nowK.getMonth() + 1) * 100 + nowK.getDate();
  // 該場技術獎得獎名單(併入成績排名顯示;成績未開放時 res.ranking 為 null,自然看不到)
  if (res.ranking && res.ranking.date) {
    res.ranking.tech = getTechList(cfg)
      .filter(function (t) { return t.date === res.ranking.date && t.winner; })
      .map(function (t) { return { award: t.award, winner: t.winner, tiao: t.tiao }; });
  }
  return res;
}

// 年度各場成績(依日期);每場含全員 + 該場總桿冠軍
// 年度成績(延後載入用):沿用 bootstrap 的「未開放最新一場對會員隱藏」過濾
// 管理員專區資料(延後載入用):展開管理員區塊時才抓
function adminData(userId, p) {
  var cfg = getConfig();
  var realAdmin = !!userId && userId === cfg.adminUserId;
  var preview = realAdmin && p && String(p.asMember) === '1';
  if (!realAdmin || preview) return { ok: false, error: '非管理員' };
  var gate = seasonResetGate_(cfg);
  return {
    ok: true,
    hcpList: getHcpList(),
    memberFees: getMemberFees(),
    binding: getBindingStatus(),
    balls: getBallStat(cfg),
    techList: getTechList(cfg),
    seasonResetAllowed: gate.allowed,
    seasonResetLastMatch: gate.lastMatch,
    teamNote: String(cfg['球隊小提醒'] == null ? '' : cfg['球隊小提醒'])
  };
}

// 設定球隊小提醒(管理員)→ 寫入 Config「球隊小提醒」,匯出 PDF 時帶入
function setNote(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以設定' };
  setConfig('球隊小提醒', String(p && p.text != null ? p.text : ''));
  return { ok: true };
}

// 同步出席名單姓名:依 userId 把名單中的名字更新成 Members 最新姓名(順便更新差點),修正舊資料用
function syncRosterNames(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以同步名單' };
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  var changed = 0;
  try {
    var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, changed: 0, roster: rosterWithHcp_(), progress: getReportProgress() };
    var mdata = membersData_(), nameByUid = {};
    for (var i = 1; i < mdata.length; i++) { var u = String(mdata[i][0] || ''); if (u) nameByUid[u] = String(mdata[i][1] || ''); }
    var hmap = getHcpMap();
    var out = [];   // 第2、3欄(name, hcp)
    for (var r = 1; r < data.length; r++) {
      var uid = String(data[r][0] || ''), curName = String(data[r][1] || ''), curHcp = String(data[r][2] || '');
      var want = nameByUid[uid];
      if (uid && want && want !== curName) {
        var h = hmap.hasOwnProperty(want) ? String(hmap[want]) : '';
        out.push([want, h]); changed++;
      } else {
        out.push([curName, curHcp]);
      }
    }
    if (changed > 0) sh.getRange(2, 2, out.length, 2).setValues(out);   // 一次寫回
  } finally { lock.releaseLock(); }
  __rosterCache = null; cacheDel_('roster');
  return { ok: true, changed: changed, roster: rosterWithHcp_(), progress: getReportProgress() };
}

// 清除共享暖快取(手動改試算表後用;管理員)
function flushCacheAction(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以清除快取' };
  flushWarmCache();
  return { ok: true };
}

// 我的本年度出席/獎金(延後抓,首屏不掃 Scores)
function myStatsAction(userId, p) {
  var cfg = getConfig();
  var member = getMember(userId);
  return { ok: true, myStats: member ? getMyStats(member.name, cfg) : null };
}

function seasonData(userId, p) {  var cfg = getConfig();
  var realAdmin = !!userId && userId === cfg.adminUserId;
  var preview = realAdmin && p && String(p.asMember) === '1';
  var isAdmin = realAdmin && !preview;
  var ranking = null;
  if (cfg.lastRanking) { try { ranking = JSON.parse(cfg.lastRanking); } catch (e) {} }
  var latestDate = ranking ? ranking.date : null;
  var rev = revealState(cfg);
  var scoreLocked = (!rev.scoreOpen && ranking) ? true : false;
  var results = getSeasonResults(cfg);
  if (!isAdmin && scoreLocked && latestDate) {
    results = results.filter(function (m) { return m.date !== latestDate; });
  }
  return { ok: true, seasonResults: results };
}

function getSeasonResults(cfg) {
  var range = seasonRange(cfg);
  var data = sheet(SHEET_SCORES).getDataRange().getValues();
  var byDate = {}, order = {}, dates = [];
  for (var i = 1; i < data.length; i++) {
    var d = mdOf_(data[i][0]); if (!d) continue;
    if (range && !inCurrentSeason(d, range)) continue;
    if (!(d in byDate)) { byDate[d] = []; order[d] = range ? scoreFullDate(d, range).getTime() : i; dates.push(d); }
    byDate[d].push({
      name: String(data[i][1] || ''),
      gross: Number(data[i][4]) || 0,
      net: Number(data[i][6]) || 0,
      rankType: String(data[i][7] || ''),
      rank: String(data[i][8] || '')
    });
  }
  dates.sort(function (a, b) { return order[a] - order[b]; });
  var links = pdfLinksObj_(cfg);
  return dates.map(function (d) {
    var rows = byDate[d], gross = '', grossScore = 0;
    rows.forEach(function (r) { if (r.rankType === '總桿') { gross = r.name; grossScore = r.gross; } });
    return { date: d, gross: gross, grossScore: grossScore, rows: rows, pdf: links[d] || '' };
  });
}

// 回報進度:年度會員人數、已登記(報名+請假)、未登記
function getReportProgress() {
  var master = getMasterList();                   // 會員名單主檔
  var memberNames = {}; master.forEach(function (m) { memberNames[m.name] = true; });
  var minfo = getMemberInfoMap();
  var roster = getRoster();
  var reg = 0, leave = 0, guests = 0, respondedMembers = {};
  roster.forEach(function (r) {
    var isMember = memberNames[r.name] || (minfo[r.name] && minfo[r.name].role !== '來賓');
    if (r.status === '請假') leave++; else reg++;
    if (isMember) respondedMembers[r.name] = true; else guests++;
  });
  var responded = Object.keys(respondedMembers).length;
  return {
    members: master.length,                       // 年度會員人數
    responded: responded,                         // 已登記會員(報名+請假)
    notYet: Math.max(0, master.length - responded),// 尚未登記
    reg: reg, leave: leave, guests: guests
  };
}

function setName(userId, p) {
  if (!userId) return { ok: false, error: 'no userId' };
  var name = (p.name || '').toString().trim();
  if (!name) return { ok: false, error: '請輸入姓名' };

  var master = getMasterMap();
  var hit = master[name];                          // 主檔有此姓名?
  var role, gender, birth;
  var phone = (p.phone || '').toString().trim();
  if (hit) {
    // 對到主檔 → 綁定為會員,性別/出生年用主檔(主檔沒填則用輸入)
    role = '會員';
    gender = hit.gender || ((p.gender === '女' || p.gender === '男') ? p.gender : '');
    birth = hit.birthYear || (p.birthYear || '').toString().trim();
  } else {
    // 沒對到 → 以來賓加入,性別/出生年用輸入
    role = '來賓';
    gender = (p.gender === '女') ? '女' : (p.gender === '男' ? '男' : '');
    birth = (p.birthYear || '').toString().trim();
  }
  var lineName = (p.lineName || '').toString().trim();
  upsertMember(userId, name, role, gender, birth, phone, lineName);
  if (hit) mergeProxyRoster_(userId, name);     // 對到會員 → 把之前的代報名接到本人
  return { ok: true, member: getMember(userId), matched: !!hit };
}

function register(userId, p) { return upsertRoster(userId, p, '報名'); }
function requestLeave(userId, p) { return upsertRoster(userId, p || {}, '請假'); }

// 會員幫邀請的來賓登記參加(會員本人需已報名)
function addGuest(userId, p) {
  if (!userId) return { ok: false, error: 'no userId' };
  var inviter = getMember(userId);
  if (!inviter) return { ok: false, error: '尚未設定姓名' };
  var name = String(p.guestName || '').trim();
  if (!name) return { ok: false, error: '請填來賓姓名' };
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    var gid = 'G:' + userId + ':' + name;
    var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues(), rowIndex = -1;
    for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === gid) { rowIndex = i + 1; break; } }
    var cart = (p.cart || '用餐').toString().trim();
    var note = '來賓・' + inviter.name + ' 邀請';
    var row = [gid, name, '', cart, note, Date.now(), '報名'];
    if (rowIndex === -1) sh.appendRow(row);
    else sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    upsertMember(gid, name, '來賓', '', '');     // 標記為來賓(計分/名單用)
  } finally { lock.releaseLock(); }
  __rosterCache = null; cacheDel_('roster'); return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

// ---------- 代會員報名(幫還沒綁定 LINE 的會員登記出席;身分=會員,非來賓)----------
// 用 userId 前綴 'M:姓名' 標記;故意不寫 Members(計分自動當會員、避免日後綁定產生同名兩列)
function proxyRegister(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以代報名' };
  var name = String(p.name || '').trim();
  if (!name) return { ok: false, error: '請選擇會員' };
  var inMaster = getMasterList().some(function (m) { return m.name === name; });
  if (!inMaster) return { ok: false, error: '「' + name + '」不在會員名單' };
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || '') === name) return { ok: false, error: '「' + name + '」已在出席名單' };
    }
    var hmap = getHcpMap();
    var hcp = hmap.hasOwnProperty(name) ? String(hmap[name]) : '';
    var cart = (p.cart || '用餐').toString().trim();
    sh.appendRow(['M:' + name, name, hcp, cart, '代報名', Date.now(), '報名']);
  } finally { lock.releaseLock(); }
  __rosterCache = null; cacheDel_('roster');
  return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

function proxyRemove(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以移除代報名' };
  var name = String(p.name || '').trim();
  var target = 'M:' + name;
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues(), row = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '') === target) { row = i + 1; break; }   // 只允許移除 M: 開頭
    }
    if (row === -1) return { ok: false, error: '找不到此代報名' };
    sh.deleteRow(row);
  } finally { lock.releaseLock(); }
  __rosterCache = null; cacheDel_('roster');
  return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

// 綁定時把「代報名(M:姓名)」那列接到本人帳號
function mergeProxyRoster_(realUserId, name) {
  var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues();
  var proxyRow = -1, hasReal = false;
  for (var i = 1; i < data.length; i++) {
    var uid = String(data[i][0] || '');
    if (uid === 'M:' + name) proxyRow = i + 1;
    else if (uid === realUserId) hasReal = true;
  }
  if (proxyRow === -1) return;
  if (!hasReal) {
    sh.getRange(proxyRow, 1).setValue(realUserId);   // 換成本人 userId
    sh.getRange(proxyRow, 5).setValue('');           // 清除「代報名」註記
  } else {
    sh.deleteRow(proxyRow);                          // 本人已有列 → 刪代報名列避免重複
  }
  __rosterCache = null; cacheDel_('roster');
}

// 移除會員自己登記的來賓
function removeGuest(userId, p) {
  if (!userId) return { ok: false, error: 'no userId' };
  var name = String(p.guestName || '').trim();
  var gid = 'G:' + userId + ':' + name;
  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === gid) sh.deleteRow(i + 1); }
    var ms = sheet(SHEET_MEMBERS), md = ms.getDataRange().getValues();
    for (var j = md.length - 1; j >= 1; j--) { if (String(md[j][0]) === gid) ms.deleteRow(j + 1); }
    __memInfoCache = __membersDataCache = null; cacheDel_('members');   // 已刪會員列 → 清快取
  } finally { lock.releaseLock(); }
  __rosterCache = null; cacheDel_('roster'); return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

function upsertRoster(userId, p, status) {
  if (!userId) return { ok: false, error: 'no userId' };
  var member = getMember(userId);
  if (!member) return { ok: false, error: '尚未設定姓名' };

  var lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER);
    var data = sh.getDataRange().getValues();        // 含表頭
    var rowIndex = -1, prev = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === userId) { rowIndex = i + 1; prev = data[i]; break; }
    }
    var hmap = getHcpMap();
    var hcp = hmap.hasOwnProperty(member.name) ? String(hmap[member.name]) : '';
    // 請假 → 自動帶「不用餐」;報名 → 用傳入值
    var cart = (status === '請假')
      ? '不用餐'
      : (p.cart || '用餐').toString().trim();
    var note = (p.note !== undefined && p.note !== null && String(p.note) !== '')
      ? String(p.note).trim()
      : (prev ? String(prev[4] || '') : '');
    var row = [userId, member.name, hcp, cart, note, Date.now(), status];
    if (rowIndex === -1) sh.appendRow(row);
    else sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  __rosterCache = null; cacheDel_('roster'); return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

function cancel(userId) {
  if (!userId) return { ok: false, error: 'no userId' };
  var lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER);
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === userId) { sh.deleteRow(i + 1); }
    }
  } finally {
    lock.releaseLock();
  }
  __rosterCache = null; cacheDel_('roster'); return { ok: true, roster: rosterWithHcp_(), progress: getReportProgress() };
}

function updateEvent(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以修改活動資訊' };
  }
  var fields = ['title','titleZh','date','course','tee','fee','deadline','cap','par','matchDate','groupRevealAt','scoreRevealAt','teamName','term','president','vicePresident','secretary','treasurer','seasonFrom','seasonTo','prizeGross','prizeNet1','prizeNet2','prizeNet3','prizeLucky','luckyPlace','prizeSkip5','skipStep','prizeBB','ballInitBoxes','ballInitNote','noticeEvent','announce','bankInfo'];
  fields.forEach(function (k) {
    if (p[k] !== undefined) setConfig(k, p[k]);
  });
  return { ok: true, event: getConfig() };
}

function clearRoster(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以清空名單' };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    var sh = sheet(SHEET_ROSTER);
    var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  } finally {
    lock.releaseLock();
  }
  __rosterCache = null; cacheDel_('roster'); return { ok: true, roster: [] };
}

function claimAdmin(userId) {
  if (!userId) return { ok: false, error: 'no userId' };
  var cfg = getConfig();
  if (cfg.adminUserId) {
    return { ok: false, error: '已有管理員', isAdmin: userId === cfg.adminUserId };
  }
  setConfig('adminUserId', userId);
  return { ok: true, isAdmin: true };
}

// ---------- 自動分組 ----------

function makeGroups(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以分組' };
  }
  var size = parseInt(p.size, 10); if (size !== 3 && size !== 4) size = 4;
  var method = p.method || 'order';
  var start = parseTime((p.start || '').trim());
  var interval = parseInt(p.interval, 10); if (isNaN(interval) || interval < 0) interval = 0;

  // 只排「確定參加」的人:排除請假;候補(超過名額)也不納入分組
  var roster = getRoster().filter(function (r) { return r.status !== '請假'; });
  var minfo = getMemberInfoMap();
  var cap = parseInt(cfg.cap, 10);
  if (!isNaN(cap) && cap > 0 && roster.length > cap) roster = roster.slice(0, cap);
  var players = roster.map(function (r) {
    return { name: r.name, hcp: parseFloat(r.hcp), cart: r.cart, userId: r.userId, note: r.note,
             gender: (minfo[r.name] ? minfo[r.name].gender : '') };
  });
  var n = players.length;
  if (n === 0) return { ok: false, error: '名單目前是空的,無法分組' };

  // 解析「備註(同組)」:把互相點名的人綁成同一群(第一順位)
  var cl = buildSameGroupClusters(players);

  var numGroups = Math.max(1, Math.ceil(n / size));
  var groups = [];
  for (var g = 0; g < numGroups; g++) groups.push([]);

  if (cl.multi.length === 0) {
    // 沒有同組需求 → 依方法分配
    if (method === 'gender') {
      // 二男二女:男生先平均分到各組,女生再平均分到各組
      players.filter(function (p) { return p.gender === '男'; })
             .forEach(function (p) { placeByGenderBalance(p, groups, size); });
      players.filter(function (p) { return p.gender === '女'; })
             .forEach(function (p) { placeByGenderBalance(p, groups, size); });
      players.filter(function (p) { return p.gender !== '男' && p.gender !== '女'; })
             .forEach(function (p) { groups[emptiestGroupIdx(groups, size)].push(p); });
      groups = groups.filter(function (gp) { return gp.length > 0; });
    } else {
      if (method === 'random') {
        for (var i = n - 1; i > 0; i--) {
          var j = Math.floor(Math.random() * (i + 1));
          var t = players[i]; players[i] = players[j]; players[j] = t;
        }
      } else if (method === 'hcp') {
        players.sort(function (a, b) {
          var ha = isNaN(a.hcp) ? 9999 : a.hcp, hb = isNaN(b.hcp) ? 9999 : b.hcp;
          return ha - hb;
        });
      }
      if (method === 'hcp') {
        // 蛇形分配:讓每組強弱平均
        var dir = 1, gi = 0;
        for (var k = 0; k < n; k++) {
          groups[gi].push(players[k]);
          if (dir === 1) { if (gi === numGroups - 1) dir = -1; else gi++; }
          else           { if (gi === 0)            dir = 1;  else gi--; }
        }
      } else {
        // 報名順序 / 隨機:平均切塊(避免最後一組只剩 1 人)
        var base = Math.floor(n / numGroups), rem = n % numGroups, idx = 0;
        for (var g2 = 0; g2 < numGroups; g2++) {
          var cnt = base + (g2 < rem ? 1 : 0);
          groups[g2] = players.slice(idx, idx + cnt);
          idx += cnt;
        }
      }
    }
  } else {
    // 有同組需求 → 同組為第一順位,先綁同組;其餘名額再依方法填、盡量平均
    var singles = cl.singles.slice();
    if (method === 'random') {
      for (var s1 = singles.length - 1; s1 > 0; s1--) {
        var s2 = Math.floor(Math.random() * (s1 + 1));
        var st = singles[s1]; singles[s1] = singles[s2]; singles[s2] = st;
      }
    } else if (method === 'hcp') {
      singles.sort(function (a, b) {
        var ha = isNaN(a.hcp) ? 9999 : a.hcp, hb = isNaN(b.hcp) ? 9999 : b.hcp;
        return ha - hb;
      });
    }
    var multi = cl.multi.slice().sort(function (a, b) { return b.length - a.length; });
    multi.forEach(function (c) { placeCluster(c, groups, size); });
    if (method === 'gender') {
      // 同組綁完後,剩餘單人依性別平均填(男先女後)
      singles.filter(function (p) { return p.gender === '男'; })
             .forEach(function (p) { placeByGenderBalance(p, groups, size); });
      singles.filter(function (p) { return p.gender === '女'; })
             .forEach(function (p) { placeByGenderBalance(p, groups, size); });
      singles.filter(function (p) { return p.gender !== '男' && p.gender !== '女'; })
             .forEach(function (p) { groups[emptiestGroupIdx(groups, size)].push(p); });
    } else {
      singles.forEach(function (s) { groups[emptiestGroupIdx(groups, size)].push(s); });
    }
    groups = groups.filter(function (gp) { return gp.length > 0; });
  }

  // 排序規則(除了「二男二女」):
  //  1) 組內女生越多 → 排越前面(例:4 女組在 4 男組前面)
  //  2) 女生數相同時 → 人數少的組在前(3 人組排在 4 人組前面)
  if (method !== 'gender') {
    groups = groups
      .map(function (gp, i) {
        var w = gp.filter(function (x) { return x.gender === '女'; }).length;
        return { gp: gp, w: w, i: i };
      })
      .sort(function (a, b) {
        return (b.w - a.w) || (a.gp.length - b.gp.length) || (a.i - b.i);
      })
      .map(function (o) { return o.gp; });
  }

  var times = [];
  if (start !== null) {
    for (var g3 = 0; g3 < groups.length; g3++) times.push(fmtTime(start + g3 * interval));
  }

  var payload = {
    ts: Date.now(), method: method, size: size,
    groups: groups.map(function (gp) {
      return gp.map(function (pl) {
        return { name: pl.name, hcp: isNaN(pl.hcp) ? '' : String(pl.hcp),
                 cart: pl.cart, userId: pl.userId };
      });
    }),
    times: times
  };
  setConfig('groups', JSON.stringify(payload));
  return { ok: true, groups: payload };
}

function clearGroups(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以清除分組' };
  }
  setConfig('groups', '');
  return { ok: true, groups: null };
}

// ---------- 手動微調分組(交換兩人 / 移動一人 / 改開球時間)----------
function adjustGroups(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以調整分組' };
  var payload = null;
  if (cfg.groups) { try { payload = JSON.parse(cfg.groups); } catch (e) {} }
  if (!payload || !payload.groups || !payload.groups.length) return { ok: false, error: '目前沒有分組可調整' };
  var groups = payload.groups, times = payload.times || [];
  var size = parseInt(payload.size, 10) || 4;
  var op = p.op || '', warnings = [];

  function locate(uid) {
    for (var g = 0; g < groups.length; g++)
      for (var i = 0; i < groups[g].length; i++)
        if (String(groups[g][i].userId) === String(uid)) return { g: g, i: i };
    return null;
  }
  function dropEmpty() {
    for (var g = groups.length - 1; g >= 0; g--) {
      if (groups[g].length === 0) { groups.splice(g, 1); if (times.length > g) times.splice(g, 1); }
    }
  }

  if (op === 'swap') {
    var la = locate(p.a), lb = locate(p.b);
    if (!la || !lb) return { ok: false, error: '找不到要交換的人' };
    if (la.g === lb.g) return { ok: false, error: '兩人已在同一組' };
    var t = groups[la.g][la.i]; groups[la.g][la.i] = groups[lb.g][lb.i]; groups[lb.g][lb.i] = t;

  } else if (op === 'move') {
    var lm = locate(p.who);
    if (!lm) return { ok: false, error: '找不到要移動的人' };
    var to = parseInt(p.to, 10);
    if (isNaN(to) || to < 0 || to >= groups.length) return { ok: false, error: '目標組別不存在' };
    if (to === lm.g) return { ok: false, error: '此人已在該組' };
    var pl = groups[lm.g].splice(lm.i, 1)[0];
    groups[to].push(pl);
    if (groups[to].length > size) warnings.push('第' + (to + 1) + '組已有 ' + groups[to].length + ' 人(超過 ' + size + ' 人)');
    dropEmpty();

  } else if (op === 'times') {
    if (Object.prototype.toString.call(p.times) === '[object Array]') {
      times = groups.map(function (_, i) { return p.times[i] != null ? String(p.times[i]) : (times[i] || ''); });
    } else {
      var start = parseTime(String(p.start || '').trim());
      var interval = parseInt(p.interval, 10); if (isNaN(interval) || interval < 0) interval = 0;
      if (start === null) return { ok: false, error: '起始時間格式需為 HH:MM(例 11:00)' };
      times = [];
      for (var g3 = 0; g3 < groups.length; g3++) times.push(fmtTime(start + g3 * interval));
    }

  } else {
    return { ok: false, error: '未知的操作' };
  }

  payload.groups = groups; payload.times = times; payload.ts = Date.now();
  setConfig('groups', JSON.stringify(payload));
  return { ok: true, groups: payload, warnings: warnings };
}

// 依「備註(同組)」把互相點名的人綁成同一群;回傳 {multi:[群], singles:[單人]}
function buildSameGroupClusters(players) {
  var parent = {};
  players.forEach(function (p) { parent[p.name] = p.name; });
  function find(x) { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; }
  function union(a, b) { var ra = find(a), rb = find(b); if (ra !== rb) parent[ra] = rb; }
  // 若某人的備註裡出現另一位報名者的姓名,就把兩人綁在一起(單向點名即成立)
  players.forEach(function (p) {
    if (!p.note) return;
    players.forEach(function (q) {
      if (q.name && q.name !== p.name && p.note.indexOf(q.name) >= 0) union(p.name, q.name);
    });
  });
  // 來賓與邀請人強制同組(依 userId:guest 的 userId 形如 G:<邀請人userId>:<姓名>)
  var uidName = {};
  players.forEach(function (p) { if (p.userId) uidName[p.userId] = p.name; });
  players.forEach(function (p) {
    if (p.userId && p.userId.indexOf('G:') === 0) {
      var invName = uidName[p.userId.split(':')[1]];
      if (invName) union(p.name, invName);
    }
  });
  var byRoot = {};
  players.forEach(function (p) { var r = find(p.name); (byRoot[r] = byRoot[r] || []).push(p); });
  var multi = [], inMulti = {};
  Object.keys(byRoot).forEach(function (k) {
    if (byRoot[k].length >= 2) {
      multi.push(byRoot[k]);
      byRoot[k].forEach(function (p) { inMulti[p.name] = true; });
    }
  });
  var singles = players.filter(function (p) { return !inMulti[p.name]; });  // 保留報名順序
  return { multi: multi, singles: singles };
}
// 把一群人放進同一組(最佳適配:塞進剛好放得下的最小剩餘組);放不下才拆到最空的組
function placeCluster(cluster, groups, size) {
  var target = -1, bestRem = 1e9;
  for (var i = 0; i < groups.length; i++) {
    var rem = size - groups[i].length;
    if (rem >= cluster.length && rem < bestRem) { bestRem = rem; target = i; }
  }
  if (target >= 0) { for (var k = 0; k < cluster.length; k++) groups[target].push(cluster[k]); return; }
  for (var k2 = 0; k2 < cluster.length; k2++) groups[emptiestGroupIdx(groups, size)].push(cluster[k2]);
}
// 剩餘容量最大的組(讓人數平均)
function emptiestGroupIdx(groups, size) {
  var best = 0, bestRem = -1e9;
  for (var i = 0; i < groups.length; i++) {
    var rem = size - groups[i].length;
    if (rem > bestRem) { bestRem = rem; best = i; }
  }
  return best;
}
// 把人放進「同性別人數最少、且還有空位」的組,讓各組男女人數平均
function placeByGenderBalance(player, groups, size) {
  var g = player.gender, best = -1, bestCnt = 1e9, bestRem = -1;
  for (var i = 0; i < groups.length; i++) {
    var rem = size - groups[i].length;
    if (rem <= 0) continue;
    var cnt = 0;
    for (var j = 0; j < groups[i].length; j++) if (groups[i][j].gender === g) cnt++;
    if (cnt < bestCnt || (cnt === bestCnt && rem > bestRem)) { bestCnt = cnt; bestRem = rem; best = i; }
  }
  if (best < 0) best = emptiestGroupIdx(groups, size);
  groups[best].push(player);
}

// ---------- 差點:匯入與更新 ----------

// 匯入去年初始差點。text 每行「姓名 差點」或「姓名,差點」
function importHcp(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以匯入差點' };
  }
  var rows = parseLines(p.text);
  var count = 0, hmap = {};
  rows.forEach(function (r) {
    var name = r[0], hcp = parseFloat(r[1]);
    if (name && !isNaN(hcp)) { hmap[name] = round1(hcp); count++; }
  });
  setHcpBothBatch_(hmap);                 // 一次寫回所有差點(原本逐人 setHcpBoth)
  return { ok: true, imported: count, hcpList: getHcpList() };
}

// 輸入本月桿數,系統自動更新差點。text 每行「姓名 總桿」
// 新差點 = 舊差點 × (1−k) + (總桿 − 標準桿) × k
// 輸入本月桿數「姓名 OUT IN」,系統算總桿/淨桿排名(只限會員)並更新差點
// ===== 本月成績鎖定 / 重送覆蓋 =====
function lockedDatesArr_(cfg) {
  try { var a = JSON.parse(cfg.lockedDates || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
// 各場匯出的 PDF 連結(以日期為鍵),供逐月成績顯示
function pdfLinksObj_(cfg) {
  try { var o = JSON.parse(cfg.pdfLinks || '{}'); return (o && typeof o === 'object') ? o : {}; } catch (e) { return {}; }
}
function isDateLocked_(cfg, date) {
  var d = mdOf_(date) || String(date || '').trim();
  return lockedDatesArr_(cfg).indexOf(d) >= 0;
}
function lockScoresAction(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以鎖定成績' };
  var d = mdOf_(p.date) || String(p.date || '').trim();
  if (!d) return { ok: false, error: '缺少日期' };
  var arr = lockedDatesArr_(cfg);
  if (arr.indexOf(d) < 0) { arr.push(d); setConfig('lockedDates', JSON.stringify(arr)); }
  return { ok: true, locked: true, date: d };
}
function unlockScoresAction(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以解鎖成績' };
  var d = mdOf_(p.date) || String(p.date || '').trim();
  var arr = lockedDatesArr_(cfg).filter(function (x) { return x !== d; });
  setConfig('lockedDates', JSON.stringify(arr));
  return { ok: true, locked: false, date: d };
}
// 刪掉某日期的技術獎(批次:整批重寫,不逐列 deleteRow)
function deleteTechForDate_(date) {
  var tsh = sheet(SHEET_TECH);
  removeRowsForDate_(tsh, tsh.getDataRange().getValues(), date);
}
// 批次移除某日期的列(1 次清空 + 1 次寫回,取代逐列 deleteRow)
function removeRowsForDate_(sh, data, date) {
  if (!data || data.length < 2) return;
  var width = data[0].length, kept = [];
  for (var i = 1; i < data.length; i++) {
    if (mdOf_(data[i][0]) !== mdOf_(date)) kept.push(data[i]);
  }
  if (kept.length === data.length - 1) return;            // 沒有要刪的
  sh.getRange(2, 1, data.length - 1, width).clearContent();
  if (kept.length) sh.getRange(2, 1, kept.length, width).setValues(kept);
}
// 批次更新「目前差點」(1 次讀、1 次寫;新名字補列),取代多次 setHcp
function setHcpBatch_(map) {
  var names = Object.keys(map || {}); if (!names.length) return;
  var sh = sheet(SHEET_HCP), data = sh.getDataRange().getValues();
  var seen = {}, col2 = [];
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][0] || '');
    if (map.hasOwnProperty(nm)) { col2.push([map[nm]]); seen[nm] = true; }
    else col2.push([data[i][1]]);
  }
  if (data.length > 1) sh.getRange(2, 2, data.length - 1, 1).setValues(col2);   // 一次寫回「目前差點」欄
  var add = [];
  names.forEach(function (nm) { if (!seen[nm]) add.push([nm, map[nm], map[nm]]); });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  __hcpMapCache = null; cacheDel_('hcp');
}
// 批次更新「目前差點 + 季初差點」(1 次讀、1 次寫 col2/col3;新名字補列),供匯入/年初重設用
function setHcpBothBatch_(map) {
  var names = Object.keys(map || {}); if (!names.length) return;
  var sh = sheet(SHEET_HCP), data = sh.getDataRange().getValues();
  var seen = {}, cols = [];
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][0] || '');
    if (map.hasOwnProperty(nm)) { cols.push([map[nm], map[nm]]); seen[nm] = true; }
    else cols.push([data[i][1], data[i][2]]);
  }
  if (data.length > 1) sh.getRange(2, 2, data.length - 1, 2).setValues(cols);   // 一次寫回 目前+季初 兩欄
  var add = [];
  names.forEach(function (nm) { if (!seen[nm]) add.push([nm, map[nm], map[nm]]); });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  __hcpMapCache = null; cacheDel_('hcp');
}
// 重送=覆蓋:還原該場差點(批次)+ 刪掉該日期的成績與技術獎(批次)
function clearMatchForResubmit_(date) {
  var minfo = getMemberInfoMap();
  var ssh = sheet(SHEET_SCORES), sdata = ssh.getDataRange().getValues();
  var restore = {};
  for (var i = 1; i < sdata.length; i++) {
    if (mdOf_(sdata[i][0]) === mdOf_(date)) {
      var nm = String(sdata[i][1] || ''), info = minfo[nm], pre = sdata[i][5];
      if (nm && (!info || info.role !== '來賓') && pre !== '' && pre != null && !isNaN(parseFloat(pre))) restore[nm] = Number(pre);
    }
  }
  setHcpBatch_(restore);                       // 批次還原賽前差點(撤銷上次扣桿)
  removeRowsForDate_(ssh, sdata, date);        // 批次刪該日期成績
  deleteTechForDate_(date);                    // 批次刪該日期技術獎
  var cfg = getConfig(), links = pdfLinksObj_(cfg), k = mdOf_(date);
  if (links[k]) { delete links[k]; setConfig('pdfLinks', JSON.stringify(links)); }   // 成績被覆蓋 → 清掉舊 PDF 連結(避免顯示過時報表,請重新匯出)
}

// 差點上限(預設 36,可由 Config maxHcp 調整)+ 夾限到 [0, 上限]、四捨五入整數
function maxHcp_(cfg) { var m = parseFloat((cfg || getConfig()).maxHcp); return (isNaN(m) || m <= 0) ? 36 : m; }
function clampHcp_(v, cfg) { var m = maxHcp_(cfg); v = Math.round(v); if (v < 0) v = 0; if (v > m) v = m; return v; }

// 新會員判定 + 取得本年度過往成績(用來累積前兩場算差點)
function isNewMember_(t) { return String(t || '').indexOf('新') >= 0; }
function priorSeasonScores_(name, cfg, excludeDate) {
  var range = seasonRange(cfg), sinceTs = parseFloat(cfg.seasonStartTs) || 0;
  var data = sheet(SHEET_SCORES).getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '') !== name) continue;
    var d = mdOf_(data[i][0]);
    if (excludeDate && d === mdOf_(excludeDate)) continue;          // 排除本場(已清)
    var ok = range ? inCurrentSeason(d, range) : (Number(data[i][10] || 0) >= sinceTs);
    if (!ok) continue;
    var g = Number(data[i][4]) || 0;
    if (g > 0) out.push({ gross: g, ts: Number(data[i][10]) || 0 });
  }
  out.sort(function (a, b) { return a.ts - b.ts; });                // 依時間先後(第一場在最前)
  return out;
}

function submitScores(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以輸入成績' };
  }
  var date = (p.date || cfg.date || '').toString().trim();
  if (isDateLocked_(cfg, date)) return { ok: false, error: '本場成績已鎖定,如需修改請先到管理員專區解鎖。' };
  clearMatchForResubmit_(date);   // 重送=覆蓋:還原差點並清掉舊成績/技術獎(改錯字重送才不會重複、不會重覆扣)
  var par = parseFloat(cfg.par); if (isNaN(par)) par = 72;
  var hmap = getHcpMap();
  var minfo = getMemberInfoMap();
  var typeMap = {}; getMasterList().forEach(function (m) { typeMap[m.name] = m.memberType; });
  var newFactor = parseFloat(cfg.newHcpFactor); if (isNaN(newFactor) || newFactor <= 0 || newFactor > 1) newFactor = 0.9;

  // 解析每行:姓名 OUT IN(out/in 是前後九洞桿數)
  var entries = parseScoreLines(p.text).map(function (e) {
    var name = e.name, out = e.out, inn = e.in;
    var gross = out + inn;
    var info = minfo[name] || { role: '會員', gender: '', birthYear: '' };
    var isGuest = info.role === '來賓';
    var inHmap = hmap.hasOwnProperty(name);
    var newMember1st = false, establishNew = false, hcp, net;
    if (!isGuest && isNewMember_(typeMap[name]) && !inHmap) {
      // 新會員、差點尚未建立 → 看是第幾場
      var prior = priorSeasonScores_(name, cfg, date);
      if (prior.length === 0) {
        newMember1st = true; hcp = ''; net = '';                       // 第一場:不算差點、不排名、只記總桿
      } else {
        var g1 = prior[0].gross;                                       // 第一場總桿
        var est = clampHcp_(((g1 + gross) / 2 - par) * newFactor, cfg); // ((第一場+第二場)/2 − 標準桿)× 倍率,夾限 [0, 上限]
        establishNew = true; hcp = est; net = round1(gross - est);     // 第二場:當即建立差點 → 算淨桿 → 進排名
      }
    } else {
      hcp = inHmap ? hmap[name] : 0; net = round1(gross - hcp);
    }
    return { name: name, out: out, in: inn, gross: gross, hcp: hcp, net: net,
             role: info.role, birthYear: parseInt(info.birthYear, 10) || 0,
             newMember1st: newMember1st, establishNew: establishNew };
  });

  // 排名只算「會員」;來賓記成績但不列入名次;新會員首場(尚無差點)也不列入名次
  var members = entries.filter(function (e) { return e.role !== '來賓' && !e.newMember1st; });

  // 本年度已當過總桿冠軍者,一年只能一次 → 不能再當(依年度區間日期判斷;同一場重輸不算)
  var usedGross = getPastGrossChamps(cfg, date);

  // 總桿排名:總桿低→高;平手 差點高者得、年長(出生年小)者得
  var grossSorted = members.slice().sort(function (a, b) {
    if (a.gross !== b.gross) return a.gross - b.gross;
    if (a.hcp !== b.hcp) return b.hcp - a.hcp;            // 差點高者得
    return ageRank(a) - ageRank(b);                        // 年長者得
  });
  // 總桿冠軍 = 總桿最低、且本季尚未當過總桿冠軍的那一位
  var grossWinner = null;
  for (var gi = 0; gi < grossSorted.length; gi++) {
    if (!usedGross[grossSorted[gi].name]) { grossWinner = grossSorted[gi].name; break; }
  }

  // 淨桿排名:排除本月總桿冠軍(已得過總桿者回到淨桿池);淨桿低→高;平手 後九洞(IN)低→差點低→年長
  var netSorted = members.filter(function (e) { return e.name !== grossWinner; })
    .sort(function (a, b) {
      if (a.net !== b.net) return a.net - b.net;
      if (a.in !== b.in) return a.in - b.in;               // 後九洞低者得
      if (a.hcp !== b.hcp) return a.hcp - b.hcp;           // 差點低者得
      return ageRank(a) - ageRank(b);                      // 年長者得
    });

  // 淨桿特殊獎(限非前三的會員):
  //  幸運獎   = 淨桿第 luckyPlace 名(預設 7)
  //  逢五跳獎 = 淨桿第 5、10、15… 名(間隔 skipStep,預設 5,依人數)
  //  BB獎     = 淨桿倒數第二名
  //  優先序 BB > 幸運獎 > 逢五跳獎;BB 與逢五同名次時只發 BB(該逢五作廢)
  var netCount = netSorted.length;
  var luckyPlace = parseInt(cfg.luckyPlace, 10) || 7;
  var skipStep = parseInt(cfg.skipStep, 10) || 5;
  var bbPos = (netCount >= 5) ? (netCount - 1) : -1;     // 倒數第二的名次

  // 標記名次 + 計算賽後差點(前三名淨桿依級距扣)
  var result = [], scoreRows = [], hcpUpdates = {}, scoreSheet = sheet(SHEET_SCORES);
  entries.forEach(function (e) {
    var rankType = '', rank = '';
    if (e.role === '來賓') { rankType = '來賓'; }
    else if (e.newMember1st) { rankType = '新會員'; }                  // 首場,待第二場累積差點
    else if (e.name === grossWinner) { rankType = '總桿'; rank = '冠軍'; }
    else {
      var pos = indexByName(netSorted, e.name) + 1;
      rank = String(pos);
      rankType = '淨桿';
      if (pos > 3) {
        if (pos === bbPos) rankType = 'BB';
        else if (pos === luckyPlace) rankType = '幸運獎';
        else if (skipStep > 0 && pos % skipStep === 0) rankType = '逢五跳獎';
      }
    }
    var newHcp = e.hcp;
    if (e.newMember1st) {
      newHcp = '';                                                     // 還沒有差點
    } else if (e.establishNew) {
      newHcp = e.hcp;                                                  // 差點剛建立,本場不再扣
    } else if (e.role !== '來賓' && e.name !== grossWinner) {
      var pos2 = indexByName(netSorted, e.name) + 1;
      var adj = hcpCut(pos2, e.hcp);                       // 級距扣(整數)
      if (pos2 >= 1 && pos2 <= 3 && e.net < par) {          // 淨桿低於標準桿 → 額外再扣
        adj += -(par - e.net) * belowParPct(e.hcp);
      }
      newHcp = round1(e.hcp + Math.round(adj));             // 調整四捨五入到整數後套用
      if (newHcp < 0) newHcp = 0;
      if (newHcp > maxHcp_(cfg)) newHcp = maxHcp_(cfg);     // 不超過差點上限
    }
    if (e.role !== '來賓' && !e.newMember1st) { hcpUpdates[e.name] = newHcp; hmap[e.name] = newHcp; }
    scoreRows.push([date, e.name, e.out, e.in, e.gross, e.hcp, e.net,
                    rankType, rank, newHcp, Date.now()]);
    result.push({ name: e.name, role: e.role, out: e.out, in: e.in, gross: e.gross,
                  hcp: e.hcp, net: e.net, rankType: rankType, rank: rank, newHcp: newHcp });
  });
  setHcpBatch_(hcpUpdates);                                 // 一次寫回所有差點
  if (scoreRows.length) {                                   // 一次寫入所有成績列
    scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoreRows.length, scoreRows[0].length).setValues(scoreRows);
  }

  // 存最新一場排名供前端顯示
  var ranking = { date: date, results: result.slice().sort(rankSortForDisplay) };
  setConfig('lastRanking', JSON.stringify(ranking));

  return { ok: true, count: result.length, ranking: ranking, hcpList: getHcpList() };
}

// 前三名淨桿依差點級距扣桿
function hcpCut(pos, hcp) {
  var t = [
    [-1, -2, -3, -4],   // 冠軍
    [ 0, -1, -2, -3],   // 亞軍
    [ 0,  0, -1, -2]    // 季軍
  ];
  if (pos < 1 || pos > 3) return 0;
  return t[pos - 1][hcpBracket(hcp)];
}
// 淨桿低於標準桿時的額外扣除百分比(依差點級距)
function belowParPct(hcp) {
  return [0.25, 0.33, 0.50, 1.00][hcpBracket(hcp)];
}
// 差點級距:0=0~9, 1=10~19, 2=20~29, 3=30+
function hcpBracket(hcp) {
  return hcp < 10 ? 0 : (hcp < 20 ? 1 : (hcp < 30 ? 2 : 3));
}
function ageRank(e) { return e.birthYear > 0 ? e.birthYear : 9999; }  // 出生年小=年長=排前
// 本季已當過總桿冠軍的姓名集合(從 Scores 找 rankType='總桿',且 ts>=季初,排除指定日期)
// 本年度已當過總桿冠軍的姓名集合。優先用年度區間日期判斷;沒設年度才退回用 ts
function getPastGrossChamps(cfg, excludeDate) {
  var range = seasonRange(cfg);
  var sinceTs = parseFloat(cfg.seasonStartTs) || 0;
  var data = sheet(SHEET_SCORES).getDataRange().getValues(), set = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][7]) !== '總桿') continue;
    var d = mdOf_(data[i][0]);
    if (excludeDate && d === mdOf_(excludeDate)) continue;        // 同一場重輸不算
    var ok = range ? inCurrentSeason(d, range) : (Number(data[i][10] || 0) >= sinceTs);
    if (ok) set[String(data[i][1])] = true;
  }
  return set;
}

// ---------- 年度區間 ----------
function parseYM(s) {
  var m = String(s || '').split(/[\/\-\.]/);
  if (m.length < 2) return null;
  var y = parseInt(m[0], 10), mo = parseInt(m[1], 10);
  if (!y || !mo) return null;
  return { y: y, m: mo };
}
// 回傳 {start, end, fy, fm} 或 null(未設年度)
function seasonRange(cfg) {
  var f = parseYM(cfg.seasonFrom), t = parseYM(cfg.seasonTo);
  if (!f || !t) return null;
  var start = new Date(f.y, f.m - 1, 1);
  var end = new Date(t.y, t.m, 0); end.setHours(23, 59, 59, 999);   // 年度迄月份最後一天
  return { start: start, end: end, fy: f.y, fm: f.m };
}
// 把成績的「M/D」依年度推算完整日期(月份>=年度起月→起始年,否則→次年)
function scoreFullDate(dateStr, range) {
  var m = String(dateStr || '').split('/');
  var mo = parseInt(m[0], 10), da = parseInt(m[1], 10) || 1;
  if (!mo) return null;
  var year = (mo >= range.fm) ? range.fy : range.fy + 1;
  return new Date(year, mo - 1, da);
}
function inCurrentSeason(dateStr, range) {
  if (!range) return true;
  var d = scoreFullDate(dateStr, range);
  if (!d) return false;
  return d >= range.start && d <= range.end;
}
// 依名次回傳獎金(總桿冠軍 / 淨桿冠亞季)
function prizeFor(rankType, rank, cfg) {
  if (rankType === '總桿') return parseFloat(cfg.prizeGross) || 0;
  if (rankType === '幸運獎') return parseFloat(cfg.prizeLucky) || 0;
  if (rankType === '逢五跳獎') return parseFloat(cfg.prizeSkip5) || 0;
  if (rankType === 'BB') return parseFloat(cfg.prizeBB) || 0;
  if (rankType === '淨桿') {
    if (rank === '1') return parseFloat(cfg.prizeNet1) || 0;
    if (rank === '2') return parseFloat(cfg.prizeNet2) || 0;
    if (rank === '3') return parseFloat(cfg.prizeNet3) || 0;
  }
  return 0;
}
// 某人本年度的出席與獎金統計
function getMyStats(name, cfg) {
  if (!name) return null;
  var range = seasonRange(cfg);
  var data = sheet(SHEET_SCORES).getDataRange().getValues();
  var dateSet = {}, played = 0, prize = 0;
  for (var i = 1; i < data.length; i++) {
    var d = mdOf_(data[i][0]); if (!d) continue;
    if (range && !inCurrentSeason(d, range)) continue;
    dateSet[d] = true;
    if (String(data[i][1] || '') === name) {
      played++;
      prize += prizeFor(String(data[i][7] || ''), String(data[i][8] || ''), cfg);
    }
  }
  var total = Object.keys(dateSet).length;
  return { played: played, total: total, prize: prize, perfect: (total > 0 && played === total) };
}

// ---------- 技術獎項 + 送球庫存 ----------
var TECH_AWARDS = ['近洞獎','二近洞獎','三近洞獎','BIRDIE獎','Eagle獎','一桿進洞獎','來賓禮'];

// 登錄某場技術獎:每行「獎項 得獎者 [條數]」,條數省略預設 1
function recordTech(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以登錄技術獎' };
  var date = String(p.date || '').trim();
  if (!date) return { ok: false, error: '請填日期(如 7/21)' };
  if (isDateLocked_(cfg, date)) return { ok: false, error: '本場成績已鎖定,如需修改請先解鎖。' };
  var lines = String(p.text || '').split(/[\r\n]+/), rows = [];
  lines.forEach(function (ln) {
    var t = ln.trim(); if (!t) return;
    var toks = t.split(/[\s,，、\t]+/).filter(function (x) { return x !== ''; });
    if (toks.length < 2) return;
    var award = toks[0];
    var tiao = 1, winnerToks = toks.slice(1);
    var last = winnerToks[winnerToks.length - 1];
    if (/^\d+(\.\d+)?$/.test(last)) { tiao = parseFloat(last); winnerToks = winnerToks.slice(0, -1); }
    var winner = winnerToks.join(' ');
    rows.push([date, award, winner, tiao]);
  });
  if (!rows.length) return { ok: false, error: '沒有可登錄的資料' };
  deleteTechForDate_(date);   // 重送=覆蓋:先清掉該日期舊技術獎
  var sh = sheet(SHEET_TECH);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { ok: true, added: rows.length, balls: getBallStat(cfg), techList: getTechList(cfg) };
}

// 例賽結算:成績 + 技術獎一次送出(共用同一日期)
function submitMatch(userId, p) {
  var sres = submitScores(userId, p);          // p: date, text(成績)
  if (!sres.ok) return sres;
  var techAdded = 0;
  if (p.tech && String(p.tech).trim()) {
    var tres = recordTech(userId, { date: p.date, text: p.tech });
    if (tres.ok) techAdded = tres.added;
  }
  return { ok: true, count: sres.count, ranking: sres.ranking, hcpList: sres.hcpList, techAdded: techAdded };
}

// 刪除某日期的技術獎登錄(重輸用)
function clearTechDate(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以刪除' };
  var date = String(p.date || '').trim();
  var sh = sheet(SHEET_TECH), data = sh.getDataRange().getValues(), removed = 0;
  for (var i = 1; i < data.length; i++) { if (mdOf_(data[i][0]) === mdOf_(date)) removed++; }
  if (removed) removeRowsForDate_(sh, data, date);   // 批次刪(1 次清空 + 1 次寫回,取代逐列 deleteRow)
  return { ok: true, removed: removed, balls: getBallStat(cfg), techList: getTechList(cfg) };
}

// 技術獎點選用名單:Members(已綁定) + 會員名單主檔,去重排序

// 技術獎清單(本年度),依日期年度序
function getTechList(cfg) {
  var range = seasonRange(cfg);
  var data = sheet(SHEET_TECH).getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    var d = mdOf_(data[i][0]); if (!d) continue;
    if (range && !inCurrentSeason(d, range)) continue;
    out.push({ date: d, award: String(data[i][1] || ''), winner: String(data[i][2] || ''), tiao: parseFloat(data[i][3]) || 0,
               ord: range ? scoreFullDate(d, range).getTime() : 0 });
  }
  out.sort(function (a, b) { return a.ord - b.ord; });
  return out;
}

// 球數庫存統計(條為單位,1盒=4條)
function getBallStat(cfg) {
  var boxes = parseFloat(cfg.ballInitBoxes) || 0;
  var note = parseFloat(cfg.ballInitNote) || 0;     // 期初零散條數
  var init = boxes * 4 + note;                       // 期初總條
  var list = getTechList(cfg);
  var byDate = {}, order = {}, dates = [];
  list.forEach(function (r) {
    if (!(r.date in byDate)) { byDate[r.date] = 0; order[r.date] = r.ord; dates.push(r.date); }
    byDate[r.date] += r.tiao;
  });
  dates.sort(function (a, b) { return order[a] - order[b]; });
  var totalOut = 0, ledger = [];
  dates.forEach(function (d) {
    totalOut += byDate[d];
    ledger.push({ date: d, out: byDate[d], balance: init - totalOut });
  });
  return { boxes: boxes, note: note, init: init, out: totalOut, balance: init - totalOut, ledger: ledger };
}

// 產生「球數庫存」分頁
function buildBalls(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以產生球數庫存' };
  var st = getBallStat(cfg);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('球數庫存'); if (!sh) sh = ss.insertSheet('球數庫存');
  sh.clearContents();
  var rows = [];
  rows.push(['球數庫存(單位:條,1盒=4條)  年度 ' + (cfg.seasonFrom || '') + '–' + (cfg.seasonTo || '')]);
  rows.push(['期初', '盒 ' + st.boxes + ' + 散 ' + st.note + ' 條 = ' + st.init + ' 條']);
  rows.push(['日期', '當月送出(條)', '累計送出(條)', '結存(條)', '結存(盒+條)']);
  var cum = 0;
  st.ledger.forEach(function (r) {
    cum += r.out;
    var bal = st.init - cum;
    rows.push([r.date, r.out, cum, bal, Math.floor(bal / 4) + ' 盒 ' + (bal % 4) + ' 條']);
  });
  rows.push(['結存', '', st.out, st.balance, Math.floor(st.balance / 4) + ' 盒 ' + (((st.balance % 4) + 4) % 4) + ' 條']);
  var width = 5;
  var grid = rows.map(function (r) { while (r.length < width) r.push(''); return r; });
  sh.getRange(1, 1, grid.length, width).setValues(grid);
  sh.getRange(1, 1, 1, width).merge().setFontWeight('bold');
  sh.getRange(3, 1, 1, width).setFontWeight('bold');
  sh.setFrozenRows(3);
  return { ok: true, balance: st.balance, balls: st };
}

// ---------- 匯出當月例賽 PDF ----------
function latestScoreDate() {
  var data = sheet(SHEET_SCORES).getDataRange().getValues();
  var best = '', bestTs = -1;
  for (var i = 1; i < data.length; i++) {
    var ts = Number(data[i][10] || 0);
    if (mdOf_(data[i][0]) && ts >= bestTs) { bestTs = ts; best = mdOf_(data[i][0]); }
  }
  return best;
}
// 名次/獎別文字
function nextEventAfter(cfg) {
  var sched = getSchedule(), now = new Date();
  var todayKey = now.getFullYear() * 10000 + (now.getMonth() + 1) * 100 + now.getDate();
  for (var i = 0; i < sched.length; i++) { if (sched[i].sortKey >= todayKey) return sched[i]; }
  return null;
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
// 日期欄防呆:試算表把日期存成 Date 物件時,統一轉成 M/D;字串原樣保留
function mdOf_(v) {
  if (v instanceof Date) return (v.getMonth() + 1) + '/' + v.getDate();
  return String(v == null ? '' : v).trim();
}
// PDF 成績總表的半欄(供兩欄並排,節省版面塞進一頁)
// 解析某場球場名:行事曆同日優先,否則用設定的球場
function resolveCourse(date, cfg) {
  var sched = getSchedule();
  for (var i = 0; i < sched.length; i++) {
    if (String(sched[i].date) === date && sched[i].course) return sched[i].course;
  }
  return cfg.course || '';
}
// 球隊簡稱(去掉「高爾夫球隊/球隊/高爾夫」字尾)供檔名用
function teamShort(cfg) {
  return String(cfg.teamName || '球隊').replace(/高爾夫球隊$|球隊$|高爾夫$/, '') || '球隊';
}

// PDF 成績總表(單欄,姓名|OUT|IN|總桿|差|淨桿|調後;冠軍金底、調後變動紅字)
function pdfScoreFull_(list, champName) {
  if (!list.length) return '<p class="muted">無會員成績</p>';
  var b = list.map(function (r) {
    var changed = (r.hcpAfter !== '' && String(r.hcpAfter) !== String(r.hcp));
    var after = (r.hcpAfter === '' ? '—' : esc_(r.hcpAfter));
    var champ = (champName && r.name === champName) || r.rankType === '總桿';
    var nm = esc_(r.name) + (champ ? ' ♔' : '');
    return '<tr' + (champ ? ' class="champ"' : '') + '><td>' + nm + '</td><td class="c">' + esc_(r.out) +
      '</td><td class="c">' + esc_(r.in) + '</td><td class="c">' + esc_(r.gross) + '</td><td class="c">' + esc_(r.hcp) +
      '</td><td class="c">' + esc_(r.net) + '</td><td class="c' + (changed ? ' chg' : '') + '">' + after + '</td></tr>';
  }).join('');
  return '<table class="t sc"><tr><th>姓名</th><th class="c">OUT</th><th class="c">IN</th><th class="c">總桿</th>' +
    '<th class="c">差</th><th class="c">淨桿</th><th class="c">調後</th></tr>' + b + '</table>';
}
// 依會員人數決定字級(讓內容約填到離頁底 ~2cm、維持單頁);可用 Config pdfFontPx 覆寫
function pdfFontForCount_(n, low, cfg) {
  var ov = parseFloat(cfg && cfg.pdfFontPx);
  if (!isNaN(ov) && ov > 0) return ov;
  var fs;
  if (low) {                       // <25 人:成績表+來賓都在左欄、含統計列
    if (n <= 10) fs = 18.5;
    else if (n <= 14) fs = 17.5;
    else if (n <= 18) fs = 16;
    else fs = 14.5;                // 19~24
  } else {                         // >=25 人:來賓移右欄、無統計列
    if (n <= 28) fs = 14;
    else if (n <= 32) fs = 13;
    else if (n <= 36) fs = 12;
    else if (n <= 40) fs = 11.5;
    else if (n <= 44) fs = 10.5;
    else fs = 10;
  }
  return fs;
}
// 球隊小提醒來源:行事曆該場「備註」欄優先(若有此欄且有填),否則用 Config「球隊小提醒」
function matchNote_(date, cfg) {
  try {
    var sh = sheet(SHEET_SCHED);
    if (sh && sh.getLastRow() > 1) {
      var d = sh.getDataRange().getValues(), hdr = d[0], dc = -1, nc = -1, yc = -1, mc = -1;
      for (var c = 0; c < hdr.length; c++) {
        var h = String(hdr[c]).replace(/\s/g, '');
        if (h === '日期') dc = c;
        if (h === '備註' || h === '提醒事項' || h === '球隊小提醒') nc = c;
        if (h === '年') yc = c;
        if (h === '月') mc = c;
      }
      if (nc >= 0 && dc >= 0) {
        for (var i = 1; i < d.length; i++) {
          var disp = normSchedDate(yc < 0 ? '' : d[i][yc], mc < 0 ? '' : d[i][mc], d[i][dc]).disp;
          if (disp === mdOf_(date)) {
            var v = String(d[i][nc] == null ? '' : d[i][nc]).trim();
            if (v) return v;
            break;
          }
        }
      }
    }
  } catch (e) {}
  return String(cfg['球隊小提醒'] == null ? '' : cfg['球隊小提醒']).trim();
}

function buildMatchReportHtml(date, cfg) {
  var data = sheet(SHEET_SCORES).getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (mdOf_(data[i][0]) !== mdOf_(date)) continue;
    rows.push({ name: String(data[i][1] || ''), out: data[i][2], in: data[i][3], gross: data[i][4],
                hcp: data[i][5], net: data[i][6], rankType: String(data[i][7] || ''), rank: String(data[i][8] || ''),
                hcpAfter: (data[i][9] === '' || data[i][9] == null) ? '' : data[i][9] });
  }
  var grp = function (rt) { return rt === '總桿' ? 0 : (rt === '來賓' ? 2 : 1); };
  rows.sort(function (a, b) {
    if (grp(a.rankType) !== grp(b.rankType)) return grp(a.rankType) - grp(b.rankType);
    return (parseInt(a.rank, 10) || 99) - (parseInt(b.rank, 10) || 99);
  });
  var members = rows.filter(function (r) { return r.rankType !== '來賓'; });
  var guests = rows.filter(function (r) { return r.rankType === '來賓'; });
  var champ = '';
  rows.forEach(function (r) { if (r.rankType === '總桿') champ = r.name; });

  var LOW = members.length < 25;
  var FS = pdfFontForCount_(members.length, LOW, cfg);
  var PAD = Math.round(FS * 0.42 * 10) / 10;
  var H1 = Math.round((FS * 1.18 + 3) * 10) / 10;
  var H2 = Math.round((FS * 1.02 + 1.5) * 10) / 10;
  var H2MT = Math.round(FS * 0.8 * 10) / 10, H2MB = Math.round(FS * 0.33 * 10) / 10;
  var CPAD = Math.round((PAD * 2 + 5) * 10) / 10;
  var CB = Math.round(FS * 1.32 * 10) / 10, CE = Math.round(FS * 0.72 * 10) / 10;
  var NF = Math.round(FS * 0.92 * 10) / 10;

  // 來賓成績表(OUT/IN/總桿)
  var guestTbl = guests.length
    ? ('<h2>來賓成績(' + guests.length + ' 人,不列入排名)</h2><table class="t"><tr><th>姓名</th><th class="c">OUT</th><th class="c">IN</th><th class="c">總桿</th></tr>' +
       guests.map(function (r) { return '<tr><td>' + esc_(r.name) + '</td><td class="c">' + esc_(r.out) + '</td><td class="c">' + esc_(r.in) + '</td><td class="c">' + esc_(r.gross) + '</td></tr>'; }).join('') +
       '</table>')
    : '';

  // 桿數排名(名次與獎金)+ 合計
  var prizeRows = [], prizeSum = 0;
  function addPrize(label, name, amt, strong) {
    if (!name) return;
    var a = Number(amt) || 0; prizeSum += a;
    prizeRows.push('<tr><td' + (strong ? ' class="p"' : '') + '>' + esc_(label) + '</td><td>' + esc_(name) +
      '</td><td class="r">' + (a ? 'NT$' + a.toLocaleString() : '') + '</td></tr>');
  }
  var byType = {};
  rows.forEach(function (r) { (byType[r.rankType] = byType[r.rankType] || []).push(r); });
  (byType['總桿'] || []).forEach(function (r) { addPrize('總桿冠軍', r.name, cfg.prizeGross, true); });
  (byType['淨桿'] || []).forEach(function (r) {
    if (r.rank === '1') addPrize('淨桿冠軍', r.name, cfg.prizeNet1, true);
    else if (r.rank === '2') addPrize('淨桿亞軍', r.name, cfg.prizeNet2, true);
    else if (r.rank === '3') addPrize('淨桿季軍', r.name, cfg.prizeNet3, true);
  });
  (byType['幸運獎'] || []).forEach(function (r) { addPrize('幸運獎(淨' + r.rank + ')', r.name, cfg.prizeLucky); });
  (byType['逢五跳獎'] || []).forEach(function (r) { addPrize('逢五跳獎(淨' + r.rank + ')', r.name, cfg.prizeSkip5); });
  (byType['BB'] || []).forEach(function (r) { addPrize('BB獎(淨' + r.rank + ')', r.name, cfg.prizeBB); });
  var prizeHtml = '<h2>桿數排名(名次與獎金)</h2><table class="t"><tr><th>獎項</th><th>得獎者</th><th class="r">獎金</th></tr>' +
    (prizeRows.join('') || '<tr><td colspan="3" class="muted">無</td></tr>') +
    (prizeSum ? '<tr><td class="p">合計</td><td></td><td class="r p">NT$' + prizeSum.toLocaleString() + '</td></tr>' : '') +
    '</table>';

  // 技術獎(送球)+ 合計送球
  var tdata = sheet(SHEET_TECH).getDataRange().getValues(), techRows = [], techSum = 0;
  for (var j = 1; j < tdata.length; j++) {
    if (mdOf_(tdata[j][0]) !== mdOf_(date)) continue;
    techSum += Number(tdata[j][3]) || 0;
    techRows.push('<tr><td>' + esc_(tdata[j][1]) + '</td><td>' + esc_(tdata[j][2]) + '</td><td class="c">' + esc_(tdata[j][3]) + ' 條</td></tr>');
  }
  var techHtml = techRows.length
    ? ('<h2>技術獎項表(送球)</h2><table class="t"><tr><th>技術獎</th><th>得獎者</th><th class="c">送球</th></tr>' +
       techRows.join('') + '<tr><td class="p">合計送球</td><td></td><td class="c p">' + techSum + ' 條</td></tr></table>')
    : '';

  // 下次例賽
  var nx = nextEventAfter(cfg);
  var nextHtml = '<h2>下次例賽資料</h2>' + (nx
    ? ('<table class="t info"><tr><td class="lb">日期</td><td>' + esc_(nx.date) + (nx.dow ? '(' + String(nx.dow).replace('週', '') + ')' : '') + '</td></tr>' +
       '<tr><td class="lb">球場</td><td>' + esc_(nx.course) + '</td></tr>' +
       '<tr><td class="lb">開球</td><td>' + esc_(nx.tee) + '</td></tr>' +
       (nx.deadline ? '<tr><td class="lb">截止</td><td>' + esc_(nx.deadline) + '</td></tr>' : '') + '</table>')
    : '<p class="muted">行事曆暫無下一場資料。</p>');

  // 球隊小提醒
  var noteItems = matchNote_(date, cfg).split(/\r?\n|｜|\|/);
  var noteHtml = '';
  var ni = [];
  for (var k = 0; k < noteItems.length; k++) { var s = String(noteItems[k]).trim(); if (s) ni.push(s); }
  if (ni.length) {
    noteHtml = '<h2>📣 球隊小提醒</h2><table class="t note">' +
      ni.map(function (s) { return '<tr><td>• ' + esc_(s) + '</td></tr>'; }).join('') + '</table>';
  }

  // 統計列(僅低人數版)
  var countsHtml = LOW
    ? ('<table class="cnt"><tr><td><b>' + rows.length + '</b><em>擊球人數</em></td>' +
       '<td><b>' + members.length + '</b><em>會員擊球</em></td>' +
       '<td><b>' + guests.length + '</b><em>來賓擊球</em></td></tr></table>')
    : '';

  var seasonTxt = (cfg.seasonFrom || cfg.seasonTo) ? ('年度 ' + (cfg.seasonFrom || '') + '–' + (cfg.seasonTo || '')) : '';
  var head = (cfg.teamName || cfg.title || '球隊') + (cfg.term ? ' 第' + cfg.term + '屆' : '');

  var leftCol =
    '<h2>本月例賽</h2>' +
    '<table class="t info"><tr><td class="lb">日期</td><td>' + esc_(date) + '</td><td class="lb">球場</td><td>' + esc_(resolveCourse(date, cfg)) + '</td></tr>' +
    '<tr><td class="lb">開球</td><td>' + esc_(cfg.tee || '') + '</td><td class="lb">類型</td><td>' + esc_((nx && '') || schedTypeOf_(date) || '月例賽') + '</td></tr></table>' +
    countsHtml +
    '<h2>成績總表(會員 ' + members.length + ' 人)<span style="font-size:' + Math.round(FS * 0.62) + 'px;color:#888;font-weight:400">　差=賽前 · 調後=賽後</span></h2>' +
    pdfScoreFull_(members, champ) +
    (LOW ? guestTbl : '');

  var rightCol =
    prizeHtml +
    techHtml +
    (LOW ? '' : guestTbl) +
    nextHtml +
    noteHtml;

  return '' +
  '<html><head><meta charset="utf-8"><style>' +
  '@page{size:A4;margin:0.8cm}' +
  '*{box-sizing:border-box}' +
  'body{font-family:"Noto Sans CJK TC","PingFang TC",sans-serif;color:#1a1a1a;font-size:' + FS + 'px;margin:0}' +
  'h1{font-size:' + H1 + 'px;margin:0;color:#14532d}' +
  '.sub{color:#666;font-size:11px;margin:3px 0 8px}' +
  '.rule{height:3px;background:#2d6a4f;margin:0 0 9px}' +
  'h2{font-size:' + H2 + 'px;border-left:5px solid #2d6a4f;padding-left:8px;margin:' + H2MT + 'px 0 ' + H2MB + 'px;color:#14532d;text-align:left}' +
  '.wrap{border-collapse:collapse;width:100%}.wrap>tbody>tr>td{vertical-align:top;border:0;padding:0}' +
  'td.L{width:56%;padding-right:10px}td.R{width:44%;padding-left:10px;border-left:1px solid #e1e1d8}' +
  'table.t{border-collapse:collapse;width:100%;margin:2px 0 5px}' +
  '.t td,.t th{border:1px solid #cfcfc6;padding:' + PAD + 'px 6px;font-size:' + FS + 'px}' +
  '.t th{background:#eef3ee;color:#14532d;font-weight:700}' +
  '.c{text-align:center}.r{text-align:right}.muted{color:#888}' +
  '.info .lb{background:#f6f6f2;color:#555;width:46px;white-space:nowrap}' +
  '.cnt{border-collapse:collapse;width:100%;margin:6px 0 2px}' +
  '.cnt td{border:1px solid #e3e3da;background:#f7f7f2;text-align:center;padding:' + CPAD + 'px 0}' +
  '.cnt b{display:block;font-size:' + CB + 'px;color:#14532d}.cnt em{font-style:normal;font-size:' + CE + 'px;color:#777}' +
  'tr.champ td{background:#fff1b8;font-weight:700}.chg{color:#bc4b3c;font-weight:700}.p{font-weight:700;color:#14532d}' +
  '.note td{font-size:' + NF + 'px;line-height:1.35;color:#333}' +
  '.foot{margin-top:6px;border-top:1px solid #ddd;padding-top:4px;color:#999;font-size:8.5px;text-align:center}' +
  '</style></head><body>' +
  '<h1>' + esc_(head) + ' · 例賽紀錄</h1>' +
  '<p class="sub">' + esc_(seasonTxt) + ' ｜ 比賽日期 ' + esc_(date) + ' ｜ 球場 ' + esc_(resolveCourse(date, cfg)) + ' ｜ 列印 ' + fmtDate(new Date()) + ' ｜ 會員 ' + members.length + ' 人</p>' +
  '<div class="rule"></div>' +
  '<table class="wrap"><tbody><tr><td class="L">' + leftCol + '</td><td class="R">' + rightCol + '</td></tr></tbody></table>' +
  '<div class="foot">' + esc_(head) + ' 例賽管理系統 · 本頁為例賽成績紀錄</div>' +
  '</body></html>';
}

// 取某場類型(行事曆),供本月例賽顯示
function schedTypeOf_(date) {
  var sched = getSchedule();
  for (var i = 0; i < sched.length; i++) { if (String(sched[i].date) === mdOf_(date) && sched[i].type) return sched[i].type; }
  return '';
}

function getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
// 清除本場成績(整場移除):還原差點 + 刪成績/技術獎 + 清該場 PDF 連結 + 清本場快取
function clearMatch(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以清除本場成績' };
  var date = String(p && p.date || '').trim();
  if (!date && cfg.lastRanking) { try { date = mdOf_(JSON.parse(cfg.lastRanking).date); } catch (e) {} }
  if (!date) date = latestScoreDate();
  if (!date) return { ok: false, error: '沒有可清除的本場成績' };
  if (isDateLocked_(cfg, date)) return { ok: false, error: '本場已鎖定,請先解鎖再清除' };
  clearMatchForResubmit_(date);                 // 還原差點 + 刪成績 + 刪技術獎 + 清該場 PDF 連結
  var rk = null; if (cfg.lastRanking) { try { rk = JSON.parse(cfg.lastRanking); } catch (e) {} }
  if (!rk || mdOf_(rk.date) === mdOf_(date)) setConfig('lastRanking', '');   // 清本場成績快取
  return { ok: true, date: date };
}

function exportPdf(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以匯出' };
  var date = String(p.date || '').trim() || latestScoreDate();
  if (!date) return { ok: false, error: '沒有可匯出的成績(請先輸入桿數)' };
  var folderName = String(p.folder || '').trim() || cfg.pdfFolder || '例賽報表';
  setConfig('pdfFolder', folderName);                       // 記住資料夾名

  var html = buildMatchReportHtml(date, cfg);
  var blob = Utilities.newBlob(html, MimeType.HTML, 'report.html').getAs(MimeType.PDF);
  var courseName = resolveCourse(date, cfg);
  var fname = date.replace(/\//g, '-') + '_' + (courseName || '球場') + '_' + teamShort(cfg) + '.pdf';
  blob.setName(fname);

  var folder = getOrCreateFolder(folderName);
  var old = folder.getFilesByName(fname);                   // 同名舊檔自動覆蓋
  while (old.hasNext()) old.next().setTrashed(true);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var purl = file.getUrl();
  var links = pdfLinksObj_(cfg); links[mdOf_(date)] = purl;   // 記住該場 PDF 連結
  setConfig('pdfLinks', JSON.stringify(links));
  return { ok: true, date: date, name: fname, folder: folderName,
           url: purl,
           download: 'https://drive.google.com/uc?export=download&id=' + file.getId() };
}

// ---------- 測試資料 ----------
var TEST_PREFIX = '【測】';
function seedTestData(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以建立測試資料' };
  var people = [
    ['【測】陳一', '男', 18, 1960], ['【測】林二', '女', 24, 1965], ['【測】張三', '男', 12, 1958],
    ['【測】王四', '男', 30, 1970], ['【測】李五', '女', 9, 1962], ['【測】趙六', '男', 20, 1955],
    ['【測】周七', '男', 15, 1968], ['【測】吳八', '女', 16, 1972]
  ];
  var ms = sheet(SHEET_MASTER), rs = sheet(SHEET_ROSTER), now = Date.now();
  people.forEach(function (pp, i) {
    ms.appendRow([pp[0], pp[1], pp[2], pp[3]]);
    setHcpBoth(pp[0], pp[2]);
    upsertMember('TEST_' + i, pp[0], '會員', pp[1], String(pp[3]));
    rs.appendRow(['TEST_' + i, pp[0], String(pp[2]), '用餐', '', now + i]);
  });
  upsertMember('TEST_G', '【測】來賓客', '來賓', '男', '');
  rs.appendRow(['TEST_G', '【測】來賓客', '', '用餐', '', now + 99]);
  __masterListCache = __masterMapCache = __rosterCache = null; cacheDel_('master','mastermap','roster','members','hcp');
  return { ok: true, seeded: people.length + 1 };
}
function delByPrefix(sheetName, nameCol) {
  var sh = sheet(sheetName), data = sh.getDataRange().getValues(), removed = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][nameCol] || '').indexOf(TEST_PREFIX) === 0) { sh.deleteRow(i + 1); removed++; }
  }
  return removed;
}
function clearTestData(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以清除測試資料' };
  var total = 0;
  total += delByPrefix(SHEET_ROSTER, 1);
  total += delByPrefix(SHEET_MEMBERS, 1);
  total += delByPrefix(SHEET_MASTER, 0);
  total += delByPrefix(SHEET_HCP, 0);
  total += delByPrefix(SHEET_SCORES, 1);
  total += delByPrefix(SHEET_PAY, 0);
  total += delByPrefix(SHEET_TECH, 2);
  __memInfoCache = __hcpMapCache = __masterListCache = __masterMapCache = __rosterCache = __membersDataCache = null; flushWarmCache();
  return { ok: true, removed: total };
}
function indexByName(arr, name) {
  for (var i = 0; i < arr.length; i++) if (arr[i].name === name) return i;
  return 999;
}
function rankSortForDisplay(a, b) {
  var grp = function (rt) { return rt === '總桿' ? 0 : (rt === '來賓' ? 3 : (rt === '新會員' ? 2 : 1)); };
  if (grp(a.rankType) !== grp(b.rankType)) return grp(a.rankType) - grp(b.rankType);
  return (parseInt(a.rank, 10) || 99) - (parseInt(b.rank, 10) || 99);
}

// 年初差點 = 上年度末差點 × 0.85(四捨五入)
// 年初重設閘門:本年度(seasonFrom~seasonTo)所有「例賽(月例賽/外地賽/國外賽)」都打完後才允許
function seasonResetGate_(cfg) {
  var now = new Date();
  var todayKey = now.getFullYear() * 10000 + (now.getMonth() + 1) * 100 + now.getDate();
  var range = seasonRange(cfg);
  var lastKey = 0;
  if (range) {
    var sK = range.start.getFullYear() * 10000 + (range.start.getMonth() + 1) * 100 + range.start.getDate();
    var eK = range.end.getFullYear() * 10000 + (range.end.getMonth() + 1) * 100 + range.end.getDate();
    getSchedule().forEach(function (e) {
      var isMatch = String(e.type || '').indexOf('賽') >= 0;   // 例賽/月例賽/外地賽/國外賽(交接不含「賽」→不算)
      if (isMatch && e.sortKey >= sK && e.sortKey <= eK && e.sortKey > lastKey) lastKey = e.sortKey;
    });
  }
  var allowed, lastMd = '';
  if (lastKey) {
    allowed = todayKey > lastKey;                 // 最後一場例賽「之後」
    lastMd = (Math.floor(lastKey / 100) % 100) + '/' + (lastKey % 100);
  } else if (range) {
    var eK2 = range.end.getFullYear() * 10000 + (range.end.getMonth() + 1) * 100 + range.end.getDate();
    allowed = todayKey > eK2;                      // 沒排程例賽 → 改用年度迄之後
  } else {
    allowed = true;                                // 無法判斷年度 → 不擋
  }
  return { allowed: allowed, lastMatch: lastMd };
}

function seasonReset(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以執行年初重設' };
  }
  var gate = seasonResetGate_(cfg);
  if (!gate.allowed) {
    return { ok: false, error: '年初重設需等本年度所有例賽結束後才能執行' + (gate.lastMatch ? '(最後一場例賽 ' + gate.lastMatch + ')' : '') + '。' };
  }
  // 倍率:優先用這次傳入的,其次設定值,預設 0.85;限 0~1
  var f = parseFloat(p && p.factor);
  if (isNaN(f)) f = parseFloat(cfg.hcpResetFactor);
  if (isNaN(f) || f <= 0 || f > 1) f = 0.85;
  setConfig('hcpResetFactor', String(f));

  var par = parseFloat(cfg.par); if (isNaN(par)) par = 72;
  var range = seasonRange(cfg);
  var minfo = getMemberInfoMap();
  var n = 0, kept = 0;

  if (!range) {
    // 沒設年度區間 → 退回舊行為:目前差點 × 倍率
    var rmap = {};
    getHcpList().forEach(function (h) {
      var v = parseFloat(h.hcp);
      if (!isNaN(v)) { rmap[h.name] = clampHcp_(v * f, cfg); n++; }
    });
    setHcpBothBatch_(rmap);                 // 一次寫回(原本逐人 setHcpBoth)
  } else {
    // 年初差點 =(去年度平均總桿 − 標準桿)× 倍率
    var sdata = sheet(SHEET_SCORES).getDataRange().getValues();
    var sum = {}, cnt = {};
    for (var i = 1; i < sdata.length; i++) {
      var dstr = String(sdata[i][0] || ''), nm = String(sdata[i][1] || '');
      if (!nm || !inCurrentSeason(dstr, range)) continue;
      var info = minfo[nm];
      if (info && info.role === '來賓') continue;          // 來賓不算差點
      var gross = parseFloat(sdata[i][4]);
      if (isNaN(gross) || gross <= 0) continue;
      sum[nm] = (sum[nm] || 0) + gross; cnt[nm] = (cnt[nm] || 0) + 1;
    }
    // 對每位有差點或有成績的會員重算
    var names = {};
    getHcpList().forEach(function (h) { names[h.name] = true; });
    Object.keys(cnt).forEach(function (k) { names[k] = true; });
    var rmap = {};
    Object.keys(names).forEach(function (nm) {
      var info = minfo[nm];
      if (info && info.role === '來賓') return;
      if (cnt[nm] > 0) {
        var avg = sum[nm] / cnt[nm];
        var nh = clampHcp_((avg - par) * f, cfg);          // (平均−標準桿)×倍率,夾限 [0, 上限]
        rmap[nm] = nh; n++;
      } else {
        kept++;                                            // 去年度沒出賽 → 維持原差點
      }
    });
    setHcpBothBatch_(rmap);                                // 一次寫回所有重算差點(原本逐人 setHcpBoth)
  }

  clearPaid();                                             // 新年度:會費收款全部歸零
  setConfig('seasonStartTs', String(Date.now()));
  return { ok: true, reset: n, kept: kept, factor: f, par: par, hcpList: getHcpList() };
}

// ---------- 年總表(像 Excel 的成績總表)----------
// 用 Scores 資料,在試算表生成/更新「年總表」分頁:每位會員 × 每月成績橫向排開
function buildSeason(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以產生年總表' };
  }

  // 1) 讀成績,只取落在年度區間內的場次,依月份分組
  var range = seasonRange(cfg);
  var sdata = sheet(SHEET_SCORES).getDataRange().getValues();
  // 欄位: date|name|out|in|gross|hcp|net|rankType|rank|hcpAfter|ts
  var dates = [], dateSeen = {}, dateOrder = {};
  var byName = {};   // name -> { date -> {gross, rankType, rank} }
  for (var i = 1; i < sdata.length; i++) {
    var date = String(sdata[i][0] || ''); if (!date) continue;
    var nm = String(sdata[i][1] || ''); if (!nm) continue;
    if (range && !inCurrentSeason(date, range)) continue;     // 只算本年度
    if (!dateSeen[date]) {
      dateSeen[date] = true; dates.push(date);
      dateOrder[date] = range ? scoreFullDate(date, range).getTime() : 0;
    }
    if (!byName[nm]) byName[nm] = {};
    byName[nm][date] = {
      gross: sdata[i][4],
      rankType: String(sdata[i][7] || ''),
      rank: String(sdata[i][8] || '')
    };
  }
  if (range) dates.sort(function (a, b) { return dateOrder[a] - dateOrder[b]; });  // 依年度月份排序

  // 2) 會員清單(依 Members 順序),帶性別/差點/會費
  var mdata = sheet(SHEET_MEMBERS).getDataRange().getValues();
  var hlist = getHcpList(), hmap = {}, smap = {};
  hlist.forEach(function (h) { hmap[h.name] = h.hcp; smap[h.name] = h.seasonStart; });

  var members = [];
  var paidMap = getPaidMap();
  for (var r = 1; r < mdata.length; r++) {
    var nm2 = String(mdata[r][1] || ''); if (!nm2) continue;
    var role = String(mdata[r][2] || '會員');
    if (role === '來賓') continue;                       // 年總表只列會員
    members.push({
      name: nm2,
      gender: String(mdata[r][3] || ''),
      seasonStart: smap[nm2] || '',
      hcp: hmap[nm2] || '',
      fee: feeOf({ role: role, gender: String(mdata[r][3] || '') }),
      paid: paidMap[nm2] ? 'v' : ''
    });
  }

  // 3) 組出表格(2D 陣列)
  var totalRounds = dates.length;   // 本年度總場次
  var head = ['序','姓名','性別','年初差點','目前差點','會費','收款','出賽','出席率','全勤','總桿冠軍','淨桿前三','獎金累計'];
  dates.forEach(function (d) { head.push(d); });
  var rows = [head];

  members.forEach(function (m, idx) {
    var played = 0, grossWins = 0, netTop3 = 0, prize = 0;
    var monthCells = dates.map(function (d) {
      var rec = byName[m.name] && byName[m.name][d];
      if (!rec) return '';
      played++;
      if (rec.rankType === '總桿') grossWins++;
      if (rec.rankType === '淨桿' && (parseInt(rec.rank, 10) || 9) <= 3) netTop3++;
      prize += prizeFor(rec.rankType, rec.rank, cfg);
      var tag = rec.rankType === '總桿' ? '總桿'
              : rec.rankType === '幸運獎' ? '幸運'
              : rec.rankType === '逢五跳獎' ? '逢五'
              : rec.rankType === 'BB' ? 'BB'
              : (rec.rankType === '淨桿' && rec.rank ? '淨' + rec.rank : '');
      return rec.gross + (tag ? ' (' + tag + ')' : '');
    });
    var rate = totalRounds > 0 ? Math.round(played / totalRounds * 100) + '%' : '';
    var perfect = (totalRounds > 0 && played === totalRounds) ? '★全勤' : '';
    rows.push([idx + 1, m.name, m.gender, m.seasonStart, m.hcp, m.fee, m.paid,
               played, rate, perfect, grossWins, netTop3, prize].concat(monthCells));
  });

  // 4) 寫入「年總表」分頁(整片覆蓋)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('年總表');
  if (!sh) sh = ss.insertSheet('年總表');
  sh.clearContents();
  var seasonTxt = (cfg.seasonFrom || cfg.seasonTo) ? ('  年度 ' + (cfg.seasonFrom || '') + '–' + (cfg.seasonTo || '')) : '';
  var title = (cfg.teamName || cfg.titleZh || '球隊') + (cfg.term ? ' 第' + cfg.term + '屆' : '') +
              ' · 年總表' + seasonTxt + '(更新 ' + fmtDate(new Date()) + ')';
  sh.getRange(1, 1).setValue(title);
  if (rows.length && rows[0].length) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(2, 1, 1, rows[0].length).setFontWeight('bold');
    sh.setFrozenRows(2);
    sh.setFrozenColumns(2);
  }

  return { ok: true, members: members.length, months: dates.length };
}

// ===== 成績總表(乾淨版:會員 × 每月總桿,像你提供的圖)=====
function officerTitleMap_(cfg) {
  var m = {};
  if (cfg.president) m[String(cfg.president).trim()] = '會長';
  if (cfg.vicePresident) m[String(cfg.vicePresident).trim()] = '副會長';
  if (cfg.secretary) m[String(cfg.secretary).trim()] = '總幹事';
  if (cfg.treasurer) m[String(cfg.treasurer).trim()] = '財務長';
  return m;
}
// 產生資料:{ dates:[{date,tee,course}], members:[{no,name,title,scores:[...]}] }
function buildScoreMatrix_(cfg) {
  var range = seasonRange(cfg);
  var sdata = sheet(SHEET_SCORES).getDataRange().getValues();
  var seen = {}, dlist = [], dorder = {}, byName = {};
  for (var i = 1; i < sdata.length; i++) {
    var d = mdOf_(sdata[i][0]); if (!d) continue;
    var nm = String(sdata[i][1] || ''); if (!nm) continue;
    if (range && !inCurrentSeason(d, range)) continue;
    if (!seen[d]) { seen[d] = true; dlist.push(d); dorder[d] = range ? scoreFullDate(d, range).getTime() : i; }
    if (!byName[nm]) byName[nm] = {};
    var g = sdata[i][4];
    byName[nm][d] = { g: (g === '' || g == null) ? '' : g, champ: String(sdata[i][7] || '') === '總桿' };
  }
  dlist.sort(function (a, b) { return dorder[a] - dorder[b]; });
  var schMap = {};
  getSchedule().forEach(function (e) { var k = mdOf_(e.date); if (k && !schMap[k]) schMap[k] = { tee: e.tee || '', course: e.course || '' }; });
  var dates = dlist.map(function (d) { var s = schMap[d] || {}; return { date: d, tee: s.tee || '', course: s.course || '' }; });
  var titles = officerTitleMap_(cfg);
  var members = getMasterList().map(function (m, idx) {
    var sc = byName[m.name] || {}, played = 0;
    var cells = dlist.map(function (d) {
      var rec = sc[d];
      if (rec && rec.g !== '') played++;
      return rec ? { g: rec.g, champ: !!rec.champ } : { g: '', champ: false };
    });
    return { no: idx + 1, name: m.name, title: titles[m.name] || '', cells: cells, played: played };
  });
  return { dates: dates, members: members,
           seasonLabel: (cfg.seasonFrom || '') + (cfg.seasonTo ? '–' + cfg.seasonTo : ''),
           term: cfg.term || '', team: cfg.teamName || cfg.titleZh || '球隊' };
}
// 寫入試算表「成績總表」分頁
function buildSeasonMatrix(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) return { ok: false, error: '只有管理員可以產生成績總表' };
  var mx = buildScoreMatrix_(cfg);
  var nCols = 3 + mx.dates.length;   // 編號 + 姓名 + 各場 + 出賽
  var hd1 = ['編號', '姓名'].concat(mx.dates.map(function (x) { return x.date; })).concat(['出賽']);
  var hd2 = ['', ''].concat(mx.dates.map(function (x) { return x.tee; })).concat(['']);
  var hd3 = ['', ''].concat(mx.dates.map(function (x) { return x.course; })).concat(['']);
  var rows = [hd1, hd2, hd3];
  mx.members.forEach(function (m) {
    rows.push([m.no, (m.title || '') + m.name].concat(m.cells.map(function (c) { return c.g; })).concat([m.played]));
  });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('成績總表'); if (!sh) sh = ss.insertSheet('成績總表');
  sh.clearContents(); try { sh.clearFormats(); } catch (e) {}
  var title = mx.team + (cfg.term ? (' 第' + cfg.term + '屆') : '') + '成績總表' +
              (mx.seasonLabel ? ('  ' + mx.seasonLabel) : '') + '(更新 ' + fmtDate(new Date()) + ')';
  sh.getRange(1, 1).setValue(title);
  if (nCols > 0) {
    sh.getRange(2, 1, rows.length, nCols).setValues(rows);
    sh.getRange(2, 1, 3, nCols).setFontWeight('bold');
    sh.setFrozenRows(4);       // 標題 + 三列表頭
    sh.setFrozenColumns(2);    // 編號 + 姓名
    // 總桿冠軍那格上色(資料列從第 5 列開始;場次欄從第 3 欄開始)
    mx.members.forEach(function (m, mi) {
      m.cells.forEach(function (c, di) {
        if (c.champ) sh.getRange(5 + mi, 3 + di).setBackground('#fff1b8');
      });
    });
  }
  return { ok: true, members: mx.members.length, months: mx.dates.length };
}

function fmtDate(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + ' ' +
         (d.getHours() < 10 ? '0' : '') + d.getHours() + ':' +
         (d.getMinutes() < 10 ? '0' : '') + d.getMinutes();
}

// ---------- 整年度例賽行事曆 ----------
var SCHED_HEAD = ['年','月','日期','星期','隊伍','類型','球場','開球時間','預訂狀態',
                  '可預約日','預約提醒','提前天數','提前月數','組數','擊球價','截止日期'];

// 建立「行事曆」分頁並填入範本(若已有資料則不覆蓋,保留你的編輯)
function buildSchedule(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以建立行事曆' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SCHED);
  if (sh && sh.getLastRow() > 1) {
    return { ok: true, existed: true };   // 已有資料,不覆蓋
  }
  if (!sh) sh = ss.insertSheet(SHEET_SCHED);

  var team = '意象太極';
  function D(y, m, d) { return new Date(y, m - 1, d); }
  // 年,月,日期,星期,隊伍,類型,球場,開球時間,預訂狀態,可預約日,預約提醒(''→公式),提前天數,提前月數,組數,擊球價
  var seed = [
    [2026,'7月','7/21','週二',team,'月例賽','長庚','11:00','已訂','', '','','',10,'3,700','6/25(三)18:00'],
    [2026,'8月','8/18','週二',team,'月例賽','新豐','11:00','已訂',D(2026,5,1),'','',3,8,'',''],
    [2026,'9月','9/15','週二',team,'月例賽','黃金海岸','11:00','已訂','', '','','',10,'',''],
    [2026,'10月','10/20','週二',team,'月例賽','龍潭','','未訂',D(2026,8,1),'','',2,10,'',''],
    [2026,'11月','11/17','週二',team,'外地賽','全國','10:53','已訂',D(2026,5,1),'','',6,10,'',''],
    [2026,'11月','11/18','週三',team,'外地賽','皇家','10:10','已訂','', '','','',10,'2,500',''],
    [2026,'12月','12/15','週二',team,'月例賽','旭陽','11:50','已訂',D(2026,6,1),'','',6,8,'',''],
    [2027,'1月','1/19','週二',team,'月例賽','大屯/國華','','未訂','', '','','',10,'',''],
    [2027,'2月','2/16','週二',team,'月例賽','老淡水','','未訂',D(2027,1,16),'','',1,10,'',''],
    [2027,'3月','3/16','週二',team,'國外賽','國外賽','','未訂','', '','','','','',''],
    [2027,'4月','4/20','週二',team,'月例賽','老爺/山溪地','','未訂','', '','','',10,'',''],
    [2027,'5月','5/18','週二',team,'月例賽','礁溪/楊梅','','未訂','', '','','',10,'',''],
    [2027,'6月','6/15','週二',team,'交接','長庚','','未訂','', '','','',10,'','']
  ];

  sh.clearContents();
  sh.getRange(1, 1, 1, SCHED_HEAD.length).setValues([SCHED_HEAD]).setFontWeight('bold');
  sh.getRange(2, 1, seed.length, SCHED_HEAD.length).setValues(seed);
  // 預約提醒(第11欄)做成公式,依「可預約日 vs 今天」自動算
  for (var r = 0; r < seed.length; r++) {
    var row = r + 2;
    sh.getRange(row, 11).setFormula(
      '=IF($I' + row + '="已訂","✅ 已預訂",IF($J' + row + '="","未訂",' +
      'IF($J' + row + '>TODAY(),"尚未開放(剩 "&TEXT($J' + row + '-TODAY(),"0")&" 天)","✅ 可預約了")))'
    );
  }
  sh.setFrozenRows(1);
  sh.getRange(2, 10, seed.length, 1).setNumberFormat('yyyy/m/d');
  __schedCache = null; cacheDel_('sched');
  return { ok: true, rows: seed.length };
}

// 讀行事曆給 App 顯示,並即時算預約提醒
// 來源可設定:Config 的 schedSheetId(外部試算表,選填)、schedSheetName(分頁名,預設「行事曆」)
// 以「欄位標題」對應,所以你的表格欄位順序不同、有多餘欄位也能讀
function getSchedule() {
  if (__schedCache) return __schedCache;
  var cachedSch = cacheGet_('sched');
  if (cachedSch) { __schedCache = cachedSch; return cachedSch; }
  var cfg = getConfig();
  var ssTarget;
  if (cfg.schedSheetId) {
    try { ssTarget = SpreadsheetApp.openById(cfg.schedSheetId); }
    catch (e) { ssTarget = SpreadsheetApp.getActiveSpreadsheet(); }
  } else {
    ssTarget = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sh = ssTarget.getSheetByName(cfg.schedSheetName || SHEET_SCHED);
  if (!sh || sh.getLastRow() < 2) { cachePut_('sched', []); return (__schedCache = []); }

  var data = sh.getDataRange().getValues();
  var header = data[0];
  function ci(name) {                       // 找欄位:標題去空白後完全相符
    for (var c = 0; c < header.length; c++) {
      if (String(header[c]).replace(/\s/g, '') === name) return c;
    }
    return -1;
  }
  function cell(row, name) { var c = ci(name); return c < 0 ? '' : row[c]; }

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var rawDate = cell(data[i], '日期');
    var yearCell = cell(data[i], '年'), monthCell = cell(data[i], '月');
    if ((yearCell === '' || yearCell == null) && (rawDate === '' || rawDate == null)) continue;
    var nd = normSchedDate(yearCell, monthCell, rawDate);
    var status = String(cell(data[i], '預訂狀態') || '');
    var teeCell = cell(data[i], '開球時間');
    var dlCell = cell(data[i], '截止日期');
    var deadline = (dlCell instanceof Date)
      ? ((dlCell.getMonth() + 1) + '/' + dlCell.getDate() + ((dlCell.getHours() || dlCell.getMinutes()) ? (' ' + fmtClock_(dlCell)) : ''))
      : String(dlCell || '');
    out.push({
      year: String(nd.y || ''),
      month: String(nd.mo || monthCell || ''),
      date: nd.disp,
      dow: String(cell(data[i], '星期') || ''),
      type: String(cell(data[i], '類型') || ''),
      course: String(cell(data[i], '球場') || ''),
      tee: (teeCell instanceof Date ? fmtClock_(teeCell) : String(teeCell || '')),
      status: status,
      groups: String(cell(data[i], '組數') || ''),
      price: String(cell(data[i], '擊球價') || ''),
      deadline: deadline,
      reminder: schedReminder(status, cell(data[i], '可預約日')),
      sortKey: nd.y * 10000 + nd.mo * 100 + nd.da
    });
  }
  out.sort(function (a, b) { return a.sortKey - b.sortKey; });
  cachePut_('sched', out);
  return (__schedCache = out);
}

// 把行事曆日期解析成 {y, mo, da, disp},容許:日期物件、「7/21」、「2026/7/21」、或只填日號(月取「月」欄)
function normSchedDate(yearCell, monthCell, dateCell) {
  if (dateCell instanceof Date) {
    return { y: dateCell.getFullYear(), mo: dateCell.getMonth() + 1, da: dateCell.getDate(),
             disp: (dateCell.getMonth() + 1) + '/' + dateCell.getDate() };
  }
  var s = String(dateCell == null ? '' : dateCell).trim();
  var y = parseInt(yearCell, 10) || 0;
  if (s.indexOf('/') >= 0) {
    var p = s.split('/');
    if (p.length >= 3) {                       // 2026/7/21
      y = parseInt(p[0], 10) || y;
      var mo3 = parseInt(p[1], 10) || 0, da3 = parseInt(p[2], 10) || 0;
      return { y: y, mo: mo3, da: da3, disp: mo3 + '/' + da3 };
    }
    var mo = parseInt(p[0], 10) || 0, da = parseInt(p[1], 10) || 0;   // 7/21
    return { y: y, mo: mo, da: da, disp: mo + '/' + da };
  }
  var mo2 = parseInt(monthCell, 10) || 0, da2 = parseInt(s, 10) || 0; // 只填日號
  return { y: y, mo: mo2, da: da2, disp: (mo2 && da2) ? (mo2 + '/' + da2) : s };
}

function schedReminder(status, openRaw) {
  if (status === '已訂') return '✅ 已預訂';
  if (openRaw === '' || openRaw === null || openRaw === undefined) return '未訂';
  var d = (openRaw instanceof Date) ? openRaw : new Date(openRaw);
  if (isNaN(d.getTime())) return '未訂';
  var t = new Date(); t.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0);
  var days = Math.round((d - t) / 86400000);
  if (days > 0) return '尚未開放(剩 ' + days + ' 天)';
  return '✅ 可預約了';
}
// 由 年 + 日期「M/D」組出可排序數值
// ---------- 開放時間鎖(分組 08:00、成績 18:00)----------
// 比賽日:優先用 cfg.matchDate(YYYY/M/D);沒設就抓行事曆「下一場」的日期
function getMatchDate(cfg) {
  if (cfg.matchDate) {
    var m = String(cfg.matchDate).split(/[\/\-\.]/);
    if (m.length >= 3) {
      var d = new Date(parseInt(m[0], 10), parseInt(m[1], 10) - 1, parseInt(m[2], 10));
      if (!isNaN(d.getTime())) return d;
    }
  }
  var sched = getSchedule();
  var cutKey = graceCutoffKey_(cfg);
  for (var i = 0; i < sched.length; i++) {
    if (sched[i].sortKey >= cutKey) {
      var k = sched[i].sortKey;
      return new Date(Math.floor(k / 10000), Math.floor((k % 10000) / 100) - 1, k % 100);
    }
  }
  return null;
}
function hm(str, defH, defM) {
  var t = parseTime(str);                  // 回傳分鐘數或 null
  if (t === null) return { h: defH, m: defM };
  return { h: Math.floor(t / 60), m: t % 60 };
}
// 回傳分組/成績是否已開放,以及鎖定時的提示文字
function revealState(cfg) {
  var md = getMatchDate(cfg);
  if (!md) return { groupsOpen: true, scoreOpen: true };  // 無法判定日期 → 不鎖
  var g = hm(cfg.groupRevealAt, 8, 0), s = hm(cfg.scoreRevealAt, 18, 0);
  var gT = new Date(md.getFullYear(), md.getMonth(), md.getDate(), g.h, g.m);
  var sT = new Date(md.getFullYear(), md.getMonth(), md.getDate(), s.h, s.m);
  var now = new Date();
  var ds = (md.getMonth() + 1) + '/' + md.getDate();
  return {
    groupsOpen: now >= gT,
    scoreOpen: now >= sT,
    groupLabel: ds + ' ' + pad2(g.h) + ':' + pad2(g.m) + ' 開放查看分組',
    scoreLabel: ds + ' ' + pad2(s.h) + ':' + pad2(s.m) + ' 開放查看成績'
  };
}
function pad2(n) { return (n < 10 ? '0' : '') + n; }

// 報名頁顯示用:管理員留空的日期/球場/開球時間,自動帶行事曆下一場;填了則以手動為準
// 賽後寬限:把今天往前推 N 天當作挑「目前場次」的基準
// → 比賽日起算 N 天內,仍把剛打完的這場當作目前場次(報名卡/成績/分組時間鎖都留在本場)
function graceCutoffKey_(cfg) {
  var grace = parseInt(cfg && cfg.eventGraceDays, 10); if (isNaN(grace)) grace = 3;
  var now = new Date();
  var cut = new Date(now.getFullYear(), now.getMonth(), now.getDate() - grace);
  return cut.getFullYear() * 10000 + (cut.getMonth() + 1) * 100 + cut.getDate();
}
function effectiveEvent(cfg, schedule) {
  var next = null;
  var cutKey = graceCutoffKey_(cfg);
  (schedule || []).forEach(function (e) {
    if (!next && e.sortKey >= cutKey) next = e;
  });
  // 行事曆下一場優先;手動欄位只在行事曆沒有對應值時當備用
  function pick(manual, auto) {
    return (auto !== undefined && String(auto).trim() !== '') ? auto : (manual || '');
  }
  return {
    titleZh: cfg.titleZh || '',
    date: pick(cfg.date, next ? (next.date + (next.dow ? '(' + next.dow.replace('週', '') + ')' : '')) : ''),
    course: pick(cfg.course, next ? next.course : ''),
    tee: pick(cfg.tee, next ? next.tee : ''),
    fee: pick(cfg.fee, next ? next.price : ''),                         // 費用 ← 行事曆「擊球價」(沒填才用設定)
    cap: (next && next.groups && parseInt(next.groups, 10) > 0)          // 名額(人數)← 行事曆「組數」×4(每組4人)
           ? String(parseInt(next.groups, 10) * 4) : (cfg.cap || ''),
    deadline: pick(cfg.deadline, next ? next.deadline : ''),            // 截止 ← 行事曆「截止日期」(沒填才用設定)
    type: next ? next.type : '',
    groups: next ? next.groups : '',
    price: next ? next.price : '',
    key: next ? next.sortKey : 0,                                      // 目前場次的排序鍵(換場偵測用)
    fromSchedule: !!next
  };
}

// 換場自動處理:偵測「目前場次」是否已前進到新的一場 → 封存上一場報名 + 清空名單
// 在 bootstrap 讀名單前呼叫;常態(沒換場)只做比對、不寫入
function rolloverRosterIfSwitched_(cfg) {
  var ev = effectiveEvent(cfg, getSchedule());
  var curKey = ev && ev.key ? ev.key : 0;
  if (!curKey) return;                                    // 沒有有效場次 → 不動
  var storedKey = parseInt(cfg.rosterEventKey, 10) || 0;
  if (!storedKey) {                                       // 第一次:只記錄,不清空
    setConfig('rosterEventKey', String(curKey));
    setConfig('rosterEventDate', ev.date || '');
    return;
  }
  if (curKey <= storedKey) return;                        // 沒前進到新場次 → 不動
  var lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return; }      // 拿不到鎖就跳過,下次 bootstrap 再處理
  try {
    var cfg2 = getConfig();
    var storedKey2 = parseInt(cfg2.rosterEventKey, 10) || 0;
    if (curKey <= storedKey2) return;                     // 鎖內再確認,避免多人同時重複封存
    archiveRoster_(cfg2.rosterEventDate || '');           // 封存「上一場」名單
    var sh = sheet(SHEET_ROSTER), last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);             // 清空名單
    setConfig('noticeEvent', '');                         // 換月一併清掉上一場的注意事項
    setConfig('rosterEventKey', String(curKey));
    setConfig('rosterEventDate', ev.date || '');
    __rosterCache = null; cacheDel_('roster');
  } finally { lock.releaseLock(); }
}

// 把目前名單封存到「報名紀錄」分頁(每列加上場次日期),供分月留存查閱
function archiveRoster_(eventDate) {
  var sh = sheet(SHEET_ROSTER), data = sh.getDataRange().getValues();
  if (data.length < 2) return;                            // 空名單不封存
  var arch = sheet(SHEET_ROSTER_ARCHIVE), rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // Roster: userId|name|hcp|cart|note|ts|status → 封存欄序:場次|姓名|狀態|用餐|差點|備註|userId|報名時間
    rows.push([eventDate || '', String(data[i][1] || ''), String(data[i][6] || '報名'),
               String(data[i][3] || ''), String(data[i][2] || ''), String(data[i][4] || ''),
               String(data[i][0] || ''), data[i][5] ? new Date(Number(data[i][5])) : '']);
  }
  if (rows.length) arch.getRange(arch.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// 解析「姓名 OUT IN」,容許逗號/空白/全形
function parseScoreLines(text) {
  return String(text || '').split(/[\r\n]+/).map(function (line) {
    var t = line.trim();
    if (!t) return null;
    var m = t.split(/[\s,，、\t]+/).filter(function (x) { return x !== ''; });
    if (m.length < 3) return null;
    var inn = parseFloat(m.pop()), out = parseFloat(m.pop());
    var name = m.join(' ').trim();
    if (!name || isNaN(out) || isNaN(inn)) return null;
    return { name: name, out: out, in: inn };
  }).filter(function (x) { return x; });
}

// 把「姓名 數值」每行拆成 [name, value](用於差點匯入)
function parseLines(text) {
  return String(text || '').split(/[\r\n]+/).map(function (line) {
    var t = line.trim();
    if (!t) return null;
    var m = t.split(/[\s,，、\t]+/);
    var val = m.pop();
    var name = m.join(' ').trim();
    return name ? [name, val] : null;
  }).filter(function (x) { return x; });
}
function round1(n) { return Math.round(n * 10) / 10; }

function parseTime(s) {
  var m = String(s).match(/^(\d{1,2})[:：](\d{2})$/);
  if (!m) return null;
  var h = parseInt(m[1], 10), mi = parseInt(m[2], 10);
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}
function fmtTime(mins) {
  mins = ((mins % 1440) + 1440) % 1440;
  var h = Math.floor(mins / 60), mi = mins % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mi < 10 ? '0' : '') + mi;
}

// ---------- 資料存取 ----------

function sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SHEET_ROSTER)  sh.appendRow(['userId','name','hcp','cart','note','ts','status']);
    if (name === SHEET_MEMBERS) sh.appendRow(['userId','name','role','gender','birthYear','phone','lineName']);
    if (name === SHEET_TECH)    sh.appendRow(['日期','獎項','得獎者','條數']);
    if (name === SHEET_MASTER)  sh.appendRow(['姓名','性別','年初差點','出生年']);
    if (name === SHEET_CONFIG)  sh.appendRow(['key','value']);
    if (name === SHEET_HCP)     sh.appendRow(['name','hcp','seasonStart']);
    if (name === SHEET_SCORES)  sh.appendRow(['date','name','out','in','gross','hcp','net','rankType','rank','hcpAfter','ts']);
    if (name === SHEET_PAY)     sh.appendRow(['name','paid']);
    if (name === SHEET_ROSTER_ARCHIVE) sh.appendRow(['場次','姓名','狀態','用餐','差點','備註','userId','報名時間']);
  }
  return sh;
}

// ---- 差點資料層(以姓名為鍵)----
function getHcpMap() {
  if (__hcpMapCache) return __hcpMapCache;
  var cachedHM = cacheGet_('hcp');
  if (cachedHM) { __hcpMapCache = cachedHM; return cachedHM; }
  var data = sheet(SHEET_HCP).getDataRange().getValues(), map = {};
  for (var i = 1; i < data.length; i++) {
    var v = data[i][1];
    if (v === '' || v === null || v === undefined) v = data[i][2];   // 目前差點(第2欄)空 → 退回期初差點(第3欄)
    if (data[i][0] !== '' && v !== '' && v !== null && v !== undefined) map[String(data[i][0])] = Number(v);
  }
  // 最後一層退路:Handicaps 沒有這個人 → 用「會員名單」的年初差點(還沒匯入/還沒比賽時也看得到)
  var master = getMasterList();
  for (var j = 0; j < master.length; j++) {
    var mn = master[j].name, mh = master[j].hcp;
    if (mn && !map.hasOwnProperty(mn) && mh !== '' && mh !== null && mh !== undefined) {
      var nh = Number(mh);
      if (!isNaN(nh)) map[mn] = nh;
    }
  }
  __hcpMapCache = map;
  cachePut_('hcp', map);
  return map;
}
function getHcpList() {
  var data = sheet(SHEET_HCP).getDataRange().getValues(), out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') out.push({
      name: String(data[i][0]),
      hcp: String(data[i][1]),
      seasonStart: String(data[i][2] !== undefined && data[i][2] !== '' ? data[i][2] : data[i][1])
    });
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return out;
}
// setHcp 只更新「目前差點」(第2欄),季初差點(第3欄)保留
function setHcp(name, hcp) {
  var sh = sheet(SHEET_HCP), data = sh.getDataRange().getValues();
  __hcpMapCache = null; cacheDel_('hcp');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) { sh.getRange(i + 1, 2).setValue(hcp); return; }
  }
  sh.appendRow([name, hcp, hcp]);
}
// setSeasonStart 同時設定季初與目前(匯入/季初重設時用)
function setHcpBoth(name, hcp) {
  var sh = sheet(SHEET_HCP), data = sh.getDataRange().getValues();
  __hcpMapCache = null; cacheDel_('hcp');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      sh.getRange(i + 1, 2, 1, 2).setValues([[hcp, hcp]]); return;
    }
  }
  sh.appendRow([name, hcp, hcp]);
}

function getRoster() {
  if (__rosterCache) return __rosterCache;
  var cachedR = cacheGet_('roster');
  if (cachedR) { __rosterCache = cachedR; return cachedR; }
  var data = sheet(SHEET_ROSTER).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      userId: String(data[i][0]),
      name:   String(data[i][1]),
      hcp:    String(data[i][2] || ''),
      cart:   String(data[i][3] || ''),
      note:   String(data[i][4] || ''),
      ts:     Number(data[i][5] || 0),
      status: String(data[i][6] || '報名')          // 報名 / 請假
    });
  }
  out.sort(function (a, b) { return a.ts - b.ts; });   // 依報名先後
  cachePut_('roster', out);
  return (__rosterCache = out);
}

function getMember(userId) {
  if (!userId) return null;
  var data = membersData_();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      return {
        name: String(data[i][1]),
        role: String(data[i][2] || '會員'),
        gender: String(data[i][3] || ''),
        birthYear: String(data[i][4] || ''),
        phone: String(data[i][5] || ''),
        lineName: String(data[i][6] || '')
      };
    }
  }
  return null;
}

function upsertMember(userId, name, role, gender, birth, phone, lineName) {
  var sh = sheet(SHEET_MEMBERS);
  __memInfoCache = __membersDataCache = null; cacheDel_('members');
  var data = sh.getDataRange().getValues();
  var row = [userId, name, role || '會員', gender || '', birth || '', phone || '', lineName || ''];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      if (phone === undefined)    row[5] = (data[i][5] != null ? data[i][5] : '');  // 未帶電話時保留原值
      if (lineName === undefined) row[6] = (data[i][6] != null ? data[i][6] : '');  // 未帶 LINE 名稱時保留原值
      sh.getRange(i + 1, 1, 1, 7).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

// 姓名 -> {role, gender, birthYear}
function getMemberInfoMap() {
  if (__memInfoCache) return __memInfoCache;
  var map = {};
  // 先用「會員名單」打底:未綁定的會員也有 role(會員)/性別/出生年 → 修正淨桿平手比年齡、二男二女分組
  getMasterList().forEach(function (m) {
    if (m.name) map[m.name] = { role: '會員', gender: m.gender || '', birthYear: m.birthYear || '' };
  });
  // 再用 Members(綁定/來賓)覆蓋:role 以 Members 為準(來賓在此標記);性別/出生年空白時保留會員名單值
  var data = membersData_();
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][1] || '');
    if (!nm) continue;
    var base = map[nm] || {};
    map[nm] = { role: String(data[i][2] || '會員'),
                gender: String(data[i][3] || '') || base.gender || '',
                birthYear: String(data[i][4] || '') || base.birthYear || '' };
  }
  __memInfoCache = map;
  return map;
}

// 姓名 -> 性別(男/女),供前端文字上色

// ---------- 會員名單主檔 ----------
// 姓名 -> {gender, hcp, birthYear}
function getMasterMap() {
  if (__masterMapCache) return __masterMapCache;
  var cachedMM = cacheGet_('mastermap');
  if (cachedMM) { __masterMapCache = cachedMM; return cachedMM; }
  var data = sheet(SHEET_MASTER).getDataRange().getValues(), map = {};
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][0] || '').trim();
    if (nm) map[nm] = { gender: String(data[i][1] || ''),
                       hcp: String(data[i][2] || ''),
                       birthYear: String(data[i][3] || '') };
  }
  __masterMapCache = map;
  cachePut_('mastermap', map);
  return map;
}
function getMasterList() {
  if (__masterListCache) return __masterListCache;
  var cachedML = cacheGet_('master');
  if (cachedML) { __masterListCache = cachedML; return cachedML; }
  var data = sheet(SHEET_MASTER).getDataRange().getValues(), out = [];
  // 以標題列偵測「會員身份」欄(放哪一欄都可)
  var idIdx = -1;
  if (data.length) {
    for (var c = 0; c < data[0].length; c++) {
      var hh = String(data[0][c] || '');
      if (hh.indexOf('身份') >= 0 || hh.indexOf('身分') >= 0) { idIdx = c; break; }
    }
  }
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][0] || '').trim();
    if (nm) out.push({ name: nm, gender: String(data[i][1] || ''),
                      hcp: String(data[i][2] || ''), birthYear: String(data[i][3] || ''),
                      memberType: idIdx >= 0 ? String(data[i][idIdx] || '').trim() : '' });
  }
  __masterListCache = out;
  cachePut_('master', out);
  return out;
}

// 匯入會員名單主檔。每行「姓名 [性別] [差點] [出生年]」,順序不拘:
// 男/女=性別、四位數(>=1900)=出生年、其他數字=差點
function importMembers(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以匯入會員名單' };
  }
  var rows = String(p.text || '').split(/[\r\n]+/), parsed = [];
  rows.forEach(function (line) {
    var t = line.trim(); if (!t) return;
    var toks = t.split(/[\s,，、\t]+/).filter(function (x) { return x !== ''; });
    if (!toks.length) return;
    var name = toks.shift();
    var gender = '', hcp = '', birth = '';
    toks.forEach(function (tk) {
      if (tk === '男' || tk === '女') gender = tk;
      else if (/^\d+$/.test(tk)) {
        var num = parseInt(tk, 10);
        if (num >= 1900 && tk.length === 4) birth = tk;     // 四位數當出生年
        else hcp = tk;                                       // 其餘數字當差點
      }
    });
    if (name) parsed.push([name, gender, hcp, birth]);
  });

  // 寫入主檔(整片覆蓋:姓名/性別/差點/出生年)
  var sh = sheet(SHEET_MASTER);
  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['姓名','性別','年初差點','出生年']]).setFontWeight('bold');
  if (parsed.length) sh.getRange(2, 1, parsed.length, 4).setValues(parsed);
  sh.setFrozenRows(1);
  __masterListCache = __masterMapCache = null; cacheDel_('master','mastermap');

  // 有填差點的,順便種年初差點(季初=目前=該值)→ 一次寫回
  var hmap = {};
  parsed.forEach(function (r) {
    var hv = parseFloat(r[2]);
    if (r[0] && !isNaN(hv)) hmap[r[0]] = round1(hv);
  });
  setHcpBothBatch_(hmap);

  return { ok: true, imported: parsed.length, binding: getBindingStatus() };
}

// 綁定狀態:主檔每個人是否已有 LINE 帳號綁定(Members 有同名且有 userId)
function getBindingStatus() {
  var master = getMasterList();
  var mdata = sheet(SHEET_MEMBERS).getDataRange().getValues();
  var bound = {}, phoneOf = {}, birthOf = {}, lineOf = {}, typeByName = {};
  for (var i = 1; i < mdata.length; i++) {
    var uid = String(mdata[i][0] || ''), nm = String(mdata[i][1] || '');
    if (uid && nm) { bound[nm] = true; phoneOf[nm] = String(mdata[i][5] || ''); birthOf[nm] = String(mdata[i][4] || ''); lineOf[nm] = String(mdata[i][6] || ''); }
  }
  master.forEach(function (m) { typeByName[m.name] = m.memberType || ''; });
  var list = master.map(function (m) {
    return { name: m.name, gender: m.gender, hcp: m.hcp, bound: !!bound[m.name], phone: phoneOf[m.name] || '', lineName: lineOf[m.name] || '', memberType: m.memberType || '' };
  });
  var boundCount = list.filter(function (x) { return x.bound; }).length;
  // 額外:已綁定但不在主檔的(通常是來賓)
  var extra = [];
  for (var j = 1; j < mdata.length; j++) {
    var u2 = String(mdata[j][0] || ''), n2 = String(mdata[j][1] || ''), role = String(mdata[j][2] || '');
    if (u2 && n2 && !master.some(function (m) { return m.name === n2; })) {
      extra.push({ name: n2, role: role || '來賓', phone: String(mdata[j][5] || ''), lineName: String(mdata[j][6] || '') });
    }
  }
  return { list: list, total: list.length, bound: boundCount, extra: extra };
}

// 會費:會員依性別,女 10000、男 11000(可由 Config feeFemale / feeMale 調整);來賓無年費
function feeOf(member) {
  if (!member || member.role === '來賓') return '';
  var cfg = getConfig();
  if (member.gender === '女') return parseInt(cfg.feeFemale, 10) || 10000;
  if (member.gender === '男') return parseInt(cfg.feeMale, 10) || 11000;
  return '';
}

// 把目前所有會員(非來賓)列入 Payments 分頁,保留已打的 v,讓管理員直接在試算表編輯
function ensurePayList(userId) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以整理收款名單' };
  }
  var sh = sheet(SHEET_PAY);
  var existing = {};
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var nm = String(data[i][0] || '');
    if (nm) existing[nm] = true;
  }
  var mdata = sheet(SHEET_MEMBERS).getDataRange().getValues();
  var added = 0, count = 0, newRows = [];
  for (var j = 1; j < mdata.length; j++) {
    var name = String(mdata[j][1] || '');
    var role = String(mdata[j][2] || '會員');
    if (!name || role === '來賓') continue;
    count++;
    if (!existing[name]) { newRows.push([name, '']); existing[name] = true; added++; }
  }
  if (newRows.length) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);  // 一次補齊
  return { ok: true, count: count, added: added };
}

// ---------- 會費收款(本年度)----------
function getPaidMap() {
  var data = sheet(SHEET_PAY).getDataRange().getValues(), map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '') map[String(data[i][0])] = (String(data[i][1]) === 'v');
  }
  return map;
}
function setPaid(name, paid) {
  var sh = sheet(SHEET_PAY), data = sh.getDataRange().getValues();
  var val = paid ? 'v' : '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) { sh.getRange(i + 1, 2).setValue(val); return; }
  }
  sh.appendRow([name, val]);
}
function clearPaid() {
  var sh = sheet(SHEET_PAY), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 2, last - 1, 1).clearContent();   // 只清收款欄,保留姓名
}
// 會員的會費/收款清單(給管理員勾選用)
function getMemberFees() {
  var mdata = sheet(SHEET_MEMBERS).getDataRange().getValues();
  var paid = getPaidMap(), out = [], paidSum = 0, total = 0;
  for (var i = 1; i < mdata.length; i++) {
    var name = String(mdata[i][1] || ''); if (!name) continue;
    var role = String(mdata[i][2] || '會員');
    if (role === '來賓') continue;
    var gender = String(mdata[i][3] || '');
    var fee = feeOf({ role: role, gender: gender });
    var isPaid = !!paid[name];
    out.push({ name: name, gender: gender, fee: String(fee), paid: isPaid });
    if (typeof fee === 'number') { total += fee; if (isPaid) paidSum += fee; }
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return { list: out, total: total, paidSum: paidSum };
}
function setPaidAction(userId, p) {
  var cfg = getConfig();
  if (!userId || userId !== cfg.adminUserId) {
    return { ok: false, error: '只有管理員可以登記收款' };
  }
  var name = (p.name || '').toString().trim();
  if (!name) return { ok: false, error: '缺少姓名' };
  setPaid(name, String(p.paid) === '1' || p.paid === true);
  return { ok: true, memberFees: getMemberFees() };
}

// 把試算表存成日期/時間物件的設定值,依欄位格式化成乾淨字串
function fmtClock_(d) { var h = d.getHours(), m = d.getMinutes(); return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m; }
function fmtConfigVal_(key, v) {
  if (v instanceof Date) {
    if (key === 'seasonFrom' || key === 'seasonTo') return v.getFullYear() + '/' + (v.getMonth() + 1);
    if (key === 'groupRevealAt' || key === 'scoreRevealAt' || key === 'tee') return fmtClock_(v);
    if (key === 'matchDate') return v.getFullYear() + '/' + (v.getMonth() + 1) + '/' + v.getDate();
    if (key === 'date') return (v.getMonth() + 1) + '/' + v.getDate();
    return v.getFullYear() + '/' + (v.getMonth() + 1) + '/' + v.getDate();
  }
  return String(v == null ? '' : v);
}
var __cfgCache = null;   // 請求內快取:同一次請求多次 getConfig 不重覆讀表(寫入時清空)
var __memInfoCache = null, __hcpMapCache = null, __masterListCache = null, __masterMapCache = null; // 請求內快取(寫入對應表時清空)
var __schedCache = null, __rosterCache = null;   // 行事曆/報名清單 請求內快取(寫入對應表時清空)
var __membersDataCache = null;                   // Members 原始列 請求內快取(getMember 與 getMemberInfoMap 共用,寫入時清空)
// Members 分頁原始列(同一次請求只讀一次;getMember/getMemberInfoMap 共用)
function membersData_() {
  if (__membersDataCache) return __membersDataCache;
  var cachedMD = cacheGet_('members');
  if (cachedMD) { __membersDataCache = cachedMD; return cachedMD; }
  var d = sheet(SHEET_MEMBERS).getDataRange().getValues();
  cachePut_('members', d);
  return (__membersDataCache = d);
}
// ===== 跨請求暖快取(CacheService,所有人共用;App 寫入即時失效,手動改表需等 TTL 或按「清除快取」)=====
var CACHE_TTL = 600;   // 共享快取存活秒數(預設 10 分;足以涵蓋賽前大量登入,又能讓手動改表 10 分內自癒)
function _scache(){ try { return CacheService.getScriptCache(); } catch (e) { return null; } }
function cacheGet_(key){ var c=_scache(); if(!c) return null; try { var v=c.get(key); return v?JSON.parse(v):null; } catch(e){ return null; } }
function cachePut_(key,obj){ var c=_scache(); if(!c) return; try { c.put(key, JSON.stringify(obj), CACHE_TTL); } catch(e){} }
function cacheDel_(){ var c=_scache(); if(!c) return; try { c.removeAll(Array.prototype.slice.call(arguments)); } catch(e){} }
function flushWarmCache(){ cacheDel_('cfg','sched','master','mastermap','hcp','members','roster'); }

function getConfig() {
  if (__cfgCache) return __cfgCache;
  var cached = cacheGet_('cfg');
  if (cached) { __cfgCache = cached; return cached; }
  var sh = sheet(SHEET_CONFIG);
  var data = sh.getDataRange().getValues();
  var cfg = {};
  Object.keys(DEFAULT_CONFIG).forEach(function (k) { cfg[k] = DEFAULT_CONFIG[k]; });
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) cfg[String(data[i][0])] = fmtConfigVal_(String(data[i][0]), data[i][1]);
  }
  // 第一次跑時把預設值寫進表裡,方便之後直接在試算表編輯
  if (data.length <= 1) {
    Object.keys(DEFAULT_CONFIG).forEach(function (k) {
      sh.appendRow([k, DEFAULT_CONFIG[k]]);
    });
  }
  __cfgCache = cfg;
  cachePut_('cfg', cfg);
  return cfg;
}

function setConfig(key, value) {
  var sh = sheet(SHEET_CONFIG);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(value);
      if (__cfgCache) __cfgCache[key] = fmtConfigVal_(key, value);   // 就地更新快取(免同請求重讀)
      cacheDel_('cfg');
      return;
    }
  }
  sh.appendRow([key, value]);
  if (__cfgCache) __cfgCache[key] = fmtConfigVal_(key, value);       // 就地更新快取
  cacheDel_('cfg');
}
