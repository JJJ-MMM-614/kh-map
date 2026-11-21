/**
 * 萬事通建案查詢 LINE 機器人（Google Apps Script，綁在 Google 試算表上）
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   LINE_TOKEN = Channel access token
 *   JOIN_CODE  = 員工加入驗證碼（例如 8 位數字，自訂）
 * 試算表分頁會自動建立：「員工名單」「查詢紀錄」
 */
var DATA = 'https://jjj-mmm-614.github.io/kh-map/bot/';
var MAP = 'https://jjj-mmm-614.github.io/kh-map/';
var BG = '#0B1A2A', PANEL = '#12293D', CY = '#22D3EE', OR = '#FB923C', PU = '#C084FC', TX = '#E6F1FA', SUB = '#8FA8BF';
var CN = ['預售', '一手', '轉手'], COL = [CY, OR, PU];

function prop(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function doGet() { return ContentService.createTextOutput('ok'); }

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    (body.events || []).forEach(handle);
  } catch (err) { console.error(err); }
  return ContentService.createTextOutput('ok');
}

// ---------- 名單 ----------
function sheet(name, header) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function staff() { return sheet('員工名單', ['LINE ID（勿改）', '姓名', '狀態（啟用/停用）', '加入時間']); }
function findStaff(uid) {
  var v = staff().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === uid) return { row: i + 1, name: v[i][1], on: String(v[i][2]).trim() === '啟用' };
  return null;
}
function profileName(uid) {
  try {
    var r = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + uid, { headers: { Authorization: 'Bearer ' + prop('LINE_TOKEN') } });
    return JSON.parse(r.getContentText()).displayName;
  } catch (e) { return ''; }
}
function log(name, q, result) {
  sheet('查詢紀錄', ['時間', '姓名', '輸入', '結果']).appendRow([new Date(), name, q, result]);
}

// ---------- 主流程 ----------
function handle(ev) {
  var uid = ev.source && ev.source.userId;
  if (!uid || !ev.replyToken) return;
  if (ev.type === 'follow') return reply(ev.replyToken, [txt('您好！這是內部建案查詢。\n第一次使用請輸入主管給您的驗證碼。')]);
  var text = ev.type === 'message' && ev.message.type === 'text' ? ev.message.text.trim() : null;
  var pick = ev.type === 'postback' ? ev.postback.data : null;
  if (text === null && pick === null) return;

  var s = findStaff(uid);
  if (!s) {
    if (text && prop('JOIN_CODE') && text === prop('JOIN_CODE')) {
      var nm = profileName(uid);
      staff().appendRow([uid, nm, '啟用', new Date()]);
      log(nm, '（驗證碼）', '加入');
      return reply(ev.replyToken, [txt('驗證成功，歡迎 ' + nm + '！\n直接輸入建案名稱就能查詢，例如：都更好')]);
    }
    return reply(ev.replyToken, [txt('您尚未開通查詢權限，請輸入主管給您的驗證碼。')]);
  }
  if (!s.on) return reply(ev.replyToken, [txt('您的查詢權限已停用，如有疑問請洽主管。')]);

  var idx = loadIndex();
  if (pick && pick.indexOf('k=') === 0) {
    var it = idx.list.filter(function (x) { return x[3] === pick.slice(2); })[0];
    if (!it) return reply(ev.replyToken, [txt('找不到這個建案，請重新輸入案名。')]);
    log(s.name, '（點選）' + it[0], it[1] + ' ' + it[0]);
    return reply(ev.replyToken, [card(it, idx.asof)]);
  }
  if (/^(說明|help|\?|？)$/i.test(text)) return reply(ev.replyToken, [txt('輸入建案名稱（可以只打部分字）就會回傳實登資訊卡。\n找到多個相近案名時，請點下方按鈕選擇。')]);

  var hits = search(idx.list, text);
  if (!hits.length) { log(s.name, text, '查無'); return reply(ev.replyToken, [txt('查無「' + text + '」，請換個關鍵字試試（例如只打案名的兩三個字）。')]); }
  if (hits.length === 1 || hits[0].exact) { log(s.name, text, hits[0].it[1] + ' ' + hits[0].it[0]); return reply(ev.replyToken, [card(hits[0].it, idx.asof)]); }
  log(s.name, text, '列出 ' + hits.length + ' 案');
  var top = hits.slice(0, 13);
  var m = txt('找到 ' + hits.length + ' 個相近建案' + (hits.length > 13 ? '（列前 13 個，可打完整一點）' : '') + '，請點選：');
  m.quickReply = { items: top.map(function (h) {
    var lb = h.it[0] + '·' + h.it[1].replace('區', '');
    return { type: 'action', action: { type: 'postback', label: lb.length > 20 ? lb.slice(0, 19) + '…' : lb, data: 'k=' + h.it[3], displayText: h.it[0] + '（' + h.it[1] + '）' } };
  }) };
  return reply(ev.replyToken, [m]);
}

// ---------- 搜尋 ----------
function norm(s) {
  return String(s).replace(/[！-～]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/臺/g, '台').replace(/[\s·・\-－_()（）]/g, '').toLowerCase();
}
function search(list, q) {
  var nq = norm(q); if (!nq) return [];
  var out = [];
  list.forEach(function (it) {
    var n = norm(it[0]), sc = 0, exact = false;
    if (n === nq) { sc = 1000; exact = true; }
    else if (n.indexOf(nq) >= 0) sc = 500 - n.length;
    else if (nq.indexOf(n) >= 0 && n.length >= 2) sc = 400 - (nq.length - n.length);
    else { // 字元重疊（容錯打錯字）
      var hit = 0; for (var i = 0; i < nq.length; i++) if (n.indexOf(nq[i]) >= 0) hit++;
      var r = hit / Math.max(nq.length, 2);
      if (nq.length >= 2 && r >= 0.66) sc = 100 * r;
    }
    if (sc) out.push({ it: it, sc: sc + Math.min(it[4], 300) / 1000, exact: exact });
  });
  out.sort(function (a, b) { return b.sc - a.sc; });
  var ex = out.filter(function (h) { return h.exact; });
  if (ex.length > 1) return ex.map(function (h) { h.exact = false; return h; }); // 同名不同區 → 讓員工選
  return out;
}

// ---------- 資料 ----------
function loadIndex() {
  var c = CacheService.getScriptCache(), n = c.get('idx_n');
  if (n) { var parts = []; for (var i = 0; i < +n; i++) parts.push(c.get('idx_' + i)); if (parts.every(Boolean)) return JSON.parse(parts.join('')); }
  var t = UrlFetchApp.fetch(DATA + 'index.json').getContentText(), sz = 90000, o = {};
  var k = Math.ceil(t.length / sz); for (var j = 0; j < k; j++) o['idx_' + j] = t.slice(j * sz, (j + 1) * sz);
  o.idx_n = String(k); c.putAll(o, 21600);
  return JSON.parse(t);
}
function loadP(key) { return JSON.parse(UrlFetchApp.fetch(DATA + 'p/' + key + '.json').getContentText()); }

// ---------- LINE ----------
function reply(token, msgs) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + prop('LINE_TOKEN') },
    payload: JSON.stringify({ replyToken: token, messages: msgs })
  });
}
function txt(t) { return { type: 'text', text: t }; }

// ---------- Flex 卡片 ----------
function avg(a) { return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : null; }
function f1(x) { return x == null ? '—' : x.toFixed(1); }
function ym(v) { return Math.floor(v / 100) + '/' + ('0' + v % 100).slice(-2); }
function comma(x) { return Math.round(x).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function T(text, k) { var d = { type: 'text', text: String(text), color: TX, size: 'sm' }; for (var x in (k || {})) d[x] = k[x]; return d; }
function B(contents, layout, k) { var d = { type: 'box', layout: layout || 'vertical', contents: contents }; for (var x in (k || {})) d[x] = k[x]; return d; }
function kpi(label, val, unit) {
  return B([T(label, { size: 'xxs', color: SUB }), B([T(val, { size: 'lg', weight: 'bold', flex: 0 }), T(unit, { size: 'xxs', color: SUB, flex: 0, margin: 'xs', gravity: 'bottom' })], 'baseline')],
    'vertical', { backgroundColor: PANEL, cornerRadius: '8px', paddingAll: '8px', flex: 1 });
}

function card(it, asof) {
  var p = loadP(it[3]), R = p.rw, url = MAP + '?n=' + encodeURIComponent(p.n) + '&d=' + encodeURIComponent(p.d);
  var head = [B([T(p.n, { size: 'xl', weight: 'bold', color: CY, flex: 0, wrap: true }),
    B([T(p.kind + (p.kind === '預售' ? '建案' : ''), { size: 'xxs', color: CY, align: 'center' })], 'vertical', { borderColor: CY, borderWidth: '1px', cornerRadius: '4px', paddingAll: '2px', paddingStart: '6px', paddingEnd: '6px', flex: 0, margin: 'md' })],
    'horizontal', { alignItems: 'center' })];
  var footer = B([{ type: 'button', style: 'primary', color: '#0E7490', height: 'sm', action: { type: 'uri', label: '看完整明細（線上地圖）', uri: url } },
    T('資料更新至 ' + ym(asof) + '｜內部參考資料，請勿外傳', { size: 'xxs', color: SUB, align: 'center', margin: 'sm', wrap: true })], 'vertical', { spacing: 'sm', backgroundColor: BG, paddingAll: '12px' });
  function bubble(body) { return { type: 'bubble', size: 'giga', body: B(head.concat(body), 'vertical', { spacing: 'md', backgroundColor: BG, paddingAll: '16px' }), footer: footer }; }

  var ns = R.filter(function (r) { return r[2] != null; }), units = ns.map(function (r) { return r[2]; });
  var cnt = [0, 1, 2].map(function (k) { return R.filter(function (r) { return r[1] === k; }).length; });
  var yrs = R.map(function (r) { return Math.floor(r[0] / 100); }).sort();
  // 頁1 摘要
  var line2 = (p.hh ? '住家 ' + p.hh + ' 戶' : '總戶數 —') + (p.sell ? '｜' + p.kind + '去化 ' + Math.round(p.sell * 100) + '%' : '｜去化率 —');
  var p1 = [T(p.d + '｜' + (p.ty || '—') + (p.fl ? '｜' + p.fl + (/^\d+$/.test(p.fl) ? ' 層' : '') : '') + (p.done ? '｜完工 ' + p.done : ''), { color: SUB, margin: 'sm', wrap: true }), T(line2, { color: SUB })];
  if (p.bld) p1.push(T('建設：' + p.bld, { color: SUB, size: 'xs' }));
  if (p.addr) p1.push(T(p.addr, { color: SUB, wrap: true, size: 'xs' }));
  p1 = p1.concat([
    B([kpi('成交筆數', comma(R.length), '筆'), kpi('平均單價', f1(avg(units)), '萬/坪')], 'horizontal', { spacing: 'sm', margin: 'lg' }),
    B([kpi('平均總價', ns.length ? comma(avg(ns.map(function (r) { return r[3]; }))) : '—', '萬'), kpi('平均坪數', f1(avg(ns.map(function (r) { return r[4]; }))), '坪')], 'horizontal', { spacing: 'sm' }),
    B([kpi('單價區間', units.length ? Math.min.apply(null, units).toFixed(1) + '–' + Math.max.apply(null, units).toFixed(1) : '—', '萬/坪')], 'horizontal'),
    B(cnt.map(function (c, k) { return c ? B([], 'vertical', { flex: c, height: '8px', backgroundColor: COL[k] }) : null; }).filter(Boolean), 'horizontal', { cornerRadius: '4px', margin: 'md' }),
    B(cnt.map(function (c, k) { return T('● ' + CN[k] + ' ' + c, { size: 'xs', color: COL[k], flex: 0 }); }), 'horizontal', { spacing: 'lg' }),
    T('統計期間 ' + yrs[0] + '–' + yrs[yrs.length - 1] + ' 全部成交；單價已扣車位，店面不列入均價', { size: 'xxs', color: SUB, wrap: true, margin: 'md' })]);
  // 頁2 明細：近半年最多 15 筆，不足補到 5 筆
  var y = Math.floor(asof / 100), m = asof % 100 - 5; while (m <= 0) { m += 12; y--; } var cut = y * 100 + m;
  var rs = R.slice().sort(function (a, b) { return b[0] - a[0]; });
  var nRecent = rs.filter(function (r) { return r[0] >= cut; }).length, recent = rs.slice(0, Math.min(nRecent, 15));
  var note = '近半年（' + ym(cut) + ' 起）' + nRecent + ' 筆';
  if (recent.length < 5) { recent = rs.slice(0, 5); note += '，不足 5 筆，改列最近 ' + recent.length + ' 筆'; } else if (nRecent > 15) note += '，列最近 15 筆';
  var W = [5, 3, 3, 3, 4, 3, 2];
  function row(cells, hdr) { return B(cells.map(function (c, i) { var col = hdr ? SUB : TX; if (c instanceof Array) { col = c[1]; c = c[0]; } return T(c, { size: 'xs', flex: W[i], align: i ? 'end' : 'start', color: col }); }), 'horizontal', { spacing: 'xs' }); }
  var rows = [row(['年月', '類別', '樓層', '坪數', '總價萬', '單價', '房'], true), { type: 'separator', color: '#2A4560' }];
  recent.forEach(function (r) { rows.push(row([ym(r[0]), [CN[r[1]], COL[r[1]]], r[5] ? r[5] + 'F' : '—', r[4].toFixed(1), comma(r[3]), r[2] != null ? r[2].toFixed(1) : ['店面', OR], r[6] >= 0 ? String(r[6]) : '—'])); });
  var p2 = [T('成交明細', { weight: 'bold', size: 'md' }), T(note, { size: 'xs', color: SUB, wrap: true }), B(rows, 'vertical', { spacing: 'sm', margin: 'md', paddingEnd: '4px' })];
  // 頁3 圖表
  function hbar(label, frac, t) {
    return B([T(label, { size: 'xs', color: SUB, flex: 3 }), B([B([], 'vertical', { backgroundColor: CY, height: '10px', width: Math.round(Math.max(frac, 0.02) * 100) + '%', cornerRadius: '2px' })], 'vertical', { flex: 6, justifyContent: 'center' }),
      T(t, { size: 'xs', flex: 4, align: 'end' })], 'horizontal', { spacing: 'sm', alignItems: 'center' });
  }
  var q = {}; R.forEach(function (r) { var k = Math.floor(r[0] / 100) * 10 + Math.floor((r[0] % 100 - 1) / 3) + 1; (q[k] = q[k] || []).push(r); });
  var qs = Object.keys(q).map(Number).sort().slice(-8), mx = Math.max.apply(null, qs.map(function (k) { return q[k].length; }));
  var sec1 = qs.map(function (k) { var a = q[k]; return hbar(Math.floor(k / 10) + 'Q' + k % 10, a.length / mx, a.length + '筆 ' + f1(avg(a.filter(function (r) { return r[2] != null; }).map(function (r) { return r[2]; })))); });
  var fl = {}; ns.forEach(function (r) { if (r[5]) (fl[r[5]] = fl[r[5]] || []).push(r[2]); });
  var fks = Object.keys(fl).map(Number).sort(function (a, b) { return a - b; }), fa = {}; fks.forEach(function (k) { fa[k] = avg(fl[k]); });
  var vals = fks.map(function (k) { return fa[k]; }), lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
  var sec2 = fks.slice(-20).map(function (k) { return hbar(k + 'F', 0.25 + 0.75 * (hi > lo ? (fa[k] - lo) / (hi - lo) : 1), fa[k].toFixed(1)); });
  var rm = {}; ns.forEach(function (r) { if (r[6] >= 0) { var k = Math.min(r[6], 4); (rm[k] = rm[k] || []).push(r); } });
  function rn(k) { return k == 0 ? '開放/0房' : k == 4 ? '4房以上' : k + '房'; }
  var hd = [['房型', 3], ['筆數', 2], ['均坪', 2], ['均單價', 2], ['均總價', 3]];
  var rw3 = [B(hd.map(function (h, i) { return T(h[0], { size: 'xs', color: SUB, flex: h[1], align: i ? 'end' : 'start' }); }), 'horizontal')];
  Object.keys(rm).sort().forEach(function (k) { var a = rm[k];
    var c = [rn(k), String(a.length), f1(avg(a.map(function (r) { return r[4]; }))), f1(avg(a.map(function (r) { return r[2]; }))), comma(avg(a.map(function (r) { return r[3]; })))];
    rw3.push(B(c.map(function (x, i) { return T(x, { size: 'xs', flex: hd[i][1], align: i ? 'end' : 'start' }); }), 'horizontal')); });
  var p3 = [T('季成交量與平均單價', { weight: 'bold', size: 'md' }), T('最近 8 季，右側＝筆數與均價（萬/坪）', { size: 'xxs', color: SUB }), B(sec1, 'vertical', { spacing: 'xs' })];
  if (fks.length) p3 = p3.concat([T('樓層單價（萬/坪）', { weight: 'bold', size: 'md', margin: 'lg' }), B(sec2, 'vertical', { spacing: 'xs' })]);
  if (Object.keys(rm).length) p3 = p3.concat([T('房型統計', { weight: 'bold', size: 'md', margin: 'lg' }), B(rw3, 'vertical', { spacing: 'xs' })]);
  return { type: 'flex', altText: p.n + ' 實登資訊', contents: { type: 'carousel', contents: [bubble(p1), bubble(p2), bubble(p3)] } };
}
