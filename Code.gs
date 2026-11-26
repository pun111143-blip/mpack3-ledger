/**
 * ============================================================
 *  📊 엠팩3노래연습장 가계부 — 웹 대시보드용 API (Code.gs)
 * ============================================================
 *  쓰는 법
 *   1) 가계부 스프레드시트에서 [확장 프로그램] → [Apps Script] 를 엽니다.
 *   2) 기본으로 있던 내용을 모두 지우고 이 파일 전체를 붙여넣습니다.
 *   3) 아래 CONFIG.PIN 의 '0000' 을 원하는 4자리 비밀번호로 바꿉니다.
 *      (대시보드 잠금화면과 같은 번호여야 합니다)
 *   4) [배포] → [새 배포] → 유형 '웹 앱'
 *        · 설명: 엠팩3 가계부 API
 *        · 다음 사용자로 실행: 나
 *        · 액세스 권한: 모든 사용자
 *      → [배포] → 나온 '웹 앱 URL'(…/exec)을 복사
 *   5) 대시보드 첫 화면의 [시트 연결 설정]에 그 URL과 비밀번호를 넣으면 끝.
 *
 *  · 시트를 고치면 대시보드는 5분 안에 새 숫자로 바뀝니다(캐시).
 *    바로 보고 싶으면 대시보드의 새로고침(↻) 버튼을 누르세요.
 *  · 이 스크립트는 시트를 '읽기만' 합니다. 지우거나 고치지 않습니다.
 * ============================================================
 */

var CONFIG = {
  SPREADSHEET_ID: '154v50zIsQA5tkSELePVHpSVDEfsyPWxEmUAQIzXYeUM',

  // ▼▼▼ 여기를 4자리 비밀번호로 바꾸세요 (예: '1234') ▼▼▼
  PIN: '0000',
  // ▲▲▲ 대시보드 잠금화면과 같은 번호 ▲▲▲

  CACHE_SECONDS: 300,          // 같은 결과를 5분간 재사용 (시트 호출 절약)
  TZ: 'Asia/Seoul'
};

/* ------------------------------------------------------------
 *  웹 요청 입구
 *    GET ?pin=0000                 → 그 달 전체 데이터(JSON)
 *    GET ?pin=0000&month=12월       → 지정한 달
 *    GET ?pin=0000&ping=1          → 비밀번호 확인 + 달 목록만
 *    GET ...&callback=함수명        → JSONP (브라우저가 fetch를 막을 때 대비)
 * ---------------------------------------------------------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = /^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(String(p.callback || '')) ? String(p.callback) : '';
  var out;
  try {
    if (String(p.pin || '').trim() !== String(CONFIG.PIN)) {
      out = { ok: false, code: 'pin', message: '비밀번호가 맞지 않습니다' };
    } else if (String(p.ping || '') === '1' && !p.month) {
      out = ping_();
    } else {
      out = buildPayload_(p.month || '');
    }
  } catch (err) {
    out = { ok: false, code: 'error', message: String((err && err.message) || err) };
  }
  return respond_(out, cb);
}

function respond_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------
 *  달 목록 · 비밀번호 확인
 * ---------------------------------------------------------- */
function monthTabs_(ss) {
  var names = ss.getSheets().map(function (s) { return s.getName().trim(); });
  return names
    .filter(function (n) { return /^\d{1,2}월$/.test(n); })
    .sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); });
}

function ping_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var months = monthTabs_(ss);
  return { ok: true, ping: true, months: months, month: months[months.length - 1] || '' };
}

/* ------------------------------------------------------------
 *  본문 만들기
 * ---------------------------------------------------------- */
function buildPayload_(monthParam) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var months = monthTabs_(ss);
  var month = String(monthParam || '').trim();
  if (months.indexOf(month) < 0) month = months[months.length - 1] || '';
  if (!month) return { ok: false, code: 'nomon', message: '월 시트를 찾지 못했습니다' };

  var cache = CacheService.getScriptCache();
  var key = 'ledger1:' + month;
  var hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }

  var sh = ss.getSheetByName(month);
  if (!sh) return { ok: false, code: 'nomon', message: '시트를 찾지 못했습니다: ' + month };

  var values = sh.getDataRange().getValues().map(function (r) { return r.slice(0, 6); });
  var payload = parseMonth_(values, month, ss.getName(), months);
  payload.trend = parseTrend_(ss.getSheetByName('월별 추이'));

  var now = new Date();
  payload.updatedAt = Utilities.formatDate(now, CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");

  var json = JSON.stringify(payload);
  if (json.length < 95000) cache.put(key, json, CONFIG.CACHE_SECONDS);
  return payload;
}

/* ------------------------------------------------------------
 *  월 시트 하나 → 화면용 데이터
 *    행 위치가 달마다 달라도(30일/31일) 소계 줄을 기준으로 찾아갑니다.
 * ---------------------------------------------------------- */
function parseMonth_(values, month, title, months) {
  var T = txt_, N = num_, NM = norm_;

  var revenue = N(labelVal_(values, '월 총매출', 0));
  var expense = N(labelVal_(values, '월 총지출', 0));
  var profit  = N(labelVal_(values, '월 순익', 0));
  var ymd = String(title || '').match(/(20\d{2})/);

  /* --- 일별 매출 --- */
  var daily = [];
  var hi = -1;
  for (var i = 0; i < values.length; i++) {
    if (NM(values[i][0]) === '날짜' && NM(values[i][1]) === '현금') { hi = i; break; }
  }
  var sc = 0, sk = 0, st = 0;
  if (hi >= 0) {
    for (var j = hi + 1; j < values.length; j++) {
      var a = T(values[j][0]);
      if (NM(a) === '합계') break;
      var m = a.match(/^(\d{1,2})\s*일$/);
      if (!m) { if (a === '') continue; break; }
      var cash = N(values[j][1]), card = N(values[j][2]), tr = N(values[j][3]);
      var total = N(values[j][4]) || (cash + card + tr);
      daily.push({ day: Number(m[1]), cash: cash, card: card, transfer: tr, total: total });
      sc += cash; sk += card; st += tr;
    }
  }

  /* --- 지출 3분류 (소계 줄을 기준으로 위로 거슬러 항목을 모음) --- */
  var fixed  = section_(values, '고정비 소계');
  var online = section_(values, '온라인 소계');
  var onsite = section_(values, '현장 소계');
  var liquor = section_(values, '주류·음료 소계');
  var repair = section_(values, '수리·관리 소계');
  var labor  = section_(values, '인건비 소계');

  var fixedSub  = fixed.subtotal  || N(labelVal_(values, '① 고정비', 3));
  var onlineSub = online.subtotal;
  var onsiteSub = onsite.subtotal;
  var liquorSub = liquor.subtotal;
  var goodsSub  = N(labelVal_(values, '② 물품비용', 3)) || (onlineSub + onsiteSub + liquorSub);
  var etcSub    = N(labelVal_(values, '③ 수리·관리·인건비', 3)) || (repair.subtotal + labor.subtotal);

  /* --- 참고 메모 ('■ 참고' 줄 아래의 '·' 줄만) --- */
  var notes = [];
  var noteStart = -1;
  for (var s0 = 0; s0 < values.length; s0++) {
    if (NM(values[s0][0]).indexOf('■참고') === 0) { noteStart = s0; break; }
  }
  for (var k = Math.max(0, noteStart); k < values.length; k++) {
    if (noteStart < 0) break;
    var cell = T(values[k][0]);
    if (cell.indexOf('·') === 0) notes.push(cell.replace(/^·\s*/, ''));
  }

  var mixBase = revenue || 1;
  return {
    ok: true,
    month: month,
    title: title,
    year: ymd ? Number(ymd[1]) : new Date().getFullYear(),
    months: months,
    summary: { revenue: revenue, expense: expense, profit: profit },
    mix: {
      cash: sc, card: sk, transfer: st,
      cashPct: sc / mixBase, cardPct: sk / mixBase, transferPct: st / mixBase
    },
    daily: daily,
    expense: {
      fixed: { subtotal: fixedSub, items: fixed.items },
      goods: {
        subtotal: goodsSub,
        online: { subtotal: onlineSub, items: online.items },
        onsite: { subtotal: onsiteSub, items: onsite.items },
        liquor: { subtotal: liquorSub, items: liquor.items }
      },
      etc: {
        subtotal: etcSub,
        repair: { subtotal: repair.subtotal, items: repair.items },
        labor: { subtotal: labor.subtotal, items: labor.items }
      }
    },
    notes: notes
  };
}

/* 소계 줄(예: '① 고정비 소계')을 찾아 그 위의 표(머리글 + 항목)를 읽는다 */
function section_(values, subtotalNeedle) {
  var T = txt_, NM = norm_;
  var e = -1;
  for (var i = 0; i < values.length; i++) {
    if (NM(values[i][0]).indexOf(NM(subtotalNeedle)) >= 0) { e = i; break; }
  }
  if (e < 0) return { subtotal: 0, items: [] };

  var h = -1;
  for (var b = e - 1; b >= 0 && b >= e - 40; b--) {
    var a = NM(values[b][0]);
    if (a === '유형' || a === '날짜') { h = b; break; }
    if (a.indexOf('소계') >= 0 || a.indexOf('■') >= 0) break;
  }

  var items = [];
  if (h >= 0) {
    var header = values[h].map(T);
    var isDate = NM(header[0]) === '날짜';
    var isName = false;
    for (var c = 0; c < header.length; c++) if (NM(header[c]) === '이름') isName = true;

    for (var r = h + 1; r < e; r++) {
      var row = values[r];
      var any = false;
      for (var q = 0; q < row.length; q++) if (T(row[q]) !== '') { any = true; break; }
      if (!any) continue;
      if (isDate) {
        items.push({ date: T(row[0]), vendor: T(row[1]), memo: T(row[2]), amount: num_(row[3]) });
      } else if (isName) {
        items.push({ type: T(row[0]), name: T(row[1]), amount: num_(row[2]), memo: T(row[3]) });
      } else {
        items.push({ type: T(row[0]), name: T(row[1]), amount: num_(row[2]), memo: T(row[3]) });
      }
    }
  }
  return { subtotal: lastNumber_(values[e]), items: items };
}

/* ------------------------------------------------------------
 *  '월별 추이' 시트 → 월별 배열
 * ---------------------------------------------------------- */
function parseTrend_(sh) {
  if (!sh) return [];
  var T = txt_, N = num_, NM = norm_;
  var values = sh.getDataRange().getValues().map(function (r) { return r.slice(0, 7); });
  var map = {}, order = [];

  var hi = -1;
  for (var i = 0; i < values.length; i++) {
    if (NM(values[i][0]) === '월' && NM(values[i][1]).indexOf('총매출') >= 0) { hi = i; break; }
  }
  if (hi >= 0) {
    for (var j = hi + 1; j < values.length; j++) {
      var m = T(values[j][0]);
      if (!/^\d{1,2}월$/.test(m)) { if (m === '') continue; break; }
      map[m] = {
        month: m, revenue: N(values[j][1]), expense: N(values[j][2]), profit: N(values[j][3]),
        cash: N(values[j][4]), card: N(values[j][5]), transfer: N(values[j][6]),
        fixed: 0, goods: 0, online: 0, onsite: 0, liquor: 0, etc: 0
      };
      order.push(m);
    }
  }

  var hj = -1;
  for (var p = 0; p < values.length; p++) {
    if (NM(values[p][0]) === '월' && NM(values[p][1]).indexOf('①고정비') >= 0) { hj = p; break; }
  }
  if (hj >= 0) {
    for (var q = hj + 1; q < values.length; q++) {
      var mm = T(values[q][0]);
      if (!/^\d{1,2}월$/.test(mm)) { if (mm === '') continue; break; }
      if (!map[mm]) {
        map[mm] = { month: mm, revenue: 0, expense: 0, profit: 0, cash: 0, card: 0, transfer: 0 };
        order.push(mm);
      }
      map[mm].fixed  = N(values[q][1]);
      map[mm].goods  = N(values[q][2]);
      map[mm].online = N(values[q][3]);
      map[mm].onsite = N(values[q][4]);
      map[mm].liquor = N(values[q][5]);
      map[mm].etc    = N(values[q][6]);
    }
  }
  return order.map(function (m) { return map[m]; });
}

/* ------------------------------------------------------------
 *  작은 도우미들
 * ---------------------------------------------------------- */
function labelVal_(values, label, col) {
  var target = norm_(label);
  for (var i = 0; i < values.length; i++) {
    if (norm_(values[i][col]) === target) return values[i][col + 1];
  }
  return '';
}

function lastNumber_(row) {
  for (var i = row.length - 1; i >= 0; i--) {
    if (typeof row[i] === 'number' && isFinite(row[i])) return row[i];
    var s = txt_(row[i]).replace(/[,\s원]/g, '');
    if (s !== '' && /^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  }
  return 0;
}

function txt_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'MM/dd');
  return String(v).trim();
}

function norm_(v) {
  return txt_(v).replace(/\s+/g, '').replace(/[—–]/g, '-');
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = txt_(v).replace(/[,\s원₩]/g, '');
  if (s === '' || s === '-') return 0;
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

/* ------------------------------------------------------------
 *  디버그용: 편집기에서 이 함수를 실행하면 로그에 결과가 찍힙니다.
 * ---------------------------------------------------------- */
function testRun() {
  var data = buildPayload_('');
  Logger.log(JSON.stringify({
    month: data.month,
    summary: data.summary,
    goods: data.expense && data.expense.goods && data.expense.goods.subtotal,
    daily: data.daily && data.daily.length,
    trend: (data.trend || []).map(function (t) { return t.month + ':' + t.revenue; })
  }, null, 2));
}
