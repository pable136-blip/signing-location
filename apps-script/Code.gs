/**
 * 행사 근무자 QR 출석 — 서버(Google Apps Script + 구글 시트)
 *
 * 종이 서명부를 대신한다. 부스 태블릿이 30초마다 바뀌는 QR 을 띄우고,
 * 근무자가 휴대폰 카메라로 찍으면 checkin.html 이 이 스크립트로 출석을 보낸다.
 *
 * 「QR 사진을 찍어 다른 곳에서 출석」을 막는 장치(모두 서버에서 판정한다):
 *   1. 시간 토큰  — QR 안의 토큰은 HMAC(부스키, 행사ID.시간칸)이라 30초마다 바뀌고
 *                   약 1~1.5분 뒤 만료된다. 사진·캡처·링크를 돌려도 금방 쓸모가 없다.
 *                   부스키는 부스 태블릿과 이 시트에만 있으므로 토큰을 위조할 수 없다.
 *   2. GPS 반경   — 출석 순간의 휴대폰 위치가 행사장 반경 밖이거나 정확도가 낮으면 거부.
 *   3. 기기 고정  — 근무자 1명 = 휴대폰 1대. 다른 폰으로 대신 찍거나, 한 폰으로
 *                   여러 사람을 찍어 주는 대리출석을 거부.
 *   4. 행사 기간  — 시작~종료 일시 밖의 출석은 거부.
 * 거부된 시도는 모두 「거부기록」 시트에 사유·거리와 함께 남는다.
 *
 * 개인정보 보호(외부 유출 방지):
 *   - 기록은 이 스프레드시트에만 남는다. GitHub 저장소·웹 화면에는 개인정보가 없다.
 *   - 휴대폰 좌표(위도·경도)는 저장하지 않고 「행사장까지 거리」만 남긴다.
 *   - 연락처는 뒤 4자리만, 브라우저 정보는 받지 않는다.
 *   - 웹 API 는 어떤 경우에도 명단·기록을 돌려주지 않는다(본인 출석 결과만).
 *   - 보존기간(CONFIG.RETENTION_DAYS)이 지난 기록은 메뉴로 일괄 삭제한다.
 *
 * 설치 방법은 attendance/README.md 참고.
 */

// ── 설정 ────────────────────────────────────────────────────────────
var CONFIG = {
  SLOT_SECONDS: 30,      // QR 이 바뀌는 주기(초). 부스 화면은 서버에서 이 값을 받아 쓴다.
  SLOTS_PAST: 2,         // 몇 칸 전 토큰까지 받아 줄지(스캔·입력 지연 허용) → 최대 약 90초
  SLOTS_FUTURE: 1,       // 부스 태블릿 시계가 약간 빠른 경우 허용
  MAX_ACCURACY_M: 100,   // GPS 오차가 이보다 크면 위치를 믿을 수 없어 거부
  MAX_FIX_AGE_MS: 120000,// 2분보다 오래된 위치값(캐시)은 거부
  DEFAULT_RADIUS_M: 200, // 행사 시트에 반경이 비어 있을 때
  RETENTION_DAYS: 365,   // 출석·거부 기록 보존기간(일) — 메뉴 「보존기간 지난 기록 삭제」 기준
  TZ: 'Asia/Seoul'
};

var SHEET = {
  EVENTS: '행사',
  WORKERS: '근무자',
  LOG: '출석기록',
  REJECT: '거부기록'
};

var HEADERS = {};
HEADERS[SHEET.EVENTS] = ['행사ID', '행사명', '위도', '경도', '허용반경(m)', '시작일시', '종료일시',
                         '사전명단만허용(Y/N)', '부스키', '사용(Y/N)'];
HEADERS[SHEET.WORKERS] = ['행사ID', '이름', '연락처뒤4자리', '기기ID', '기기등록일시'];
HEADERS[SHEET.LOG] = ['일시', '행사ID', '이름', '연락처뒤4자리', '구분', '거리(m)', 'GPS오차(m)'];
HEADERS[SHEET.REJECT] = ['일시', '행사ID', '이름', '연락처뒤4자리', '구분', '사유', '거리(m)',
                         'GPS오차(m)', '기기ID'];

// ── 시트 메뉴 / 초기 설정 ───────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('출석관리')
    .addItem('시트 초기 설정', 'setup')
    .addItem('빈 행사ID·부스키 채우기', 'fillEventKeys')
    .addItem('보존기간 지난 기록 삭제', 'purgeOldRecords')
    .addToUi();
}

/** 필요한 시트와 머리글을 만든다. 여러 번 실행해도 안전하다. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TZ);     // 시트에 적힌 일시를 한국 시간으로 해석
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    }
  });
  // 연락처 뒷자리 「0123」 의 앞자리 0 이 사라지지 않도록 텍스트 서식
  ss.getSheetByName(SHEET.WORKERS).getRange('C:C').setNumberFormat('@');
  ss.getSheetByName(SHEET.LOG).getRange('D:D').setNumberFormat('@');
  ss.getSheetByName(SHEET.REJECT).getRange('D:D').setNumberFormat('@');
  fillEventKeys();
}

/** 행사 시트에서 행사ID·부스키가 빈 행을 무작위 값으로 채운다. */
function fillEventKeys() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.EVENTS);
  if (!sh || sh.getLastRow() < 2) return;
  var range = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS[SHEET.EVENTS].length);
  var rows = range.getValues();
  rows.forEach(function (r) {
    if (!r[1]) return;                       // 행사명이 없는 빈 줄은 건너뜀
    if (!r[0]) r[0] = 'EV' + randomString_(6).toUpperCase();
    if (!r[8]) r[8] = Utilities.getUuid().replace(/-/g, '');   // 122비트 무작위
    if (!r[7]) r[7] = 'N';                   // 기본: 명단 없이 첫 출석 때 자동 등록
    if (!r[9]) r[9] = 'Y';
  });
  range.setValues(rows);
}

/** 출석기록·거부기록에서 보존기간(CONFIG.RETENTION_DAYS)이 지난 행을 지운다. */
function purgeOldRecords() {
  var cutoff = Utilities.formatDate(new Date(Date.now() - CONFIG.RETENTION_DAYS * 86400000), CONFIG.TZ, 'yyyy-MM-dd');
  var removed = 0;
  [SHEET.LOG, SHEET.REJECT].forEach(function (name) {
    var sh = sheet_(name);
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {      // 아래에서부터 지워야 행 번호가 밀리지 않는다
      if (dayOf_(rows[i][0]) < cutoff) { sh.deleteRow(i + 1); removed++; }
    }
  });
  SpreadsheetApp.getUi().alert(cutoff + ' 이전 기록 ' + removed + '건을 삭제했습니다.');
}

// ── 웹 API ─────────────────────────────────────────────────────────
/**
 * GET ?action=time            → 서버 시각·QR 주기 (부스 화면 시계 맞춤)
 * GET ?action=event&e=행사ID&t=토큰 → 행사명 (출석 화면 표시용). 유효한 QR 토큰이 있어야
 *                                    답하므로 행사ID만으로 행사명을 알아낼 수 없다.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'time') {
    return json_({ ok: true, now: Date.now(), slotSeconds: CONFIG.SLOT_SECONDS });
  }
  if (p.action === 'event') {
    var ev = findEvent_(p.e);
    if (!ev || !ev.active || !tokenValid_(ev, p.t, Date.now())) {
      return json_({ ok: false, code: 'TOKEN_EXPIRED', message: 'QR이 만료되었습니다. 부스 화면의 QR을 다시 스캔해 주세요.' });
    }
    return json_({ ok: true, name: ev.name });
  }
  return json_({ ok: false, code: 'BAD_REQUEST', message: '알 수 없는 요청입니다.' });
}

/**
 * POST (본문은 JSON 문자열, Content-Type: text/plain — CORS 사전요청을 피하기 위함)
 *   {action:'boothcheck', eventId, token}  부스 화면이 부스키를 맞게 넣었는지 확인
 *   {action:'checkin', eventId, token, deviceId, name, phone4, type, lat, lng, accuracy, fixTime}
 *   위도·경도는 거리 계산에만 쓰고 저장하지 않는다.
 */
function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, code: 'BAD_REQUEST', message: '요청 형식이 올바르지 않습니다.' });
  }
  if (req.action === 'boothcheck') return json_(boothCheck_(req));
  if (req.action === 'checkin') {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);                    // 동시 출석 시 기기 등록이 꼬이지 않도록 직렬화
    try {
      return json_(checkin_(req));
    } finally {
      lock.releaseLock();
    }
  }
  return json_({ ok: false, code: 'BAD_REQUEST', message: '알 수 없는 요청입니다.' });
}

function boothCheck_(req) {
  var ev = findEvent_(req.eventId);
  if (!ev || !ev.active) return { ok: false, code: 'NO_EVENT', message: '행사ID를 찾을 수 없거나 사용(Y)이 아닙니다.' };
  if (!tokenValid_(ev, req.token, Date.now())) return { ok: false, code: 'BAD_KEY', message: '부스키가 맞지 않습니다.' };
  return { ok: true, name: ev.name, slotSeconds: CONFIG.SLOT_SECONDS };
}

function checkin_(req) {
  var now = Date.now();
  var name = String(req.name || '').trim();
  var phone4 = String(req.phone4 || '').replace(/\D/g, '');
  var type = req.type === '퇴근' ? '퇴근' : '출근';
  var deviceId = String(req.deviceId || '');
  var lat = Number(req.lat), lng = Number(req.lng), acc = Number(req.accuracy);
  var ctx = { req: req, name: name, phone4: phone4, type: type, deviceId: deviceId,
              lat: lat, lng: lng, acc: acc, dist: '' };

  var ev = findEvent_(req.eventId);
  if (!ev || !ev.active) return reject_(ctx, 'NO_EVENT', '등록되지 않았거나 종료된 행사입니다.');
  ctx.eventId = ev.id;

  // 1) 시간 토큰 — QR 사진·링크 재사용 차단
  if (!tokenValid_(ev, req.token, now)) {
    return reject_(ctx, 'TOKEN_EXPIRED', 'QR이 만료되었습니다. 부스 화면의 QR을 지금 다시 스캔해 주세요.');
  }
  // 2) 행사 기간
  if (ev.start && now < ev.start.getTime()) return reject_(ctx, 'NOT_STARTED', '아직 출석 시간이 아닙니다.');
  if (ev.end && now > ev.end.getTime()) return reject_(ctx, 'ENDED', '출석 시간이 지났습니다.');
  // 3) 입력값
  if (!name || phone4.length !== 4) return reject_(ctx, 'BAD_INPUT', '이름과 연락처 뒤 4자리를 입력해 주세요.');
  if (!/^[A-Za-z0-9-]{16,64}$/.test(deviceId)) return reject_(ctx, 'BAD_DEVICE', '기기 정보를 확인할 수 없습니다. 일반 브라우저(크롬·사파리)로 열어 주세요.');
  // 4) GPS
  if (!isFinite(lat) || !isFinite(lng) || !isFinite(acc)) {
    return reject_(ctx, 'NO_GPS', '위치 정보가 필요합니다. 위치 권한을 허용해 주세요.');
  }
  if (req.fixTime && Math.abs(now - Number(req.fixTime)) > CONFIG.MAX_FIX_AGE_MS) {
    return reject_(ctx, 'STALE_GPS', '위치 정보가 오래되었습니다. 다시 시도해 주세요.');
  }
  if (acc > CONFIG.MAX_ACCURACY_M) {
    return reject_(ctx, 'LOW_ACCURACY', '위치 정확도가 낮습니다(오차 ' + Math.round(acc) + 'm). 와이파이·GPS를 켜고 다시 시도해 주세요.');
  }
  if (isFinite(ev.lat) && isFinite(ev.lng)) {
    ctx.dist = Math.round(distanceM_(lat, lng, ev.lat, ev.lng));
    if (ctx.dist > ev.radius) {
      return reject_(ctx, 'OUT_OF_RANGE', '행사장 밖에서는 출석할 수 없습니다(행사장까지 약 ' + ctx.dist + 'm).');
    }
  }
  // 5) 근무자·기기 고정 — 대리출석 차단
  var ws = sheet_(SHEET.WORKERS);
  var rows = ws.getDataRange().getValues();
  var me = -1, deviceOwner = -1, hasRoster = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== ev.id) continue;
    hasRoster = true;
    if (normName_(rows[i][1]) === normName_(name) && pad4_(rows[i][2]) === phone4) me = i;
    if (String(rows[i][3]) === deviceId) deviceOwner = i;
  }
  if (deviceOwner !== -1 && deviceOwner !== me) {
    return reject_(ctx, 'DEVICE_OTHER_WORKER', '이 휴대폰은 이미 다른 근무자로 등록되어 있습니다. 본인 휴대폰으로 출석해 주세요.');
  }
  if (me === -1) {
    if (ev.rosterOnly) return reject_(ctx, 'NOT_IN_ROSTER', '근무자 명단에 없습니다. 이름·연락처 뒤 4자리를 확인하거나 담당자에게 문의해 주세요.');
    ws.appendRow([ev.id, name, "'" + phone4, deviceId, stamp_(now)]);
  } else if (!rows[me][3]) {
    ws.getRange(me + 1, 4, 1, 2).setValues([[deviceId, stamp_(now)]]);   // 첫 출석 기기로 고정
  } else if (String(rows[me][3]) !== deviceId) {
    return reject_(ctx, 'WORKER_OTHER_DEVICE', '처음 출석한 휴대폰이 아닙니다. 본인 휴대폰으로 출석해 주세요(휴대폰을 바꿨다면 담당자에게 기기 초기화를 요청).');
  }
  // 6) 같은 날 같은 구분 중복
  var today = Utilities.formatDate(new Date(now), CONFIG.TZ, 'yyyy-MM-dd');
  var log = sheet_(SHEET.LOG);
  var logRows = log.getDataRange().getValues();
  for (var j = logRows.length - 1; j >= 1; j--) {
    var r = logRows[j];
    if (String(r[1]) === ev.id && normName_(r[2]) === normName_(name) && pad4_(r[3]) === phone4 &&
        r[4] === type && dayOf_(r[0]) === today) {
      return { ok: true, duplicate: true, message: '이미 ' + type + ' 처리되었습니다(' + timeOf_(r[0]) + ').',
               event: ev.name, name: name, type: type, time: timeOf_(r[0]) };
    }
  }
  log.appendRow([stamp_(now), ev.id, name, "'" + phone4, type, ctx.dist, Math.round(acc)]);
  return { ok: true, message: type + ' 완료', event: ev.name, name: name, type: type,
           time: Utilities.formatDate(new Date(now), CONFIG.TZ, 'HH:mm:ss'), distance: ctx.dist };
}

// ── 토큰 ────────────────────────────────────────────────────────────
/** token = base64url(HMAC-SHA256(부스키, 행사ID + '.' + 시간칸)) 앞 16자. booth.html 과 같은 식. */
function makeToken_(secret, eventId, slot) {
  var sig = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(eventId + '.' + slot).getBytes(),
    Utilities.newBlob(secret).getBytes());
  return Utilities.base64EncodeWebSafe(sig).substring(0, 16);
}

function tokenValid_(ev, token, now) {
  if (!token || typeof token !== 'string' || token.length !== 16) return false;
  var slot = Math.floor(now / 1000 / CONFIG.SLOT_SECONDS);
  for (var s = slot - CONFIG.SLOTS_PAST; s <= slot + CONFIG.SLOTS_FUTURE; s++) {
    if (makeToken_(ev.secret, ev.id, s) === token) return true;
  }
  return false;
}

// ── 도우미 ──────────────────────────────────────────────────────────
function findEvent_(id) {
  id = String(id || '').trim();
  if (!id) return null;
  var sh = sheet_(SHEET.EVENTS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0]).trim() !== id) continue;
    return {
      id: id,
      name: String(r[1]),
      lat: r[2] === '' ? NaN : Number(r[2]),
      lng: r[3] === '' ? NaN : Number(r[3]),
      radius: Number(r[4]) || CONFIG.DEFAULT_RADIUS_M,
      start: toDate_(r[5]),
      end: toDate_(r[6]),
      rosterOnly: String(r[7]).trim().toUpperCase() === 'Y',
      secret: String(r[8]),
      active: String(r[9]).trim().toUpperCase() !== 'N' && String(r[8]).length >= 8
    };
  }
  return null;
}

function reject_(ctx, code, message) {
  try {
    // 없는 행사ID 로 들어온 요청은 기록하지 않는다(외부에서 시트를 채우는 장난 방지).
    if (ctx.eventId) {
      sheet_(SHEET.REJECT).appendRow([
        stamp_(Date.now()), ctx.eventId, String(ctx.name || '').substring(0, 30), "'" + (ctx.phone4 || '').substring(0, 4),
        ctx.type || '', code + ' ' + message, ctx.dist,
        isFinite(ctx.acc) ? Math.round(ctx.acc) : '', String(ctx.deviceId || '').substring(0, 64)
      ]);
    }
  } catch (err) { /* 기록 실패가 응답을 막지 않도록 */ }
  return { ok: false, code: code, message: message };
}

function distanceM_(lat1, lng1, lat2, lng2) {
  var R = 6371000, rad = Math.PI / 180;
  var dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.asin(Math.sqrt(a));
}

function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」 시트가 없습니다. 메뉴 출석관리 > 시트 초기 설정 을 실행하세요.');
  return sh;
}

function toDate_(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  var d = new Date(String(v).replace(' ', 'T') + (/[+Z]/.test(String(v)) ? '' : '+09:00'));
  return isNaN(d.getTime()) ? null : d;
}

function stamp_(ms) { return Utilities.formatDate(new Date(ms), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'); }
function dayOf_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd') : String(v).substring(0, 10); }
function timeOf_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, 'HH:mm:ss') : String(v).substring(11, 19); }
function normName_(v) { return String(v).replace(/\s+/g, ''); }
function pad4_(v) { return ('0000' + String(v).replace(/\D/g, '')).slice(-4); }

function randomString_(n) {
  var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < n; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
