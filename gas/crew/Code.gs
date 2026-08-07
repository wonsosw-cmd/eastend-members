/**
 * EASTEND 크루 카페 사용권 — Apps Script 백엔드
 *
 * 구조: 크루 앱(코드 발급) → 카페 공통 승인 화면(코드+금액 입력) → 시트 기록 → 관리자 집계
 * 시트: 크루 / 사용내역 / 카페 / 설정   (개인정보 포함 — 링크 공유 금지)
 *
 * 최초 1회 setup() 실행 → 로그에 관리자 토큰·카페 토큰 출력
 * 코드 수정 후: 배포 > 배포 관리 > 수정 > 새 버전 (URL 유지)
 */

var SHEET_CREW = "크루";
var SHEET_USAGE = "사용내역";
var SHEET_CAFE = "카페";
var SHEET_CONF = "설정";
var TZ = "Asia/Seoul";
var CODE_TTL = 180;          // 발급 코드 유효시간(초) = 3분
var RESULT_TTL = 600;        // 승인 결과를 크루 화면이 받아갈 수 있는 시간(초)

// 기본 정책 (설정 시트에서 덮어쓸 수 있음)
var DEFAULT_MONTHLY_LIMIT = 3;
var DEFAULT_VOUCHER_AMOUNT = 15000;

// ─────────────────────────────────────────────
// 최초 1회 실행
// ─────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  var crew = ss.getSheetByName(SHEET_CREW) || ss.insertSheet(SHEET_CREW);
  if (crew.getLastRow() === 0) {
    crew.appendRow(["등록일시", "크루ID", "이름", "휴대전화", "소속", "상태", "월한도", "메모"]);
    crew.setFrozenRows(1);
    crew.setColumnWidths(1, 8, 130);
  }

  var usage = ss.getSheetByName(SHEET_USAGE) || ss.insertSheet(SHEET_USAGE);
  if (usage.getLastRow() === 0) {
    usage.appendRow([
      "승인일시", "사용ID", "크루ID", "이름", "소속", "카페",
      "결제금액", "회사부담", "자기부담", "승인자", "코드", "년월", "상태", "메모"
    ]);
    usage.setFrozenRows(1);
  }

  var cafe = ss.getSheetByName(SHEET_CAFE) || ss.insertSheet(SHEET_CAFE);
  if (cafe.getLastRow() === 0) {
    cafe.appendRow(["카페명", "상태", "메모"]);
    cafe.appendRow(["제휴 카페", "활성", "관리자 화면에서 이름을 바꾸거나 추가하세요"]);
    cafe.setFrozenRows(1);
  }

  var conf = ss.getSheetByName(SHEET_CONF) || ss.insertSheet(SHEET_CONF);
  if (conf.getLastRow() === 0) {
    conf.appendRow(["항목", "값", "설명"]);
    conf.appendRow(["월한도", DEFAULT_MONTHLY_LIMIT, "1인당 월 최대 방문 횟수"]);
    conf.appendRow(["1회한도", DEFAULT_VOUCHER_AMOUNT, "1회당 회사 부담 상한(원). 초과분은 크루 자부담"]);
    conf.setFrozenRows(1);
  }

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("ADMIN_TOKEN")) {
    props.setProperty("ADMIN_TOKEN", Utilities.getUuid().replace(/-/g, "").slice(0, 20));
  }
  if (!props.getProperty("CAFE_TOKEN")) {
    props.setProperty("CAFE_TOKEN", Utilities.getUuid().replace(/-/g, "").slice(0, 8));
  }

  Logger.log("관리자 토큰(admin.html): " + props.getProperty("ADMIN_TOKEN"));
  Logger.log("카페 토큰(cafe.html): " + props.getProperty("CAFE_TOKEN"));
}

/** 카페 토큰을 원하는 값으로 바꾸고 싶을 때 이 함수 본문을 고쳐 1회 실행 */
function setCafeToken() {
  PropertiesService.getScriptProperties().setProperty("CAFE_TOKEN", "eastcafe");
  Logger.log("카페 토큰 변경 완료: " + PropertiesService.getScriptProperties().getProperty("CAFE_TOKEN"));
}

// ─────────────────────────────────────────────
// 공용 유틸
// ─────────────────────────────────────────────
function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function props_() {
  return PropertiesService.getScriptProperties();
}
function cache_() {
  return CacheService.getScriptCache();
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function toIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
  return String(v || "");
}
function ym_(d) {
  return Utilities.formatDate(d || new Date(), TZ, "yyyy-MM");
}
function ymd_(d) {
  return Utilities.formatDate(d || new Date(), TZ, "yyyy-MM-dd");
}

/** 휴대전화 정규화: 숫자만 남기고 010-1234-5678 형태로 */
function normPhone_(v) {
  var d = String(v == null ? "" : v).replace(/[^0-9]/g, "");
  if (d.length === 11) return d.slice(0, 3) + "-" + d.slice(3, 7) + "-" + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + "-" + d.slice(3, 6) + "-" + d.slice(6);
  return d;
}
function isPhone_(p) {
  return /^01[016789]-\d{3,4}-\d{4}$/.test(p);
}

/** 설정 시트 → {monthlyLimit, voucherAmount} */
function getConfig_() {
  var out = { monthlyLimit: DEFAULT_MONTHLY_LIMIT, voucherAmount: DEFAULT_VOUCHER_AMOUNT };
  var sh = sheet_(SHEET_CONF);
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  vals.forEach(function (v) {
    var k = String(v[0] || "").trim();
    var n = Number(v[1]);
    if (!isFinite(n)) return;
    if (k === "월한도") out.monthlyLimit = Math.max(0, Math.round(n));
    if (k === "1회한도") out.voucherAmount = Math.max(0, Math.round(n));
  });
  return out;
}

/** 크루 전체 로드 → [{row, crewId, name, phone, dept, status, limit, memo}] */
function loadCrews_() {
  var sh = sheet_(SHEET_CREW);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 8).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!String(v[1] || "").trim() && !String(v[3] || "").trim()) continue; // 빈 행
    out.push({
      row: i + 2,
      registeredAt: toIso_(v[0]),
      crewId: String(v[1] || "").trim(),
      name: String(v[2] || "").trim(),
      phone: normPhone_(v[3]),
      dept: String(v[4] || "").trim(),
      status: String(v[5] || "활성").trim() || "활성",
      limit: (v[6] === "" || v[6] == null) ? null : Math.max(0, Math.round(Number(v[6]) || 0)),
      memo: String(v[7] || "")
    });
  }
  return out;
}

function findCrewByPhone_(phone) {
  var list = loadCrews_();
  for (var i = 0; i < list.length; i++) if (list[i].phone === phone) return list[i];
  return null;
}
function findCrewById_(crewId) {
  var list = loadCrews_();
  for (var i = 0; i < list.length; i++) if (list[i].crewId === crewId) return list[i];
  return null;
}

/** 사용내역 전체 로드(정상 건만 옵션) */
function loadUsage_() {
  var sh = sheet_(SHEET_USAGE);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 14).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!String(v[1] || "").trim()) continue;
    out.push({
      row: i + 2,
      usedAt: toIso_(v[0]),
      usageId: String(v[1]),
      crewId: String(v[2]),
      name: String(v[3]),
      dept: String(v[4]),
      cafe: String(v[5]),
      amount: Number(v[6]) || 0,
      company: Number(v[7]) || 0,
      self: Number(v[8]) || 0,
      staff: String(v[9]),
      code: String(v[10]),
      ym: String(v[11]),
      status: String(v[12] || "정상").trim() || "정상",
      memo: String(v[13] || "")
    });
  }
  return out;
}

/** 특정 크루의 특정 월 사용 건(정상만) */
function usageOfMonth_(crewId, ym) {
  return loadUsage_().filter(function (u) {
    return u.crewId === crewId && u.ym === ym && u.status === "정상";
  });
}

function cafeList_() {
  var sh = sheet_(SHEET_CAFE);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return vals
    .filter(function (v) { return String(v[0] || "").trim(); })
    .map(function (v) {
      return { name: String(v[0]).trim(), status: String(v[1] || "활성").trim() || "활성" };
    });
}

function nextCrewId_() {
  var list = loadCrews_();
  var max = 0;
  list.forEach(function (c) {
    var m = /^CR(\d+)$/.exec(c.crewId);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return "CR" + ("000" + (max + 1)).slice(-3);
}

// ─────────────────────────────────────────────
// POST 라우팅
// ─────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var action = String(p.action || "");

    // 크루용
    if (action === "crew_login") return handleCrewLogin_(p);
    if (action === "issue_code") return handleIssueCode_(p);
    if (action === "check_code") return handleCheckCode_(p);
    if (action === "cancel_code") return handleCancelCode_(p);

    // 카페용
    if (action === "cafe_lookup") return handleCafeLookup_(p);
    if (action === "cafe_redeem") return handleCafeRedeem_(p);

    // 관리자용
    if (action === "admin_crew_add") return handleCrewAdd_(p);
    if (action === "admin_crew_update") return handleCrewUpdate_(p);
    if (action === "admin_crew_delete") return handleCrewDelete_(p);
    if (action === "admin_usage_status") return handleUsageStatus_(p);
    if (action === "admin_cafe_save") return handleCafeSave_(p);
    if (action === "admin_config_save") return handleConfigSave_(p);

    return json_({ result: "error", message: "unknown action" });
  } catch (err) {
    return json_({ result: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// 크루: 로그인(본인 확인)
// ─────────────────────────────────────────────
function handleCrewLogin_(p) {
  var phone = normPhone_(p.phone);
  if (!isPhone_(phone)) return json_({ result: "error", message: "휴대전화 번호 형식을 확인해 주세요" });

  var crew = findCrewByPhone_(phone);
  if (!crew) return json_({ result: "error", message: "등록되지 않은 번호예요. 관리자에게 등록을 요청해 주세요." });
  if (crew.status !== "활성") return json_({ result: "error", message: "현재 사용이 중지된 계정이에요. 관리자에게 문의해 주세요." });

  return json_(crewStatePayload_(crew));
}

/** 크루 화면이 필요로 하는 상태 한 덩어리 */
function crewStatePayload_(crew) {
  var conf = getConfig_();
  var limit = (crew.limit == null) ? conf.monthlyLimit : crew.limit;
  var thisYm = ym_();
  var all = loadUsage_().filter(function (u) { return u.crewId === crew.crewId && u.status === "정상"; });
  var thisMonth = all.filter(function (u) { return u.ym === thisYm; });

  var pending = null;
  var code = cache_().get("crew_" + crew.crewId);
  if (code) {
    var raw = cache_().get("code_" + code);
    if (raw) {
      var d = JSON.parse(raw);
      pending = { code: code, expiresAt: d.expiresAt };
    }
  }

  return {
    result: "ok",
    crew: { crewId: crew.crewId, name: crew.name, phone: crew.phone, dept: crew.dept },
    policy: { monthlyLimit: limit, voucherAmount: conf.voucherAmount },
    month: {
      ym: thisYm,
      used: thisMonth.length,
      remain: Math.max(0, limit - thisMonth.length),
      companyAmount: thisMonth.reduce(function (s, u) { return s + u.company; }, 0)
    },
    pending: pending,
    history: all.sort(function (a, b) { return a.usedAt < b.usedAt ? 1 : -1; }).slice(0, 30).map(function (u) {
      return { usedAt: u.usedAt, cafe: u.cafe, amount: u.amount, company: u.company, self: u.self, ym: u.ym };
    }),
    cafes: cafeList_().filter(function (c) { return c.status === "활성"; }).map(function (c) { return c.name; })
  };
}

// ─────────────────────────────────────────────
// 크루: 4자리 코드 발급
// ─────────────────────────────────────────────
function handleIssueCode_(p) {
  var phone = normPhone_(p.phone);
  var crew = findCrewByPhone_(phone);
  if (!crew) return json_({ result: "error", message: "등록되지 않은 번호예요" });
  if (crew.status !== "활성") return json_({ result: "error", message: "사용이 중지된 계정이에요" });

  var conf = getConfig_();
  var limit = (crew.limit == null) ? conf.monthlyLimit : crew.limit;
  var thisYm = ym_();
  var used = usageOfMonth_(crew.crewId, thisYm).length;
  if (used >= limit) {
    return json_({ result: "error", message: thisYm + " 사용 가능 횟수를 모두 사용했어요 (" + used + "/" + limit + "회)" });
  }

  var c = cache_();

  // 기존 발급 코드가 살아 있으면 그대로 돌려줌 (중복 발급 방지)
  var old = c.get("crew_" + crew.crewId);
  if (old) {
    var oraw = c.get("code_" + old);
    if (oraw) {
      var od = JSON.parse(oraw);
      return json_({ result: "ok", code: old, expiresAt: od.expiresAt, reused: true });
    }
  }

  var code = "";
  for (var t = 0; t < 30; t++) {
    var cand = ("000" + Math.floor(Math.random() * 10000)).slice(-4);
    if (!c.get("code_" + cand)) { code = cand; break; }
  }
  if (!code) return json_({ result: "error", message: "코드 발급이 혼잡해요. 잠시 후 다시 눌러주세요." });

  var expiresAt = new Date(new Date().getTime() + CODE_TTL * 1000);
  var payload = {
    crewId: crew.crewId, name: crew.name, phone: crew.phone, dept: crew.dept,
    issuedAt: toIso_(new Date()), expiresAt: toIso_(expiresAt)
  };
  c.put("code_" + code, JSON.stringify(payload), CODE_TTL);
  c.put("crew_" + crew.crewId, code, CODE_TTL);

  return json_({ result: "ok", code: code, expiresAt: toIso_(expiresAt), reused: false });
}

/** 크루 화면 폴링: 코드가 승인됐는지 확인 */
function handleCheckCode_(p) {
  var code = String(p.code || "").trim();
  if (!/^\d{4}$/.test(code)) return json_({ result: "error", message: "코드 형식 오류" });

  var c = cache_();
  var doneRaw = c.get("used_" + code);
  if (doneRaw) return json_({ result: "ok", status: "used", detail: JSON.parse(doneRaw) });

  var raw = c.get("code_" + code);
  if (!raw) return json_({ result: "ok", status: "expired" });

  var d = JSON.parse(raw);
  return json_({ result: "ok", status: "waiting", expiresAt: d.expiresAt });
}

/** 크루가 코드 발급을 취소 */
function handleCancelCode_(p) {
  var code = String(p.code || "").trim();
  var raw = cache_().get("code_" + code);
  if (raw) {
    var d = JSON.parse(raw);
    cache_().remove("code_" + code);
    cache_().remove("crew_" + d.crewId);
  }
  return json_({ result: "ok" });
}

// ─────────────────────────────────────────────
// 카페: 코드 조회 → 승인
// ─────────────────────────────────────────────
function requireCafeToken_(p) {
  var t = props_().getProperty("CAFE_TOKEN");
  return t && String(p.token || "") === t;
}

/** 승인 전 코드 확인 (누구 건지 보여주기) */
function handleCafeLookup_(p) {
  if (!requireCafeToken_(p)) return json_({ result: "error", message: "unauthorized" });
  var code = String(p.code || "").trim();
  if (!/^\d{4}$/.test(code)) return json_({ result: "error", message: "4자리 숫자를 입력해 주세요" });

  var raw = cache_().get("code_" + code);
  if (!raw) {
    if (cache_().get("used_" + code)) return json_({ result: "error", message: "이미 사용된 코드예요" });
    return json_({ result: "error", message: "유효하지 않거나 시간이 지난 코드예요" });
  }
  var d = JSON.parse(raw);
  var conf = getConfig_();
  var crew = findCrewById_(d.crewId);
  var limit = (crew && crew.limit != null) ? crew.limit : conf.monthlyLimit;
  var used = usageOfMonth_(d.crewId, ym_()).length;

  return json_({
    result: "ok",
    name: d.name, dept: d.dept, crewId: d.crewId,
    expiresAt: d.expiresAt,
    voucherAmount: conf.voucherAmount,
    monthUsed: used, monthLimit: limit
  });
}

function handleCafeRedeem_(p) {
  if (!requireCafeToken_(p)) return json_({ result: "error", message: "unauthorized" });

  var code = String(p.code || "").trim();
  var amount = Math.round(Number(p.amount));
  var cafe = String(p.cafe || "").trim().slice(0, 40);
  var staff = String(p.staff || "").trim().slice(0, 20);

  if (!/^\d{4}$/.test(code)) return json_({ result: "error", message: "4자리 숫자를 입력해 주세요" });
  if (!isFinite(amount) || amount <= 0) return json_({ result: "error", message: "결제 금액을 입력해 주세요" });
  if (amount > 500000) return json_({ result: "error", message: "금액이 너무 커요. 다시 확인해 주세요." });
  if (!cafe) return json_({ result: "error", message: "카페를 선택해 주세요" });

  var c = cache_();
  var raw = c.get("code_" + code);
  if (!raw) {
    if (c.get("used_" + code)) return json_({ result: "error", message: "이미 사용된 코드예요" });
    return json_({ result: "error", message: "유효하지 않거나 시간이 지난 코드예요" });
  }
  var d = JSON.parse(raw);

  var crew = findCrewById_(d.crewId);
  if (!crew) return json_({ result: "error", message: "크루 정보를 찾을 수 없어요" });
  if (crew.status !== "활성") return json_({ result: "error", message: "사용이 중지된 계정이에요" });

  var conf = getConfig_();
  var limit = (crew.limit == null) ? conf.monthlyLimit : crew.limit;
  var thisYm = ym_();
  var used = usageOfMonth_(crew.crewId, thisYm).length;
  if (used >= limit) {
    c.remove("code_" + code);
    c.remove("crew_" + crew.crewId);
    return json_({ result: "error", message: "이번 달 사용 가능 횟수를 모두 사용했어요 (" + used + "/" + limit + "회)" });
  }

  var company = Math.min(amount, conf.voucherAmount);
  var self = Math.max(0, amount - conf.voucherAmount);
  var usageId = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
  var now = new Date();

  sheet_(SHEET_USAGE).appendRow([
    now, usageId, crew.crewId, crew.name, crew.dept, cafe,
    amount, company, self, staff, code, thisYm, "정상", ""
  ]);

  // 코드 1회용 소진 + 크루 화면이 결과를 받아갈 수 있도록 짧게 보관
  c.remove("code_" + code);
  c.remove("crew_" + crew.crewId);
  c.put("used_" + code, JSON.stringify({
    usageId: usageId, name: crew.name, cafe: cafe, amount: amount,
    company: company, self: self, usedAt: toIso_(now)
  }), RESULT_TTL);

  return json_({
    result: "ok",
    usageId: usageId,
    name: crew.name, dept: crew.dept,
    amount: amount, company: company, self: self,
    monthUsed: used + 1, monthLimit: limit,
    usedAt: toIso_(now)
  });
}

// ─────────────────────────────────────────────
// 관리자
// ─────────────────────────────────────────────
function requireAdmin_(p) {
  var t = props_().getProperty("ADMIN_TOKEN");
  return t && String(p.token || "") === t;
}

function handleCrewAdd_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var name = String(p.name || "").trim().slice(0, 20);
  var phone = normPhone_(p.phone);
  var dept = String(p.dept || "").trim().slice(0, 30);
  var limit = (p.limit === "" || p.limit == null) ? "" : Math.max(0, Math.round(Number(p.limit) || 0));

  if (!name) return json_({ result: "error", message: "이름을 입력해 주세요" });
  if (!isPhone_(phone)) return json_({ result: "error", message: "휴대전화 번호 형식을 확인해 주세요" });
  if (findCrewByPhone_(phone)) return json_({ result: "error", message: "이미 등록된 번호예요" });

  var crewId = nextCrewId_();
  var sh = sheet_(SHEET_CREW);
  sh.appendRow([new Date(), crewId, name, phone, dept, "활성", limit, String(p.memo || "").slice(0, 100)]);
  sh.getRange(sh.getLastRow(), 4).setNumberFormat("@");
  SpreadsheetApp.flush();
  sh.getRange(sh.getLastRow(), 4).setValue(phone);

  return json_({ result: "ok", crewId: crewId });
}

function handleCrewUpdate_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var crew = findCrewById_(String(p.crewId || ""));
  if (!crew) return json_({ result: "error", message: "크루를 찾을 수 없어요" });

  var sh = sheet_(SHEET_CREW);
  var r = crew.row;

  if (p.name != null) sh.getRange(r, 3).setValue(String(p.name).trim().slice(0, 20));
  if (p.phone != null) {
    var ph = normPhone_(p.phone);
    if (!isPhone_(ph)) return json_({ result: "error", message: "휴대전화 번호 형식을 확인해 주세요" });
    var dup = findCrewByPhone_(ph);
    if (dup && dup.crewId !== crew.crewId) return json_({ result: "error", message: "다른 크루가 쓰는 번호예요" });
    sh.getRange(r, 4).setNumberFormat("@");
    SpreadsheetApp.flush();
    sh.getRange(r, 4).setValue(ph);
  }
  if (p.dept != null) sh.getRange(r, 5).setValue(String(p.dept).trim().slice(0, 30));
  if (p.status != null) sh.getRange(r, 6).setValue(String(p.status) === "비활성" ? "비활성" : "활성");
  if (p.limit != null) {
    sh.getRange(r, 7).setValue(String(p.limit) === "" ? "" : Math.max(0, Math.round(Number(p.limit) || 0)));
  }
  if (p.memo != null) sh.getRange(r, 8).setValue(String(p.memo).slice(0, 100));

  return json_({ result: "ok" });
}

function handleCrewDelete_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var crew = findCrewById_(String(p.crewId || ""));
  if (!crew) return json_({ result: "error", message: "크루를 찾을 수 없어요" });
  sheet_(SHEET_CREW).deleteRow(crew.row);
  return json_({ result: "ok" });
}

/** 잘못 승인된 건 취소/복구 */
function handleUsageStatus_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var usageId = String(p.usageId || "");
  var status = String(p.status || "") === "취소" ? "취소" : "정상";
  var list = loadUsage_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].usageId === usageId) {
      sheet_(SHEET_USAGE).getRange(list[i].row, 13).setValue(status);
      return json_({ result: "ok" });
    }
  }
  return json_({ result: "error", message: "사용 내역을 찾을 수 없어요" });
}

/** 카페 목록 통째로 저장 */
function handleCafeSave_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var rows = (p.cafes || []).filter(function (c) { return String(c.name || "").trim(); });
  var sh = sheet_(SHEET_CAFE);
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows.map(function (c) {
      return [String(c.name).trim().slice(0, 40), c.status === "비활성" ? "비활성" : "활성", String(c.memo || "").slice(0, 100)];
    }));
  }
  return json_({ result: "ok" });
}

function handleConfigSave_(p) {
  if (!requireAdmin_(p)) return json_({ result: "error", message: "unauthorized" });
  var sh = sheet_(SHEET_CONF);
  var want = {
    "월한도": Math.max(0, Math.round(Number(p.monthlyLimit))),
    "1회한도": Math.max(0, Math.round(Number(p.voucherAmount)))
  };
  var last = sh.getLastRow();
  var seen = {};
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var k = String(vals[i][0] || "").trim();
      if (want.hasOwnProperty(k) && isFinite(want[k])) {
        sh.getRange(i + 2, 2).setValue(want[k]);
        seen[k] = true;
      }
    }
  }
  Object.keys(want).forEach(function (k) {
    if (!seen[k] && isFinite(want[k])) sh.appendRow([k, want[k], ""]);
  });
  return json_({ result: "ok" });
}

// ─────────────────────────────────────────────
// GET 라우팅 (관리자 조회 / 카페 화면 초기 로드)
// ─────────────────────────────────────────────
function doGet(e) {
  var action = String((e && e.parameter && e.parameter.action) || "").toLowerCase();
  if (action === "ping") return json_({ result: "ok" });

  // 카페 화면: 카페 목록 + 오늘 승인 내역
  if (action === "cafe_init") {
    if (String(e.parameter.token || "") !== props_().getProperty("CAFE_TOKEN")) {
      return json_({ result: "error", message: "unauthorized" });
    }
    var today = ymd_();
    var mine = String(e.parameter.cafe || "");
    var todays = loadUsage_().filter(function (u) {
      return u.usedAt.slice(0, 10) === today && (!mine || u.cafe === mine) && u.status === "정상";
    }).sort(function (a, b) { return a.usedAt < b.usedAt ? 1 : -1; });

    return json_({
      result: "ok",
      cafes: cafeList_().filter(function (c) { return c.status === "활성"; }).map(function (c) { return c.name; }),
      voucherAmount: getConfig_().voucherAmount,
      today: todays.map(function (u) {
        return { usedAt: u.usedAt, name: u.name, cafe: u.cafe, amount: u.amount, company: u.company, self: u.self };
      })
    });
  }

  // 관리자
  var token = props_().getProperty("ADMIN_TOKEN");
  if (!token || String((e.parameter && e.parameter.token) || "") !== token) {
    return json_({ result: "error", message: "unauthorized" });
  }

  if (action === "admin_init") {
    var conf = getConfig_();
    var crews = loadCrews_();
    var usage = loadUsage_();
    var thisYm = ym_();

    var byCrewMonth = {};
    usage.forEach(function (u) {
      if (u.status !== "정상") return;
      var k = u.crewId + "|" + u.ym;
      if (!byCrewMonth[k]) byCrewMonth[k] = { count: 0, amount: 0, company: 0 };
      byCrewMonth[k].count++;
      byCrewMonth[k].amount += u.amount;
      byCrewMonth[k].company += u.company;
    });

    var months = {};
    usage.forEach(function (u) { if (u.ym) months[u.ym] = true; });
    months[thisYm] = true;

    return json_({
      result: "ok",
      config: conf,
      thisYm: thisYm,
      months: Object.keys(months).sort().reverse(),
      cafes: cafeList_(),
      cafeToken: props_().getProperty("CAFE_TOKEN"),
      crews: crews.map(function (c) {
        var k = c.crewId + "|" + thisYm;
        var s = byCrewMonth[k] || { count: 0, amount: 0, company: 0 };
        return {
          crewId: c.crewId, name: c.name, phone: c.phone, dept: c.dept,
          status: c.status, limit: c.limit, memo: c.memo,
          registeredAt: c.registeredAt,
          thisMonthCount: s.count, thisMonthAmount: s.amount, thisMonthCompany: s.company
        };
      }),
      usage: usage.sort(function (a, b) { return a.usedAt < b.usedAt ? 1 : -1; }).map(function (u) {
        return {
          usageId: u.usageId, usedAt: u.usedAt, crewId: u.crewId, name: u.name, dept: u.dept,
          cafe: u.cafe, amount: u.amount, company: u.company, self: u.self,
          staff: u.staff, ym: u.ym, status: u.status
        };
      })
    });
  }

  return json_({ result: "error", message: "unknown action" });
}
