/**
 * EASTEND 오프라인 멤버십 v2 — Apps Script 백엔드
 * 회원가입(웰컴 3,000P) + 마일리지 조회/차감 — 구매 적립(1%)은 외부 판매 프로그램에서 처리
 *
 * 시트: 회원 / 영수증 / 포인트   (개인정보 포함 — 링크 공유 금지)
 * 코드 업데이트 후: setup2 1회 실행 → 배포 > 배포 관리 > 수정 > 새 버전
 */

var SHEET_MEMBER = "회원";
var SHEET_RECEIPT = "영수증";
var SHEET_POINT = "포인트";
var EARN_RATE = 0.01;        // 적립률 1%
var WELCOME_POINTS = 3000;   // 신규가입 축하 적립
var REDEEM_MIN = 5000;     // 최소 사용 포인트
var REDEEM_UNIT = 1000;    // 사용 단위
var EXPIRE_DAYS = 365;     // 적립 후 소멸까지 일수 (1년)

// ─────────────────────────────────────────────
// 최초 1회: 시트 준비 + 관리자 토큰
// ─────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MEMBER) || ss.insertSheet(SHEET_MEMBER);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "가입일시", "브랜드", "매장코드", "매장명", "이름", "휴대전화",
      "생년월일", "성별", "개인정보동의", "마케팅동의", "동의일시",
      "최근방문", "방문횟수", "UserAgent", "야간광고동의"
    ]);
    sh.setFrozenRows(1);
  }
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("ADMIN_TOKEN");
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, "").slice(0, 20);
    props.setProperty("ADMIN_TOKEN", token);
  }
  Logger.log("관리자 토큰: " + token);
}

// ─────────────────────────────────────────────
// v2 업그레이드 1회 실행: 영수증/포인트 시트 + 매장 토큰 + 드라이브 폴더 + 자동확정 트리거
// ─────────────────────────────────────────────
function setup2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var r = ss.getSheetByName(SHEET_RECEIPT) || ss.insertSheet(SHEET_RECEIPT);
  if (r.getLastRow() === 0) {
    r.appendRow([
      "제출일시", "전화번호", "이름", "브랜드", "매장코드", "매장명",
      "영수증번호", "결제금액", "적립예정포인트", "적립예정일",
      "상태", "확정일시", "사진링크", "제출경로", "메모"
    ]);
    r.setFrozenRows(1);
  }

  var p = ss.getSheetByName(SHEET_POINT) || ss.insertSheet(SHEET_POINT);
  if (p.getLastRow() === 0) {
    p.appendRow([
      "일시", "전화번호", "이름", "유형", "포인트",
      "관련영수증번호", "매장코드", "매장명", "메모"
    ]);
    p.setFrozenRows(1);
  }

  var props = PropertiesService.getScriptProperties();
  var staffToken = props.getProperty("STAFF_TOKEN");
  if (!staffToken) {
    staffToken = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
    props.setProperty("STAFF_TOKEN", staffToken);
  }

  // 영수증 사진 저장 폴더
  var folderId = props.getProperty("PHOTO_FOLDER_ID");
  if (!folderId) {
    var folder = DriveApp.createFolder("멤버십_영수증사진");
    props.setProperty("PHOTO_FOLDER_ID", folder.getId());
  }

  // 매일 새벽 3시 적립 자동확정 트리거 (중복 생성 방지)
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === "confirmReceipts";
  });
  if (!has) {
    ScriptApp.newTrigger("confirmReceipts").timeBased().everyDays(1).atHour(3).create();
  }

  Logger.log("관리자 토큰: " + props.getProperty("ADMIN_TOKEN"));
  Logger.log("매장용 토큰: " + staffToken);
  Logger.log("v2 준비 완료 (영수증/포인트 시트, 사진 폴더, 자동확정 트리거)");
}

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function props_() {
  return PropertiesService.getScriptProperties();
}
function maskName_(name) {
  var s = String(name || "");
  if (s.length <= 1) return s;
  if (s.length === 2) return s.charAt(0) + "*";
  return s.charAt(0) + Array(s.length - 1).join("*") + s.charAt(s.length - 1);
}

// 가입 당일 여부 확인 (영수증 적립은 가입 익일부터 허용)
function joinedToday_(memberRow) {
  var joined = sheet_(SHEET_MEMBER).getRange(memberRow, 1).getValue();
  if (!(joined instanceof Date)) return false;
  var fmt = function (d) { return Utilities.formatDate(d, "Asia/Seoul", "yyyyMMdd"); };
  return fmt(joined) === fmt(new Date());
}

// 회원 찾기: [행번호, 이름] (없으면 행번호 0)
function findMember_(phone) {
  var sh = sheet_(SHEET_MEMBER);
  var last = sh.getLastRow();
  if (last < 2) return [0, ""];
  var vals = sh.getRange(2, 5, last - 1, 2).getValues(); // E:이름, F:전화
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][1]).trim() === phone) return [i + 2, String(vals[i][0])];
  }
  return [0, ""];
}

// 포인트 잔액(확정분) 계산
function getBalance_(phone) {
  var sh = sheet_(SHEET_POINT);
  var last = sh.getLastRow();
  var bal = 0;
  if (last < 2) return 0;
  var vals = sh.getRange(2, 2, last - 1, 4).getValues(); // B:전화, E:포인트(5번째지만 range로 B~E)
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === phone) bal += Number(vals[i][3]) || 0;
  }
  return bal;
}

// 전체 잔액 맵 (관리자 목록용)
function balanceMap_() {
  var sh = sheet_(SHEET_POINT);
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;
  var vals = sh.getRange(2, 2, last - 1, 4).getValues();
  for (var i = 0; i < vals.length; i++) {
    var ph = String(vals[i][0]).trim();
    map[ph] = (map[ph] || 0) + (Number(vals[i][3]) || 0);
  }
  return map;
}

// 대기중(미확정) 적립 합계
function getPending_(phone) {
  var sh = sheet_(SHEET_RECEIPT);
  var last = sh.getLastRow();
  var sum = 0;
  if (last < 2) return 0;
  var vals = sh.getRange(2, 2, last - 1, 10).getValues(); // B전화 ~ K상태
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === phone && String(vals[i][9]) === "대기") {
      sum += Number(vals[i][7]) || 0; // 적립예정포인트
    }
  }
  return sum;
}

// 영수증번호 중복 확인
function receiptExists_(no) {
  var sh = sheet_(SHEET_RECEIPT);
  var last = sh.getLastRow();
  if (last < 2) return false;
  var vals = sh.getRange(2, 7, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === no) return true;
  }
  return false;
}

// 사진 저장 → URL
function savePhoto_(b64, mime, receiptNo) {
  try {
    if (!b64) return "";
    var folderId = props_().getProperty("PHOTO_FOLDER_ID");
    if (!folderId) return "";
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || "image/jpeg",
      "영수증_" + receiptNo + "_" + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss") + ".jpg");
    var file = DriveApp.getFolderById(folderId).createFile(blob);
    return file.getUrl();
  } catch (err) {
    return "저장실패:" + String(err).slice(0, 80);
  }
}

// 영수증 등록 공통 — 등록 즉시 포인트 적립
function addReceipt_(o) {
  var no = String(o.receiptNo || "").trim();
  var amt = Math.round(Number(String(o.amount || "").replace(/[^0-9]/g, "")));
  if (!no) return { ok: false, message: "영수증 번호 누락" };
  if (!amt || amt <= 0) return { ok: false, message: "결제금액 오류" };
  if (receiptExists_(no)) return { ok: false, message: "이미 등록된 영수증입니다" };

  var pts = Math.floor(amt * EARN_RATE);
  var now = new Date();
  var photoUrl = savePhoto_(o.photo, o.photoType, no);

  sheet_(SHEET_RECEIPT).appendRow([
    now, o.phone, o.name, o.brand || "", o.storeId || "", o.storeName || "",
    no, amt, pts, now, "확정", now, photoUrl, o.via || "", ""
  ]);
  var r = sheet_(SHEET_RECEIPT).getLastRow();
  sheet_(SHEET_RECEIPT).getRange(r, 7).setNumberFormat("@"); // 영수증번호 텍스트

  sheet_(SHEET_POINT).appendRow([
    now, o.phone, o.name, "적립", pts,
    no, o.storeId || "", o.storeName || "", "영수증 즉시 적립"
  ]);

  return { ok: true, points: pts, immediate: true };
}

// ─────────────────────────────────────────────
// 매일 새벽 트리거: 1년 경과 포인트 소멸
// (즉시 적립으로 전환되어 '대기' 확정 처리는 없음 — 과거 대기 건이 있으면 함께 확정)
// ─────────────────────────────────────────────
function confirmReceipts() {
  var sh = sheet_(SHEET_RECEIPT);
  var last = sh.getLastRow();
  var now = new Date();
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 11).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][10]) === "대기") {
        var row = i + 2;
        sheet_(SHEET_POINT).appendRow([
          now, vals[i][1], vals[i][2], "적립", Number(vals[i][8]) || 0,
          String(vals[i][6]), vals[i][4], vals[i][5], "과거 대기분 확정"
        ]);
        sh.getRange(row, 11).setValue("확정");
        sh.getRange(row, 12).setValue(now);
      }
    }
  }
  expirePoints();
}

// 적립 후 1년(EXPIRE_DAYS) 경과한 미사용 포인트 자동 소멸 (선입선출 방식)
function expirePoints() {
  var sh = sheet_(SHEET_POINT);
  var last = sh.getLastRow();
  if (last < 2) return;
  var now = new Date();
  var cutoff = new Date(now.getTime() - EXPIRE_DAYS * 86400000);
  var vals = sh.getRange(2, 1, last - 1, 5).getValues(); // A일시~E포인트

  // 전화번호별: 1년 지난 적립 누계 vs 사용·소멸 누계 (선입선출: 사용분은 오래된 적립부터 차감된 것으로 간주)
  var oldEarn = {}, spent = {}, names = {};
  for (var i = 0; i < vals.length; i++) {
    var ph = String(vals[i][1]).trim();
    var pts = Number(vals[i][4]) || 0;
    names[ph] = names[ph] || String(vals[i][2] || "");
    if (pts > 0) {
      var d = vals[i][0];
      if (d instanceof Date && d <= cutoff) oldEarn[ph] = (oldEarn[ph] || 0) + pts;
    } else if (pts < 0) {
      spent[ph] = (spent[ph] || 0) + (-pts);
    }
  }
  for (var ph2 in oldEarn) {
    var toExpire = oldEarn[ph2] - (spent[ph2] || 0);
    if (toExpire > 0) {
      sh.appendRow([now, ph2, names[ph2], "소멸", -toExpire, "", "", "",
        "적립 후 " + EXPIRE_DAYS + "일 경과 자동 소멸"]);
    }
  }
}

// ─────────────────────────────────────────────
// POST
// ─────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var p = JSON.parse(e.postData.contents || "{}");
    var action = String(p.action || "");

    if (action === "join") return handleJoin_(p);
    if (action === "staff_lookup") return handleStaffLookup_(p);
    if (action === "staff_redeem") return handleStaffRedeem_(p);
    // 영수증 적립은 폐지 — 구매 적립(1%)은 판매 프로그램에서 자동 처리
    if (action === "receipt" || action === "staff_earn") {
      return json_({ result: "error", message: "구매 적립(1%)은 판매 프로그램에서 자동 처리됩니다" });
    }

    return json_({ result: "error", message: "unknown action" });
  } catch (err) {
    return json_({ result: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function handleJoin_(p) {
  var name = String(p.name || "").trim().slice(0, 20);
  var phone = String(p.phone || "").trim();
  if (!name) return json_({ result: "error", message: "이름 누락" });
  if (!/^01[016789]-\d{3,4}-\d{4}$/.test(phone)) return json_({ result: "error", message: "전화번호 형식 오류" });
  if (p.consentPrivacy !== true) return json_({ result: "error", message: "필수 동의 누락" });

  var sh = sheet_(SHEET_MEMBER);
  var now = new Date();
  var found = findMember_(phone);
  var status;

  if (found[0]) {
    var cnt = Number(sh.getRange(found[0], 13).getValue()) || 1;
    sh.getRange(found[0], 12).setValue(now);
    sh.getRange(found[0], 13).setValue(cnt + 1);
    if (p.consentMarketing === true) {
      sh.getRange(found[0], 10).setValue("Y");
      sh.getRange(found[0], 11).setValue(now);
    }
    if (p.consentNight === true) {
      sh.getRange(found[0], 15).setValue("Y");
    }
    status = "existing";
    name = found[1]; // 시트 등록명 기준
  } else {
    sh.appendRow([
      now, String(p.brand || ""), String(p.storeId || ""), String(p.storeName || ""),
      name, phone, String(p.birth || ""), String(p.gender || ""),
      "Y", p.consentMarketing === true ? "Y" : "N", now, now, 1,
      String(p.ua || "").slice(0, 200),
      p.consentNight === true ? "Y" : "N"
    ]);
    var r = sh.getLastRow();
    sh.getRange(r, 6, 1, 2).setNumberFormat("@");
    status = "new";
    // 신규가입 축하 포인트 즉시 지급
    if (WELCOME_POINTS > 0) {
      sheet_(SHEET_POINT).appendRow([
        now, phone, name, "적립", WELCOME_POINTS,
        "", String(p.storeId || ""), String(p.storeName || ""), "신규가입 축하 적립"
      ]);
    }
  }

  // 영수증 적립은 가입 익일부터 — 가입 요청에 딸려온 영수증은 접수하지 않음
  var receipt = null;
  if (p.receiptNo) {
    receipt = { ok: false, message: "영수증 적립은 회원가입 익일부터 가능합니다" };
  }

  return json_({ result: "ok", status: status, receipt: receipt });
}

// 기존 회원 영수증 적립 신청 (receipt.html)
function handleReceipt_(p) {
  var phone = String(p.phone || "").trim();
  if (!/^01[016789]-\d{3,4}-\d{4}$/.test(phone)) return json_({ result: "error", message: "전화번호 형식 오류" });
  var found = findMember_(phone);
  if (!found[0]) return json_({ result: "error", message: "가입되지 않은 번호입니다. 먼저 멤버십에 가입해주세요." });
  if (joinedToday_(found[0])) return json_({ result: "error", message: "영수증 적립은 회원가입 익일부터 가능합니다 (가입 당일 적립 불가)" });

  var receipt = addReceipt_({
    phone: phone, name: found[1], brand: p.brand, storeId: p.storeId,
    storeName: p.storeName, receiptNo: p.receiptNo, amount: p.amount,
    photo: p.photo, photoType: p.photoType, via: "영수증폼"
  });
  if (!receipt.ok) return json_({ result: "error", message: receipt.message });
  return json_({ result: "ok", name: maskName_(found[1]), receipt: receipt });
}

// 매장: 전화번호 조회
function handleStaffLookup_(p) {
  if (p.staffToken !== props_().getProperty("STAFF_TOKEN")) return json_({ result: "error", message: "unauthorized" });
  var phone = String(p.phone || "").trim();
  var found = findMember_(phone);
  if (!found[0]) return json_({ result: "ok", exists: false });

  var sh = sheet_(SHEET_MEMBER);
  return json_({
    result: "ok", exists: true,
    name: maskName_(found[1]),
    balance: getBalance_(phone),
    pending: getPending_(phone),
    joinedAt: toIso_(sh.getRange(found[0], 1).getValue()).slice(0, 10),
    visitCount: sh.getRange(found[0], 13).getValue()
  });
}

// 매장: 영수증 적립 등록
function handleStaffEarn_(p) {
  if (p.staffToken !== props_().getProperty("STAFF_TOKEN")) return json_({ result: "error", message: "unauthorized" });
  var phone = String(p.phone || "").trim();
  var found = findMember_(phone);
  if (!found[0]) return json_({ result: "error", message: "가입되지 않은 번호입니다" });
  if (joinedToday_(found[0])) return json_({ result: "error", message: "영수증 적립은 회원가입 익일부터 가능합니다 (가입 당일 적립 불가)" });

  var receipt = addReceipt_({
    phone: phone, name: found[1], brand: p.brand, storeId: p.storeId,
    storeName: p.storeName, receiptNo: p.receiptNo, amount: p.amount,
    photo: p.photo, photoType: p.photoType, via: "매장"
  });
  if (!receipt.ok) return json_({ result: "error", message: receipt.message });
  return json_({ result: "ok", receipt: receipt, balance: getBalance_(phone) });
}

// 매장: 마일리지 차감 (에누리)
function handleStaffRedeem_(p) {
  if (p.staffToken !== props_().getProperty("STAFF_TOKEN")) return json_({ result: "error", message: "unauthorized" });

  var amt = Math.round(Number(String(p.amount || "").replace(/[^0-9]/g, "")));
  if (!amt || amt <= 0) return json_({ result: "error", message: "차감 금액 오류" });
  if (amt < REDEEM_MIN) return json_({ result: "error", message: REDEEM_MIN.toLocaleString() + "P부터 사용 가능합니다" });
  if (amt % REDEEM_UNIT !== 0) return json_({ result: "error", message: REDEEM_UNIT.toLocaleString() + "P 단위로만 사용 가능합니다" });

  var phone = String(p.phone || "").trim();
  var found = findMember_(phone);
  if (!found[0]) return json_({ result: "error", message: "가입되지 않은 번호입니다" });

  var bal = getBalance_(phone);
  if (amt > bal) return json_({ result: "error", message: "잔액 부족 (현재 " + bal + "P)" });

  sheet_(SHEET_POINT).appendRow([
    new Date(), phone, found[1], "사용", -amt, "",
    String(p.storeId || ""), String(p.storeName || ""),
    "매장 에누리 차감" + (p.memo ? " / " + String(p.memo).slice(0, 50) : "")
  ]);
  return json_({ result: "ok", used: amt, balance: bal - amt });
}

// ─────────────────────────────────────────────
// GET (관리자)
// ─────────────────────────────────────────────
function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  if (action === "ping") return json_({ result: "ok" });

  var token = props_().getProperty("ADMIN_TOKEN");
  if (!token || e.parameter.token !== token) return json_({ result: "error", message: "unauthorized" });

  if (action === "list") {
    var sh = sheet_(SHEET_MEMBER);
    var last = sh.getLastRow();
    var rows = [];
    if (last >= 2) {
      var bmap = balanceMap_();
      var vals = sh.getRange(2, 1, last - 1, 13).getValues();
      rows = vals.map(function (v) {
        var ph = String(v[5]);
        return {
          joinedAt: toIso_(v[0]), brand: v[1], storeId: v[2], storeName: v[3],
          name: v[4], phone: ph, birth: String(v[6]), gender: v[7],
          consentPrivacy: v[8], consentMarketing: v[9],
          lastVisit: toIso_(v[11]), visitCount: v[12],
          balance: bmap[ph] || 0
        };
      });
    }
    return json_({ result: "ok", rows: rows });
  }

  if (action === "receipts") {
    var sh2 = sheet_(SHEET_RECEIPT);
    var last2 = sh2.getLastRow();
    var rows2 = [];
    if (last2 >= 2) {
      var vals2 = sh2.getRange(2, 1, last2 - 1, 14).getValues();
      rows2 = vals2.map(function (v) {
        return {
          submittedAt: toIso_(v[0]), phone: String(v[1]), name: v[2],
          brand: v[3], storeName: v[5], receiptNo: String(v[6]),
          amount: v[7], points: v[8], dueDate: toIso_(v[9]).slice(0, 10),
          status: v[10], confirmedAt: toIso_(v[11]), photoUrl: v[12], via: v[13]
        };
      });
    }
    return json_({ result: "ok", rows: rows2 });
  }

  return json_({ result: "error", message: "unknown action" });
}

function toIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd'T'HH:mm:ss");
  return String(v || "");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
