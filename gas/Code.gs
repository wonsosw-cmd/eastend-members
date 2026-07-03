/**
 * EASTEND 오프라인 멤버십 — Apps Script 백엔드
 *
 * [설치 순서]
 * 1. 새 Google 스프레드시트 생성 (이름 예: "오프라인 멤버십 회원DB") — 반드시 비공개 유지
 * 2. 확장 프로그램 > Apps Script > 이 코드 전체 붙여넣기
 * 3. 함수 setup 실행 1회 (시트 헤더 생성 + 관리자 토큰 발급, 로그에 토큰 출력됨)
 * 4. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자 (익명 포함) ← QR로 접속하는 고객이 로그인 없이 제출해야 하므로 필수
 * 5. 발급된 웹앱 URL(…/exec)을 GitHub 레포 docs/config.js 의 API_URL에 붙여넣고 push
 *
 * 시트에는 개인정보가 저장되므로 절대 링크 공유하지 말 것.
 */

var SHEET_NAME = "회원";

// ─────────────────────────────────────────────
// 최초 1회 실행: 시트 준비 + 관리자 토큰 발급
// ─────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "가입일시", "브랜드", "매장코드", "매장명", "이름", "휴대전화",
      "생년월일", "성별", "개인정보동의", "마케팅동의", "동의일시",
      "최근방문", "방문횟수", "UserAgent"
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
  Logger.log("이 토큰을 admin.html 접속 시 입력하세요.");
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// ─────────────────────────────────────────────
// POST: 가입 처리 (고객 폼에서 호출)
// ─────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // 동시 제출 대비
  try {
    var p = JSON.parse(e.postData.contents || "{}");
    if (p.action !== "join") return json_({ result: "error", message: "unknown action" });

    // 서버측 검증
    var name = String(p.name || "").trim().slice(0, 20);
    var phone = String(p.phone || "").trim();
    if (!name) return json_({ result: "error", message: "이름 누락" });
    if (!/^01[016789]-\d{3,4}-\d{4}$/.test(phone)) return json_({ result: "error", message: "전화번호 형식 오류" });
    if (p.consentPrivacy !== true) return json_({ result: "error", message: "필수 동의 누락" });

    var sh = getSheet_();
    var now = new Date();

    // 전화번호로 중복 확인 (F열)
    var last = sh.getLastRow();
    var existingRow = 0;
    if (last >= 2) {
      var phones = sh.getRange(2, 6, last - 1, 1).getValues();
      for (var i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).trim() === phone) { existingRow = i + 2; break; }
      }
    }

    if (existingRow) {
      // 기존 회원 → 최근방문(L), 방문횟수(M) 업데이트
      var cnt = Number(sh.getRange(existingRow, 13).getValue()) || 1;
      sh.getRange(existingRow, 12).setValue(now);
      sh.getRange(existingRow, 13).setValue(cnt + 1);
      // 마케팅 동의를 새로 한 경우만 갱신 (동의 철회는 폼으로 처리하지 않음)
      if (p.consentMarketing === true) {
        sh.getRange(existingRow, 10).setValue("Y");
        sh.getRange(existingRow, 11).setValue(now);
      }
      return json_({ result: "ok", status: "existing" });
    }

    sh.appendRow([
      now,                                   // 가입일시
      String(p.brand || ""),                 // 브랜드
      String(p.storeId || ""),               // 매장코드
      String(p.storeName || ""),             // 매장명
      name,                                  // 이름
      phone,                                 // 휴대전화
      String(p.birth || ""),                 // 생년월일(YYMMDD)
      String(p.gender || ""),                // 성별
      "Y",                                   // 개인정보동의 (필수라 항상 Y)
      p.consentMarketing === true ? "Y" : "N", // 마케팅동의
      now,                                   // 동의일시
      now,                                   // 최근방문
      1,                                     // 방문횟수
      String(p.ua || "").slice(0, 200)       // UserAgent (동의 증적 보조)
    ]);
    // 전화/생년월일이 숫자로 변형되지 않도록 텍스트 서식
    var r = sh.getLastRow();
    sh.getRange(r, 6, 1, 2).setNumberFormat("@");

    return json_({ result: "ok", status: "new" });
  } catch (err) {
    return json_({ result: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// GET: 관리자 조회 (admin.html에서 호출, 토큰 필수)
// ─────────────────────────────────────────────
function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  if (action === "ping") return json_({ result: "ok" });

  var token = PropertiesService.getScriptProperties().getProperty("ADMIN_TOKEN");
  if (!token || e.parameter.token !== token) {
    return json_({ result: "error", message: "unauthorized" });
  }

  if (action === "list") {
    var sh = getSheet_();
    var last = sh.getLastRow();
    var rows = [];
    if (last >= 2) {
      var vals = sh.getRange(2, 1, last - 1, 13).getValues();
      rows = vals.map(function (v) {
        return {
          joinedAt: toIso_(v[0]), brand: v[1], storeId: v[2], storeName: v[3],
          name: v[4], phone: String(v[5]), birth: String(v[6]), gender: v[7],
          consentPrivacy: v[8], consentMarketing: v[9],
          lastVisit: toIso_(v[11]), visitCount: v[12]
        };
      });
    }
    return json_({ result: "ok", rows: rows });
  }

  return json_({ result: "error", message: "unknown action" });
}

function toIso_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(v || "");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
