// 크루 카페 사용권 — 공용 API 헬퍼
(function () {
  function apiUrl() {
    var u = (window.CREW_CONFIG && window.CREW_CONFIG.API_URL) || "";
    if (!u) throw new Error("config.js의 API_URL이 비어 있어요. Apps Script 웹앱 URL을 넣어주세요.");
    return u;
  }

  // POST는 text/plain으로 보내 CORS preflight를 피한다 (Apps Script 제약)
  async function post(payload) {
    var res = await fetch(apiUrl(), {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });
    return res.json();
  }

  async function get(params) {
    var q = Object.keys(params)
      .map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]); })
      .join("&");
    var res = await fetch(apiUrl() + "?" + q);
    return res.json();
  }

  function won(n) {
    return (Number(n) || 0).toLocaleString("ko-KR") + "원";
  }

  function normPhone(v) {
    var d = String(v || "").replace(/[^0-9]/g, "").slice(0, 11);
    if (d.length <= 3) return d;
    if (d.length <= 7) return d.slice(0, 3) + "-" + d.slice(3);
    return d.slice(0, 3) + "-" + d.slice(3, 7) + "-" + d.slice(7);
  }

  // "2026-08-07T13:05:00" (KST 문자열) → 표시용
  function fmtDateTime(iso) {
    if (!iso) return "";
    var s = String(iso).replace("T", " ");
    return s.slice(5, 16).replace("-", "/");
  }
  function fmtDate(iso) {
    if (!iso) return "";
    return String(iso).slice(5, 10).replace("-", "/");
  }

  function showMsg(el, text, kind) {
    if (!el) return;
    el.textContent = text || "";
    el.className = "msg show " + (kind === "ok" ? "msg-ok" : "msg-error");
    if (!text) el.className = "msg";
  }
  function clearMsg(el) {
    if (el) el.className = "msg";
  }

  window.CrewAPI = { post: post, get: get, won: won, normPhone: normPhone, fmtDateTime: fmtDateTime, fmtDate: fmtDate, showMsg: showMsg, clearMsg: clearMsg };
})();
