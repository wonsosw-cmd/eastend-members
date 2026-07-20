// 매장 마스터 — _ETL_Masters_Blob.json storeLookup 기준 (오프라인만)
// id = 재무코드, erp = ERP 매장코드(회원등록 CSV의 '가입매장'에 사용)
// 매장 추가/삭제 시 이 파일만 수정하면 폼/QR/관리자 모두 반영됨.
window.EE_STORES = [
  // 시티브리즈 오프라인
  { id: "C001", erp: "M0001", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 판교" },
  { id: "C002", erp: "M0005", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 중동" },
  { id: "C003", erp: "M0010", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 무역센터" },
  { id: "C004", erp: "M0002", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 송도" },
  { id: "C005", erp: "M0004", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "플래그십" },
  { id: "C006", erp: "M0006", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 파주" },
  { id: "C007", erp: "M0011", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 부산" },
  { id: "C008", erp: "M0008", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 부산점" },
  { id: "C009", erp: "M0009", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 천호" },
  { id: "C010", erp: "M0003", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "더현대 여의도" },
  { id: "C011", erp: "M5157", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 동부산" },
  { id: "C012", erp: "M5172", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "더팩토리" },
  { id: "C013", erp: "M5183", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 킨텍스" },
  { id: "C014", erp: "M5184", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 울산" },
  { id: "C015", erp: "M5187", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 인천" },
  { id: "C016", erp: "M5188", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 의정부" },
  { id: "C017", erp: "M5189", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 대전" },
  { id: "C018", erp: "M5194", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 잠실점" },
  { id: "C019", erp: "M5202", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 대전점" },
  { id: "C020", erp: "M5203", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 노원점" },
  { id: "C021", erp: "M5204", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 안산점" },
  // 아티드 오프라인
  { id: "A001", erp: "M0013", brand: "ARTID", brandKo: "아티드", name: "신세계 대구" },
  { id: "A002", erp: "M0012", brand: "ARTID", brandKo: "아티드", name: "신세계 타임스퀘어" },
  { id: "A003", erp: "M5167", brand: "ARTID", brandKo: "아티드", name: "롯데 부산" },
  { id: "A004", erp: "M5185", brand: "ARTID", brandKo: "아티드", name: "현대 판교" },
  { id: "A005", erp: "M5186", brand: "ARTID", brandKo: "아티드", name: "신세계 의정부" },
  { id: "A006", erp: "M5172", brand: "ARTID", brandKo: "아티드", name: "더팩토리" },
  { id: "A007", erp: "M0014", brand: "ARTID", brandKo: "아티드", name: "신세계 광주" },
];

window.EE_STORE_BY_ID = Object.fromEntries(window.EE_STORES.map(s => [s.id, s]));
