// 매장 마스터 — _ETL_Masters_Blob.json storeLookup 기준 (오프라인만)
// id = 재무코드. 매장 추가/삭제 시 이 파일만 수정하면 폼/QR/관리자 모두 반영됨.
window.EE_STORES = [
  // 시티브리즈 오프라인
  { id: "C001", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 판교" },
  { id: "C002", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 중동" },
  { id: "C003", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 무역센터" },
  { id: "C004", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 송도" },
  { id: "C005", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "플래그십" },
  { id: "C006", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 파주" },
  { id: "C007", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 부산" },
  { id: "C008", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 부산점" },
  { id: "C009", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 천호" },
  { id: "C010", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "더현대 여의도" },
  { id: "C011", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 동부산" },
  { id: "C012", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "더팩토리" },
  { id: "C013", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 킨텍스" },
  { id: "C014", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "현대 울산" },
  { id: "C015", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 인천" },
  { id: "C016", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 의정부" },
  { id: "C017", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "신세계 대전" },
  { id: "C018", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 잠실점" },
  { id: "C019", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 대전점" },
  { id: "C020", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 노원점" },
  { id: "C021", brand: "CITYBREEZE", brandKo: "시티브리즈", name: "롯데 안산점" },
  // 아티드 오프라인
  { id: "A001", brand: "ARTID", brandKo: "아티드", name: "신세계 대구" },
  { id: "A002", brand: "ARTID", brandKo: "아티드", name: "신세계 타임스퀘어" },
  { id: "A003", brand: "ARTID", brandKo: "아티드", name: "롯데 부산" },
  { id: "A004", brand: "ARTID", brandKo: "아티드", name: "현대 판교" },
  { id: "A005", brand: "ARTID", brandKo: "아티드", name: "신세계 의정부" },
  { id: "A006", brand: "ARTID", brandKo: "아티드", name: "더팩토리" },
  { id: "A007", brand: "ARTID", brandKo: "아티드", name: "신세계 광주" },
];

window.EE_STORE_BY_ID = Object.fromEntries(window.EE_STORES.map(s => [s.id, s]));
