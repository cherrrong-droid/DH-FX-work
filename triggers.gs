/**
 * 통합 트리거 컨트롤러 (Section 5 기반)
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("자동화도구")
    .addItem("ODD 자동 세팅", "odd_runODDSetupFinal")
    .addItem("스프레드 환율 계산", "odd_runSpreadStartingCalculation")
    .addToUi();
}

/** 
 * [마스터 onEdit] 
 * 이 파일에서 모든 파일의 변화를 감지하고 호출합니다.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const sheet = e.range.getSheet();

  // 1. Command Sheet 감지 (A열) -> myFunction.gs의 myFunction() 호출
  if (typeof commandColumn !== 'undefined' && col === commandColumn) {
     if (typeof myFunction === 'function') myFunction();
  }

  // 2. 날짜 계산기 시트 감지
  const sName = sheet.getName();
  if (sName === "날짜 계산기" || sName === "날짜계산기") {
    
    // ODD 및 스프레드 계산 트리거 (B열, F열 등)
    if ((col === 2 && row >= 3) || (col === 6 && (row === 4 || row === 8 || row === 14 || row === 18))) {
      if (row <= 25 || col === 6) {
        if (typeof odd_runODDSetupFinal === 'function') odd_runODDSetupFinal(true);
      }
      if (row >= 26) {
        if (typeof odd_runSpreadStartingCalculation === 'function') odd_runSpreadStartingCalculation(true);
      }
      if (typeof odd_runConfirmationAutomation === 'function') odd_runConfirmationAutomation();
    }

    // 1-1. 스프레드 기준가(C26 또는 D26) 변경 시 계산
    if (row === 26 && (col === 3 || col === 4)) {
      if (typeof odd_runSpreadStartingCalculation === 'function') odd_runSpreadStartingCalculation(true);
    }

    // 1-2. 스프레드 증감 자동 계산 (스프레드 증감.gs 호출)
    if (typeof calculateSpreadChange === 'function') {
      calculateSpreadChange(e);
    }

    // 2-1. 컨포 실시간 감지 (P:T열)
    var lastCol = e.range.getLastColumn();
    if (lastCol >= 16 && col <= 20 && row >= 5) {
      if (typeof odd_runConfirmationAutomation === 'function') odd_runConfirmationAutomation();
    }
  }
}
