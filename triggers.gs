/**
 * 통합 트리거 & 자동화 도구
 * (ODD 자동화 + 스프레드 계산 + 컨포 자동생성 모두 통합)
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("자동화도구")
    .addItem("ODD 자동 세팅", "odd_runODDSetupFinal")
    .addItem("스프레드 환율 계산", "odd_runSpreadCalculation")
    .addToUi();
}

/** 
 * [마스터 onEdit] 
 * 1. Command Sheet 감지 (Column A) -> myFunction()
 * 2. 날짜 계산기 시트 감지 -> odd_autoCalculation(e)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const col = e.range.getColumn();
  
  // 1. A열(commandColumn) 변경 시 myFunction() 호출 - CommandSheet 기능
  // (myFunction.gs에 정의된 commandColumn 변수 사용)
  if (typeof commandColumn !== 'undefined' && col === commandColumn) {
     if (typeof myFunction === 'function') myFunction();
  }

  // 2. 날짜 계산기 자동화
  odd_autoCalculation(e);
}

function odd_autoCalculation(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sName = sheet.getName();
  if (sName !== "날짜 계산기" && sName !== "날짜계산기") return;
  
  const col = e.range.getColumn();
  const row = e.range.getRow();
  
  // 1. 달력(B3:C22) 또는 ODD(F열) 변경 시 처리
  if ((col === 2 || col === 3) && row >= 3 && row <= 22) {
    if (typeof odd_runODDSetupFinal === 'function') odd_runODDSetupFinal(true);
  }
  
  // 1-1. ODD 및 스프레드 계산 트리거 (B열, F열 등)
  if ((col === 2 && row >= 3) || (col === 6 && (row === 4 || row === 8 || row === 14 || row === 18))) {
    if (row <= 25 || col === 6) {
      if (typeof odd_runODDSetupFinal === 'function') odd_runODDSetupFinal(true);
    }
    if (row >= 26) {
      if (typeof odd_runSpreadStartingCalculation === 'function') odd_runSpreadStartingCalculation(true);
    }
    if (typeof odd_runConfirmationAutomation === 'function') odd_runConfirmationAutomation();
  }

  // 1-2. 스프레드 기준가(C26 또는 D26) 변경 시 계산
  if (row === 26 && (col === 3 || col === 4)) {
    if (typeof odd_runSpreadStartingCalculation === 'function') odd_runSpreadStartingCalculation(true);
  }

  // 2. 컨포 자동 생성 처리 (Q17:U21 영역 감지)
  if (col >= 17 && col <= 21 && row >= 5) {
     if (typeof odd_runConfirmationAutomation === 'function') odd_runConfirmationAutomation();
  }

  // 3. Swap Point 수정 시 Daily 실시간 계산 (J열 10열)
  if (col === 10 && row >= 5 && row <= 16) {
    const pt = sheet.getRange(row, 10).getValue();
    const dToday = sheet.getRange(row, 13).getValue();
    const dTom = sheet.getRange(row, 14).getValue();
    
    if (pt !== "" && !isNaN(parseFloat(pt))) {
      if (dToday && !isNaN(dToday) && dToday != 0) sheet.getRange(row, 15).setValue(parseFloat(pt) / dToday);
      if (dTom && !isNaN(dTom) && dTom != 0) sheet.getRange(row, 16).setValue(parseFloat(pt) / dTom);
    }
    odd_runODDSetupFinal(true);
  }

  // 4. 26행/27행 자동 수식 (F/G열 기반)
  if (typeof calculateSpreadChange === 'function') {
    calculateSpreadChange(e);
  }

  // 5. 27행/30행 스왑 보정 (필요시 활성화)
  if ((row === 27 || row === 30) && (col >= 2 && col <= 8)) {
     // SpreadsheetApp.flush(); Utilities.sleep(1000); if (typeof calculateSwapCorrection === 'function') calculateSwapCorrection(sheet);
  }
}

// (onOpen, onEdit, odd_autoCalculation만 유지합니다. 나머지 로직은 각 전용 파일인 날짜 계산기.gs, 컨포 생성.gs 등에 들어있습니다.)
