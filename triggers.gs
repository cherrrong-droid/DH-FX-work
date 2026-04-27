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
  if (sheet.getName() !== "날짜 계산기") return;
  
  const col = e.range.getColumn();
  const row = e.range.getRow();
  
  // 1. B열/F열 처리 (ODD 및 스프레드 계산)
  const sName = sheet.getName();
  if (sName === "날짜 계산기" || sName === "날짜계산기") {
    // B열(테너) 또는 F열(Odd입력) 변경 시
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

    // 1-2. 스프레드 증감 자동 계산 (F26/G26/F27/G27)
    if (typeof calculateSpreadChange === 'function') {
      calculateSpreadChange(e);
    }
  }

  // 2. 컨포 자동 생성 처리 (Q17:U21)
  if (col >= 17 && col <= 21 && row >= 5) {
     if (typeof odd_runConfirmationAutomation === 'function') odd_runConfirmationAutomation();
  }

  // 3. Swap Point 수정 시 Daily 실시간 계산 (J열 수정 시)
  // J(10) / M(13) -> O(15) 
  // J(10) / N(14) -> P(16)
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

  // 4. 26행/27행 자동 수식 계산 (스프레드 증감.gs 호출)
  if (typeof calculateSpreadChange === 'function') {
    calculateSpreadChange(e);
  }

  // 5. 27행/30행 스왑 보정 (필요시 활성화)
  if ((row === 27 || row === 30) && (col >= 2 && col <= 8)) {
     // SpreadsheetApp.flush(); Utilities.sleep(1000); if (typeof calculateSwapCorrection === 'function') calculateSwapCorrection(sheet);
  }
}

// ===== ODD 자동 세팅 (Section 2 기반 통합) =====
function odd_runODDSetupFinal(isAutomatic) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("날짜 계산기");
  if (!sheet) return;
  try {
    const parseODD = (v) => {
      if (!v || !v.includes("~")) return null;
      // 공백 유무에 관계없이 날짜(M/D) 두 개를 정확히 추출
      const m = v.match(/(\d+\/\d+)\s*~\s*(\d+\/\d+)/);
      return m ? { s: m[1], e: m[2] } : null;
    };
    
    // F4 / F14 입력 시 하단 날짜 셀 자동 업데이트
    const d1 = parseODD(sheet.getRange("F4").getValue());
    if (d1) { sheet.getRange("F7").setValue(d1.s); sheet.getRange("G7").setValue(d1.e); }
    const d2 = parseODD(sheet.getRange("F14").getValue());
    if (d2) { sheet.getRange("F17").setValue(d2.s); sheet.getRange("G17").setValue(d2.e); }

    const bcVal = sheet.getRange("B4:B25").getDisplayValues().flat();
    let pMap = {};
    bcVal.forEach(v => {
      let l = (v + "").trim(); if (!l || l.toUpperCase().includes("CALENDAR")) return;
      const dM = l.match(/(\d+)\s*d/i), dts = l.match(/(\d+\/\d+)/g);
      const t = (l.split(/[^A-Z0-9\/]/i)[0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, "");
      if (t && dts && dts.length >= 2 && dM) pMap[t] = { s: dts[0], f: dts[1], d: parseInt(dM[1]) };
    });

    const mkt = sheet.getRange("I5:N25").getValues();
    let mList = [];
    for (let j = 0; j < mkt.length; j++) {
      const t = (mkt[j][0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, "");
      if (pMap[t]) {
        const curR = 5 + j;
        const dy = (mkt[j][1] !== "" && !isNaN(mkt[j][1])) ? mkt[j][1] / pMap[t].d : "";
        sheet.getRange(curR, 11, 1, 3).setValues([[pMap[t].s, pMap[t].f, pMap[t].d]]);
        if (dy !== "") { sheet.getRange(curR, 14).setValue(dy); mList.push({ t: t, p: mkt[j][1], dy: dy, f: pMap[t].f }); }
      }
    }
    const spot = pMap["S/N"] ? pMap["S/N"].s : "";
    if (spot) { sheet.getRange("F6").setValue(spot); sheet.getRange("F16").setValue(spot); }

    [["F7", "F8", "F10", "G10"], ["F17", "F18", "F20", "G20"]].forEach(g => {
      const d = sheet.getRange(g[0]).getDisplayValue(), r = parseFloat(sheet.getRange(g[1]).getDisplayValue());
      if (d && !isNaN(r) && spot) {
        const pt = odd_calcH(d, spot, mList);
        sheet.getRange(g[2]).setValue((r + (pt / 100)).toFixed(2));
        sheet.getRange(g[3]).setValue((r + (pt / 100)).toFixed(2));
      }
    });
    if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ ODD 완료!");
  } catch (e) { console.log(e.message); }
}

function odd_calcH(tS, sS, list) {
  const tD = odd_pDt(tS);
  const prio = ["1W", "1M", "2M", "3M", "6M", "9M", "1Y", "S/N"];
  list.sort((a, b) => prio.indexOf(a.t) - prio.indexOf(b.t));
  let ref = list[0];
  for (let d of list) { if (odd_pDt(d.f) >= tD) { ref = d; break; } }
  return ref.p + (((tD - odd_pDt(ref.f)) / 86400000) * ref.dy);
}

function odd_pDt(mmdd) {
  const p = mmdd.split("/");
  const d = new Date(); d.setMonth(parseInt(p[0]) - 1, parseInt(p[1]));
  d.setHours(0, 0, 0, 0); return d;
}

// ===== 스프레드 계산 (Section 3 기반 통합) =====
function odd_runSpreadCalculation(isAutomatic) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("날짜 계산기");
  if (!sheet) return;
  const mkt = sheet.getRange("I5:J25").getValues();
  let pts = {};
  mkt.forEach(r => { const t = (r[0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, ""); if (t) pts[t] = parseFloat(r[1]); });
  const base = parseFloat(sheet.getRange("C25").getValue());
  if (isNaN(base)) return;
  const ins = sheet.getRange("B26:B50").getDisplayValues().flat(), curs = sheet.getRange("C26:C50").getValues();
  let res = [];
  ins.forEach((v, i) => {
    let l = (v + "").trim(); if (!l || !l.includes("*")) { res.push([curs[i][0]]); return; }
    const fPart = l.split("*")[0].trim().toUpperCase();
    let tgt = fPart === "S" ? "SN" : (isNaN(fPart) ? fPart : fPart + "M");
    if (!tgt.match(/[W|M|Y]/) && tgt !== "SN") tgt += "M";
    const p = pts[tgt.replace("/", "")];
    if (p !== undefined) res.push([(base - Math.abs(p / 100)).toFixed(2)]);
    else res.push(["-"]);
  });
  sheet.getRange(26, 3, res.length, 1).setValues(res);
  if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ 스프레드 완료!");
}
