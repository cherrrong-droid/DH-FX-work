function odd_runSpreadStartingCalculation(isAutomatic) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("날짜 계산기") || ss.getSheetByName("날짜계산기");
  if (!sheet) return;
  SpreadsheetApp.flush();

  // Market Data (I5:J25)
  const mkt = sheet.getRange("I5:J25").getValues();
  let pts = {};
  mkt.forEach(r => { 
    const t = (r[0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, ""); 
    if (t) pts[t] = parseFloat(r[1]); 
  });

  // Base Rate (C26 또는 D26 중 숫자가 있는 것 사용)
  let base = parseFloat(sheet.getRange("C26").getValue());
  if (isNaN(base)) base = parseFloat(sheet.getRange("D26").getValue());
  if (isNaN(base)) return;

  // Inputs (B27:B50) & Outputs (C27:C50)
  const ins = sheet.getRange("B27:B50").getDisplayValues().flat();
  const cursC = sheet.getRange("C27:C50").getValues();
  const cursD = sheet.getRange("D27:D50").getValues();
  let resC = [];
  let resD = [];

  ins.forEach((v, i) => {
    let l = (v + "").trim(); 
    if (!l || !l.includes("*")) { 
      resC.push([cursC[i][0]]); 
      resD.push([cursD[i][0]]);
      return; 
    }
    
    // 첫번째 테너 (P1)
    const fPart = l.split("*")[0].trim().toUpperCase();
    let tgtC = fPart === "S" ? "SN" : (isNaN(fPart) ? fPart : fPart + "M");
    if (!tgtC.match(/[W|M|Y]/) && tgtC !== "SN") tgtC += "M";
    const pC = pts[tgtC.replace("/", "")];
    
    // 두번째 테너 (P2)
    const sPart = l.split("*")[1].trim().toUpperCase();
    let tgtD = sPart === "S" ? "SN" : (isNaN(sPart) ? sPart : sPart + "M");
    if (!tgtD.match(/[W|M|Y]/) && tgtD !== "SN") tgtD += "M";
    const pD = pts[tgtD.replace("/", "")];

    // C열 계산: Base + (P1/100)
    if (pC !== undefined) resC.push([(base + (pC / 100)).toFixed(2)]);
    else resC.push(["-"]);

    // D열 계산: Base + (P2 - P1)/100 (사용자 요청: 뒤의 포인트 - 앞의 포인트)
    if (pC !== undefined && pD !== undefined) {
      const spreadDiff = pD - pC;
      resD.push([(base + (spreadDiff / 100)).toFixed(2)]);
    } else {
      resD.push(["-"]);
    }
  });

  sheet.getRange(27, 3, resC.length, 1).setValues(resC);
  sheet.getRange(27, 4, resD.length, 1).setValues(resD);
  if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ 스프레드 계산 완료!");
}
