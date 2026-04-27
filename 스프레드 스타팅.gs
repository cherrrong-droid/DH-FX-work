function odd_runSpreadStartingCalculation(isAutomatic) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("날짜 계산기") || ss.getSheetByName("날짜계산기");
  if (!sheet) return;
  SpreadsheetApp.flush();

  const mkt = sheet.getRange("I5:J25").getValues();
  let pts = {};
  mkt.forEach(r => { 
    const t = (r[0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, ""); 
    if (t) pts[t] = parseFloat(r[1]); 
  });

  let base = parseFloat(sheet.getRange("C26").getValue());
  if (isNaN(base)) base = parseFloat(sheet.getRange("D26").getValue());
  if (isNaN(base)) return;

  const ins = sheet.getRange("B27:B50").getDisplayValues().flat();
  const cursC = sheet.getRange("C27:C50").getValues();
  const cursD = sheet.getRange("D27:D50").getValues();
  let resC = [], resD = [];

  ins.forEach((v, i) => {
    let l = (v + "").trim(); 
    if (!l || !l.includes("*")) { 
      resC.push([cursC[i][0]]); 
      resD.push([cursD[i][0]]);
      return; 
    }
    
    // 첫번째 테너 (Column C 용)
    const fPart = l.split("*")[0].trim().toUpperCase();
    let tgtC = fPart === "S" ? "SN" : (isNaN(fPart) ? fPart : fPart + "M");
    if (!tgtC.match(/[W|M|Y]/) && tgtC !== "SN") tgtC += "M";
    const pC = pts[tgtC.replace("/", "")];
    if (pC !== undefined) resC.push([(base - Math.abs(pC / 100)).toFixed(2)]);
    else resC.push(["-"]);

    // 두번째 테너 (Column D 용)
    const sPart = l.split("*")[1].trim().toUpperCase();
    let tgtD = sPart === "S" ? "SN" : (isNaN(sPart) ? sPart : sPart + "M");
    if (!tgtD.match(/[W|M|Y]/) && tgtD !== "SN") tgtD += "M";
    const pD = pts[tgtD.replace("/", "")];
    if (pD !== undefined) resD.push([(base - Math.abs(pD / 100)).toFixed(2)]);
    else resD.push(["-"]);
  });

  sheet.getRange(27, 3, resC.length, 1).setValues(resC);
  sheet.getRange(27, 4, resD.length, 1).setValues(resD); // D열 업데이트
  if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ 스프레드 계산 완료!");
}
