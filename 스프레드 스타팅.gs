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
  const curs = sheet.getRange("C27:C50").getValues();
  let res = [];

  ins.forEach((v, i) => {
    let l = (v + "").trim(); 
    if (!l || !l.includes("*")) { 
      res.push([curs[i][0]]); 
      return; 
    }
    const fPart = l.split("*")[0].trim().toUpperCase();
    let tgt = fPart === "S" ? "SN" : (isNaN(fPart) ? fPart : fPart + "M");
    if (!tgt.match(/[W|M|Y]/) && tgt !== "SN") tgt += "M";
    
    const p = pts[tgt.replace("/", "")];
    if (p !== undefined) res.push([(base - Math.abs(p / 100)).toFixed(2)]);
    else res.push(["-"]);
  });

  sheet.getRange(27, 3, res.length, 1).setValues(res);
  if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ 스프레드 계산 완료!");
}
