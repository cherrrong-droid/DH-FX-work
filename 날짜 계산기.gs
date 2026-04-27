function odd_runODDSetupFinal(isAutomatic) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("날짜 계산기") || ss.getSheetByName("날짜계산기");
  if (!sheet) return;
  try {
    const parseODD = (v) => {
      if (!v || !v.includes("~")) return null;
      const m = v.match(/(\d+\/\d+)\s*~\s*(\d+\/\d+)/);
      return m ? { s: m[1], e: m[2] } : null;
    };
    
    // F4 / F14 입력 시 하단 날짜 셀 자동 업데이트
    const d1 = parseODD(sheet.getRange("F4").getValue());
    if (d1) { sheet.getRange("F7").setValue(d1.s); sheet.getRange("G7").setValue(d1.e); }
    const d2 = parseODD(sheet.getRange("F14").getValue());
    if (d2) { sheet.getRange("F17").setValue(d2.s); sheet.getRange("G17").setValue(d2.e); }

    // [추가] B3:B22 및 C3:C22에서 날짜수([30d]) 추출하여 M5:M16, N5:N16에 입력
    const todayCal = sheet.getRange("B3:B22").getDisplayValues().flat();
    const tomCal = sheet.getRange("C3:C22").getDisplayValues().flat();
    
    const extractD = (str) => {
      const m = String(str).match(/\[(\d+)[dD]\]/);
      return m ? parseInt(m[1]) : "";
    };

    // M5:M16 (Today Day Count), N5:N16 (Tomorrow Day Count)
    let mVals = [], nVals = [];
    for (let i = 2; i <= 13; i++) { // B5 to B16 (B3=index 0)
      mVals.push([extractD(todayCal[i])]);
      nVals.push([extractD(tomCal[i])]);
    }
    sheet.getRange("M5:M16").setValues(mVals);
    sheet.getRange("N5:N16").setValues(nVals);

    const bcVal = sheet.getRange("B4:B25").getDisplayValues().flat();
    let pMap = {};
    bcVal.forEach(v => {
      let l = (v + "").trim(); if (!l || l.toUpperCase().includes("CALENDAR")) return;
      const dM = l.match(/(\d+)\s*d/i), dts = l.match(/(\d+\/\d+)/g);
      const t = (l.split(/[^A-Z0-9\/]/i)[0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, "");
      if (t && dts && dts.length >= 2 && dM) pMap[t] = { s: dts[0], f: dts[1], d: parseInt(dM[1]) };
    });

    const mkt = sheet.getRange("I5:O25").getValues();
    let mList = [];
    for (let j = 0; j < mkt.length; j++) {
      const curR = 5 + j;
      if (curR > 16) break; // [중요] 16행까지만 자동화 실행

      const t = (mkt[j][0] + "").toUpperCase().replace(/[^A-Z0-9\/]/g, "");
      if (pMap[t]) {
        const dy = (mkt[j][1] !== "" && !isNaN(mkt[j][1])) ? mkt[j][1] / pMap[t].d : "";
        sheet.getRange(curR, 11, 1, 3).setValues([[pMap[t].s, pMap[t].f, pMap[t].d]]);
        if (dy !== "") { sheet.getRange(curR, 15).setValue(dy); mList.push({ t: t, p: mkt[j][1], dy: dy, f: pMap[t].f }); }
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
    if (!isAutomatic) SpreadsheetApp.getUi().alert("✅ ODD 완료 (16행 제한 적용)!");
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
