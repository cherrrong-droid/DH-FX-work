/**
 * 컨펌 생성 자동화 스크립트 (Reverted Parsing with Merge Fix)
 */

function odd_runConfirmationAutomation() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("날짜 계산기") || SpreadsheetApp.getActive().getSheetByName("날짜계산기");
  if (!sheet) return;
  const pDict = odd_parseC(sheet);
  
  for (let startRow = 5; startRow <= 205; startRow += 8) {
    const vals = sheet.getRange(startRow, 16, 2, 5).getValues().flat();
    const input = vals.find(v => String(v).includes("<") || String(v).includes(">"));
    if (!input) continue;

    const resObj = odd_gMC(input.toString().trim(), pDict);
    
    // Row 1 (상단 컨펌: P7, U7 위치)
    resObj.row1.forEach((c, idx) => {
      if (idx < 2) {
        const target = sheet.getRange(startRow + 2, 16 + (idx * 5), 2, 5);
        target.merge().setValue(c).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true).setFontSize(9);
      }
    });

    // Row 2 (하단 컨펌: P9, U9 위치)
    resObj.row2.forEach((c, idx) => {
      if (idx < 2) {
        const target = sheet.getRange(startRow + 4, 16 + (idx * 5), 2, 5);
        target.merge().setValue(c).setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true).setFontSize(9);
      }
    });
  }
}

function odd_gMC(cmd, prdDict) {
  let row1 = [], row2 = [];
  try {
    // 5. 7/27 ~ 7/28... 처럼 앞에 '번호. '이 있으면 제거
    let cleanCmd = cmd.trim();
    if (cleanCmd.match(/^\d+\.\s+/)) {
      cleanCmd = cleanCmd.replace(/^\d+\.\s+/, "");
    }

    let rawPts = cleanCmd.split(/\s+/).filter(p => p.trim() !== "");
    
    // ODD DATE 패턴 감지 (예: 7/27 ~ 7/28 [1d])
    let oddMatch = cleanCmd.match(/^(\d+\/\d+)\s*~\s*(\d+\/\d+)\s*\[(\d+)[dD]\]/);
    let tenorAbbr, dateInfo, pts;

    if (oddMatch) {
      // ODD DATE 케이스
      let sDate = oddMatch[1], eDate = oddMatch[2], days = oddMatch[3];
      tenorAbbr = sDate + "~" + eDate + "[" + days + "D]";
      
      const mD = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
      let s = sDate.split("/"), f = eDate.split("/");
      let sM = parseInt(s[0]), fM = parseInt(f[0]), sD = parseInt(s[1]), fD = parseInt(f[1]);
      let sY = 2026;
      let fY = (fM < sM || (fM === sM && fD <= sD)) ? 2027 : 2026;
      dateInfo = "VAL " + s[1].padStart(2,'0') + " " + mD[sM-1] + " " + sY + " AND " + f[1].padStart(2,'0') + " " + mD[fM-1] + " " + fY;
      
      // 날짜 부분을 제외한 나머지 토큰들 추출
      let remainingCmd = cleanCmd.substring(oddMatch[0].length).trim();
      pts = remainingCmd.split(/\s+/).filter(p => p.trim() !== "");
    } else {
      // 일반 테너 케이스
      pts = rawPts;
      tenorAbbr = odd_getTenorAbbr(pts[0]);
      dateInfo = odd_rP(pts[0], prdDict);
    }
    
    let obj = odd_ex(pts, !!oddMatch); 
    if (obj.companyList.length < 2) return { row1: [], row2: [] };
    
    const dict = odd_getB();
    const first = obj.companyList[0], last = obj.companyList[obj.companyList.length - 1];
    
    if (obj.switchComp) {
      let t1 = odd_mF(first, obj.switchComp, obj.calculationList, dict);
      let t2 = odd_mF(obj.switchComp, last, obj.calculationList, dict); 
      row1.push(tenorAbbr + " " + t1[0] + "\n" + dateInfo);
      row1.push(tenorAbbr + " " + t2[1] + "\n" + dateInfo);
      row2.push(tenorAbbr + " " + t2[0] + "\n" + dateInfo);
      row2.push(tenorAbbr + " " + t1[1] + "\n" + dateInfo);
    } else {
      let mainRs = odd_mF(first, last, obj.calculationList, dict);
      row1.push(tenorAbbr + " " + mainRs[0] + "\n" + dateInfo);
      row1.push(tenorAbbr + " " + mainRs[1] + "\n" + dateInfo);
    }
  } catch (e) { console.log("Error: " + e); }
  return { row1: row1, row2: row2 };
}

function odd_mF(l, r, calc, dict) {
  const find = (n) => { 
    if (!n) return "?";
    const cleanN = String(n).trim();
    if (dict[cleanN]) return dict[cleanN];
    for(let k in dict) { if (cleanN.startsWith(k)) return dict[k]; } 
    return cleanN; 
  };
  
  let pnt = parseFloat(calc[0] || "0");
  let am = calc[1] || "-";
  let lR = parseFloat(calc[2]); 
  
  let ln = find(l), rn = find(r);
  let nearR = isNaN(lR) ? "NaN" : lR.toFixed(2);
  let farVal = isNaN(lR) ? 0 : lR + (pnt / 100);
  let farR = isNaN(lR) ? "NaN" : (Number.isInteger(farVal * 100) ? farVal.toFixed(2) : farVal.toFixed(3));
  
  return [
    ln + " S/B USD " + am + " MIO WITH " + rn + " AT " + nearR + " & " + farR + " (@ " + pnt + ")",
    rn + " B/S USD " + am + " MIO WITH " + ln + " AT " + nearR + " & " + farR + " (@ " + pnt + ")"
  ];
}

function odd_ex(p, isOdd) {
  let o = { companyList: [], calculationList: [], switchComp: null };
  const TENORS = ["오버", "탐", "스팟넥", "on", "tn", "sn", "1w", "1m", "2m", "3m", "6m", "9m", "1y", "1달", "2달", "3달", "6달", "1년", "2년"];
  let numbersStarted = false;
  
  for (let i = 0; i < p.length; i++) {
    const s = p[i];
    const sLow = s.toLowerCase();
    const cleanS = s.replace(/[()]/g, "");
    
    if (cleanS !== "" && !isNaN(parseFloat(cleanS)) && isFinite(cleanS)) {
      o.calculationList.push(cleanS);
      numbersStarted = true;
      continue;
    }
    
    if (numbersStarted) continue;
    
    // ODD DATE인 경우 첫 토큰이 테너가 아니므로 스킵 방지
    if (!isOdd && i === 0 && TENORS.some(t => sLow.includes(t))) continue;
    if (s === "<" || s === ">" || s === "-") continue;
    
    let isSw = s.includes("(") && s.includes(")");
    let cW = (s.match(/\(?([^()]+)\)?/) || [null, s])[1].trim();
    if (cW && isNaN(parseFloat(cW))) {
      o.companyList.push(cW);
      if (isSw) o.switchComp = cW;
    }
  }
  return o;
}

function odd_parseC(sh) {
  const b = sh.getRange("B4:B50").getDisplayValues().flat();
  const mD = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  let d = {};
  b.forEach(v => {
    let l = String(v).trim(); if (!l || l.toUpperCase().includes("CALENDAR")) return;
    const dt = l.match(/(\d+\/\d+)/g);
    const t = l.split(/[^A-Z0-9\/]/i)[0].toUpperCase().replace(/[^A-Z0-9\/]/g, "");
    if (t && dt && dt.length >= 2) {
      const s = dt[0].split("/"), f = dt[1].split("/");
      let sY = 2026;
      let sM = parseInt(s[0]), sD = parseInt(s[1]);
      let fM = parseInt(f[0]), fD = parseInt(f[1]);
      // 종료 월/일이 시작 월/일과 같거나 빠르면 연도 +1 (이월 처리)
      let fY = (fM < sM || (fM === sM && fD <= sD)) ? 2027 : 2026;
      
      const val = "VAL " + s[1].padStart(2,'0') + " " + mD[sM-1] + " " + sY + " AND " + f[1].padStart(2,'0') + " " + mD[fM-1] + " " + fY;
      d[t.toLowerCase()] = val;
      if (t === "O/N") d["on"] = val; if (t === "T/N") d["tn"] = val; if (t === "S/N") d["sn"] = val;
    }
  });
  return d;
}

function odd_getTenorAbbr(k) {
  const t = k.toLowerCase();
  if (t === "오버" || t === "on") return "ON"; 
  if (t === "탐" || t === "tn") return "TN"; 
  if (t === "스팟넥" || t === "sn") return "SN";
  
  // 3*6 또는 3x6 케이스
  if (t.includes("*") || t.includes("x")) {
    return t.replace("*", "X").replace("x", "X").toUpperCase();
  }

  let n = k.replace(/[^0-9]/g, "");
  if (k.includes("주")) return n + "W"; if (k.includes("달")) return n + "M"; if (k.includes("년")) return n + "Y";
  return k.toUpperCase();
}

function odd_rP(k, dDict) {
  let key = k.replace("주", "w").replace("달", "m").replace("년", "y").toLowerCase();
  if (k === "탐") key = "tn"; if (k === "오버") key = "on"; if (k === "스팟넥") key = "sn";
  
  // 3*6 케이스 처리
  if (key.includes("*") || key.includes("x")) {
    let pts = key.split(/[*x]/);
    let left = pts[0], right = pts[1];
    
    // 숫자로만 되어 있으면 'm' 붙여줌
    if (!isNaN(left)) left += "m";
    if (!isNaN(right)) right += "m";
    
    let dL = dDict[left], dR = dDict[right];
    if (dL && dR) {
      let pL = dL.split(" AND "), pR = dR.split(" AND ");
      if (pL.length >= 2 && pR.length >= 2) {
        return "VAL " + pL[1].trim() + " AND " + pR[1].trim();
      }
    }
  }

  return dDict[key] || "VAL [DATE] AND [DATE]";
}

function odd_getB() {
  return {
    "경남":"Kyongnam(SL)","공상":"ICBC(SL)","공상원":"ICBC(SL)","국민종":"Kookmin(SL)","국민경":"Kookmin(SL)",
    "기종":"IBK(SL)","기원":"IBK(SL)","기철":"IBK(SL)","노무라윤":"Nomura(SL)","노무라민":"Nomura(SL)",
    "농철":"Nonghyup(SL)","농이":"Nonghyup(SL)","농완":"Nonghyup(SL)","농화":"Nonghyup(SL)","농욱":"Nonghyup(SL)",
    "뉴욕박":"BNYM(SL)","도수":"Deutsche(SL)","도전":"Deutsche(SL)","아이엠 ":"IMB_DAEGU(SL)",
    "대신패시브":"Daishin Sec(SL)","딥조":"DBS(SL)","딥박":"DBS(SL)","동송":"MUFG(SL)","동김":"MUFG(SL)","동준":"MUFG(SL)",
    "미규":"Mizuho(SL)","미조":"Mizuho(SL)","미영":"Mizuho(SL)","미래신":"MIRAE Sec(SL)","미래노":"MIRAE Sec(SL)","미래창":"MIRAE Sec(SL)",
    "메리구":"Meritz Sec(SL)","모건":"M.Stanley(SL)","벤식":"BNP PARIBAS(SL)","벤홍":"BNP PARIBAS(SL)","벤하":"BNP PARIBAS(SL)",
    "보아종":"BOA(SL)","보아현":"BOA(SL)","보아원":"BOA(SL)","부산철":"Busan(SL)","부산복":"Busan(SL)","부산연":"Busan(SL)",
    "산업고":"KDB(SL)","산업김":"KDB(SL)","삼증원":"Samsung Sec(SL)","삼증박":"Samsung Sec(SL)",
    "속규":"Societe(SL)","속백":"Societe(SL)","속슬":"Societe(SL)","수협":"Suhyup(SL)","스미토모":"Sumitomo(SL)",
    "신영기존":"Shinyoung Sec(SL)","신투준":"Shinhan Sec(SL)","신투란":"Shinhan Sec(SL)","신투현":"Shinhan Sec(SL)","신투최":"Shinhan Sec(SL)",
    "신서":"Shinhan(SL)","신열":"Shinhan(SL)","스탠섭":"SCB(SL)","스탠철":"SCB(SL)","스탠구":"SCB(SL)",
    "스테잇왕":"SSBT(SL)","스테잇문":"SSBT(SL)","씨훈":"CITI(SL)","씨정":"CITI(SL)","우리태":"Woori(SL)","우리용":"Woori(SL)",
    "유오석":"UOB(SL)","유오백":"UOB(SL)","잉박":"ING(SL)","잉준":"ING(SL)","잉유":"ING(SL)","체노":"JP M.Chase(SL)","체정":"JP M.Chase(SL)",
    "크민":"CA-CIB(SL)","크송":"CA-CIB(SL)","크류":"CA-CIB(SL)","하나원":"Hana(SL)","하나전":"Hana(SL)","하남궁":"Hana(SL)","하권":"Hana(SL)",
    "하증":"Hana Sec(SL)","한":"Korea Sec(SL)","한투":"Korea Sec(SL)","한투홍":"Korea Sec(SL)","한투식":"Korea Sec(SL)",
    "한화홍":"Hanwha Sec(SL)","한화":"Hanwha Sec(SL)","홍샹박":"HSBC(SL)","홍샹배":"HSBC(SL)","kb증권":"KB Sec(SL)",
    "증금조":"KSFC(SL)","현대차":"HMC Sec(SL)"
  };
}
