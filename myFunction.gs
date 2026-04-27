/**
 * myfunction (이코드는 command sheet가 실행되기 위한 코드야. 이건 절대 건드리지말고, 참고할때 써)
 * (이 파일의 onEdit은 triggers.gs에서 통합 관리되므로 주석 처리하거나 무시됩니다)
 */

/*
// 트리거 조건 설정
function onEdit(e) {

  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var newValue = range.getValue();

  // A열의 변경을 감지 (1번 열)
  if (range.getColumn() == commandColumn) {
    myFunction();
  }
}
*/

/**
 * 변수 세팅
 */
var cd = "CommandSheet" // '10달 스탠섭 > 산업강 -255 50 1380.95' 와 같은 변환을 해야하는 시트 이름
var ps = "PeriodSheet" // 날짜 기간에 대한 정보를 담은 시트 이름
var curYear = 2026; // 기준년도
var nextYear = curYear + 1;
var commandColumn = 1; // A열
var createdCommandColumnStart = 2; // B열, 여기 열부터 회사 수에 따라 S/B, B/S를 채워줌
var errorColor = '#FFCCCC'; // 연한 빨간색
var oSlashNColor = '#D3D3D3'; // 연한 회색
var date = '';

var yearDict = {};
var periodONTNSNDict = {};
var periodDict = {};
var speicalPeriodDict = {};
var monthDict = {};
var companyDict = {};
var colorRowList = [];

/**
 * 회사 약어 세팅
 */
companyDict["경남"] = "Kyongnam(SL)";
companyDict["공상"] = "ICBC(SL)";
companyDict["국민"] = "Kookmin(SL)";
companyDict["기"] = "IBK(SL)";
companyDict["노무라"] = "Nomura(SL)";
companyDict["놈"] = "Nomura(SL)";
companyDict["농"] = "Nonghyup(SL)";
companyDict["뉴욕"] = "BNYM(SL)";
companyDict["도"] = "Deutsche(SL)";
companyDict["대"] = "IMB_DAEGU(SL)(SL)";
companyDict["대신패시"] = "Daishin Sec(SL)";
companyDict["딥"] = "DBS(SL)";
companyDict["동경"] = "MUFG(SL)";
companyDict["미"] = "Mizuho(SL)";
companyDict["미래"] = "MIRAE Sec(SL)";
companyDict["메리최성"] = "Meritz Sec(SL)";
companyDict["모"] = "M.Stanley(SL)";
companyDict["벤"] = "BNP PARIBAS(SL)";
companyDict["보아"] = "BOA(SL)";
companyDict["부산"] = "Busan(SL)";
companyDict["산업"] = "KDB(SL)";
companyDict["삼증"] = "Samsung Sec(SL)";
companyDict["속"] = "Societe(SL)";
companyDict["수출입국"] = "EXIM(SL)";
companyDict["수"] = "Suhyup(SL)";
companyDict["스미토"] = "Sumitomo(SL)";
companyDict["신영기"] = "Shinyoung Sec(SL)";
companyDict["신"] = "Shinhan(SL)";
companyDict["신투"] = "Shinhan Sec(SL)";
companyDict["스탠"] = "SCB(SL)";
companyDict["스테잇"] = "SSBT(SL)";
companyDict["씨"] = "CITI(SL)";
companyDict["아이"] = "IMB_DAEGU(SL)";
companyDict["아이엠머"] = "IMB_DAEGU(SL)";
companyDict["안쯔"] = "ANZ(SL)";
companyDict["엔투"] = "NH Sec(SL)";
companyDict["우리"] = "Woori(SL)";
companyDict["유오"] = "UOB(SL)";
companyDict["유안타fic"] = "Yuanta Sec(SL)";
companyDict["잉"] = "ING(SL)";
companyDict["체"] = "JP M.Chase(SL)";
companyDict["크"] = "CA-CIB(SL)";
companyDict["키움증"] = "Kiwoom Sec(SL)";
companyDict["하나"] = "Hana(SL)";
companyDict["하남"] = "Hana(SL)";
companyDict["하증"] = "Hana Sec(SL)";
companyDict["한"] = "Korea Sec(SL)";
companyDict["한화"] = "Hanwha Sec(SL)";
companyDict["홍샹"] = "HSBC(SL)";

companyDict["abo"] = "ABOC(SL)";
companyDict["bo"] = "BOC(SL)";
companyDict["cc"] = "CCB(SL)";
companyDict["DB투자증"] = "DB Sec(SL)";
companyDict["ibk투자증"] = "IBK투자증권";
companyDict["icb"] = "ICBC(SL)";
companyDict["kb증"] = "KB Sec(SL)";
companyDict["ocb"] = "OCBC(SL)";

companyDict["스테잇진런"] = "SSBT(LDN)";
companyDict["안"] = "ANZ(SL)";
companyDict["엔투공"] = "NH Sec(SL)";
companyDict["메리김채"] = "Meritz Sec(SL)";
companyDict["메리윤재"] = "Meritz Sec(SL)";
companyDict["한"] = "Korea Sec(SL)";
companyDict["하나머"] = "Hana(SL)";
companyDict["ccb"] = "CCB(SL)";
companyDict["메리정광"] = "Meritz Sec(SL)";
companyDict["우리머"] = "Woori(SL)";
companyDict["유안타"] = "Yuanta Sec(SL)";
companyDict["메리"] = "Meritz Sec(SL)";
companyDict["메리sr"] = "Meritz Sec(SL)";
companyDict["메리SR"] = "Meritz Sec(SL)";
companyDict["증금"] = "KSFC(SL)";
companyDict["교보증"] = "교보증권";
companyDict["현대"] = "HMC Sec(SL)";
companyDict["메리구성"] = "Meritz Sec(SL)";
companyDict["수"] = "Suhyup(SL)";

/**
 * 달 약어 세팅
 */
monthDict["01"] = "JAN";
monthDict["02"] = "FEB";
monthDict["03"] = "MAR";
monthDict["04"] = "APR";
monthDict["05"] = "MAY";
monthDict["06"] = "JUN";
monthDict["07"] = "JUL";
monthDict["08"] = "AUG";
monthDict["09"] = "SEP";
monthDict["10"] = "OCT";
monthDict["11"] = "NOV";
monthDict["12"] = "DEC";

/**
 * 실행될 코드 함수
 */
function myFunction() {
  convertPeriodSheetToDict();
  console.log(periodDict);
  processCommands();
}

/**
 * PeriodSheet에 있는 내용에 대해 Dict에 세팅 함수
 */
function convertPeriodSheetToDict() {

  var periodSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ps);
  var periodData = periodSheet.getDataRange().getValues();
  var onTnSn = {};

  // PeriodSheet 에서 VAL 형식의 문자열 생성
  for (var i = 0; i < periodData.length; i++) {
    var parts = periodData[i][0].split(" ");
    var periodKey = parts[0];
    var periodValue = periodData[i][0].substring(periodKey.length + 1);
        
    if (periodValue.toLowerCase() === "new york holiday") {
      periodDict[periodKey] = periodValue;
    } else {
      // pass by value 한계 극복을 위한 객체 사용
      let context = {
        parts: [],
        diffDay: 0,
        delta: 0
      };
      var parts = [];
      var diffDay = 0;
      var endYear = 0;

      if (validateDate(periodValue, context)) {
        parts = context.parts;
        diffDay = context.diffDay;

        if (parts.length < 2) {
          continue;
        }

        // 처음 O/N, T/N, S/N에 대하서는 undefined 값이지만 아래에서 다시 돌게끔 i = i - 1을 함
        var cy = yearDict[parts[0]];
        
        var endPeriod = addDaysToDate(cy + '/' + parts[0], diffDay);
        if (equals(endPeriod[1], parts[1])) {
          endYear = endPeriod[0];
        } else {
          // 잘못된 diffDay 지만 범위에 맞게 다시 생성
          endYear = createEndYear(cy, parts, periodKey)
        }

        // 시작 날짜
        var startDate = parts[0];
        var startParts = startDate.split("/").map((x, index) => {
          return index % 2 === 0 ? x.padStart(2, '0') : x;
        });
        if (startParts.length < 2) {
          continue;
        }
        var startMonth = startParts[0];
        var startDay = startParts[1];
        
        // 종료 날짜
        var endDate = parts[1];
        var endParts = endDate.split("/").map((x, index) => {
          return index % 2 === 0 ? x.padStart(2, '0') : x;
        });
        if (endParts.length < 2) {
          continue;
        }
        var endMonth = endParts[0];
        var endDay = endParts[1];

      } else {
        // cannot match date range
        continue;
      }
                  
      // Construct the formatted period and Add to period dictionary
      periodDict[periodKey] = "VAL " + startDay + " " + monthDict[startMonth] + " " + cy + " AND " + endDay + " " + monthDict[endMonth] + " " + endYear;

      if (periodKey === 'O/N') { // TODAY
        onTnSn['0'] = parseInt(startMonth);
        if (yearDict[parts[0]] === undefined) {
          i = i - 1;
          yearDict[parts[0]] = curYear;
        }
        yearDict[parts[0]] = curYear;
        yearDict['0'] = curYear;
        speicalPeriodDict['tod'] = startDay + " " + monthDict[startMonth] + " ";
      } else if (periodKey === 'T/N') { // 다음 업무일
        onTnSn['1'] = parseInt(startMonth);
        if (yearDict[parts[0]] === undefined) {
          i = i - 1;
          if (onTnSn['0'] > onTnSn['1']) {
            yearDict[parts[0]] = nextYear;
            yearDict['1']= nextYear;
          } else {
            yearDict[parts[0]] = curYear;
            yearDict['1']= curYear;
          }
        }
        speicalPeriodDict['tom'] = startDay + " " + monthDict[startMonth] + " ";
      } else if (periodKey === 'S/N') { // 다다음 업무일
        onTnSn['2'] = parseInt(startMonth);
        if (yearDict[parts[0]] === undefined) {
          i = i - 1;
          if (onTnSn['0'] > onTnSn['2']) {
            yearDict[parts[0]] = nextYear;
            yearDict['2']= nextYear;
          } else {
            yearDict[parts[0]] = curYear;
            yearDict['2']= curYear;
          }
        }
        speicalPeriodDict['spot'] = startDay + " " + monthDict[startMonth] + " ";
      }
    }
  }
  delete  periodDict[''];

  // tod, spot에 대한 날짜 변환 결과 세팅
  speicalPeriodDict['tod'] = speicalPeriodDict['tod'] + yearDict['0'];
  speicalPeriodDict['tom'] = speicalPeriodDict['tom'] + yearDict['1'];
  speicalPeriodDict['spot'] = speicalPeriodDict['spot'] + yearDict['2']; 
  console.log('speicalPeriodDict', speicalPeriodDict);
}

/**
 * CommandSheet에 있는 내용에 대해 Convert 하는 함수
 */
function processCommands() {
  
  var commandSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cd);
  var commandData = commandSheet.getRange(1, commandColumn, commandSheet.getLastRow(), 1).getValues();
  var maxCompanyCount = 0;

  // CommandSheet에서 행별로 실행
  for (let i = 1; i < commandData.length; i++) {

    try {
      let updatedCommand = commandData[i][0].trim();
      if (updatedCommand === "") continue;
      let parts = updatedCommand.split(" ");
      
      parts = removeFirstNumberIfExist(parts);
      if (parts.length < 7) {
        colorRowList.push([i, errorColor]);
        continue;
      }
      
      let object = extractCompanyListAndCalculationListInParts(parts);
      let companyList = object.companyList;
      let calculationList = object.calculationList;
      let startCompanyIndex = object.startCompanyIndexInParts;

      // 회사는 최소 2개 이상이어야함
      if (companyList.length < 2) {
        colorRowList.push([i, errorColor]);
        continue;
      }
      
      // 아래에서 컬럼 간격 조절을 위해 최대치 계산
      maxCompanyCount = Math.max(maxCompanyCount, companyList.length);

      let commandPeriod = makeSecond(parts, i, startCompanyIndex);
      
      // S/B, B/S 결과 출력
      for (let j = 0; j < companyList.length - 1; j++) {
        
        // S/B와 B/S 결과 생성
        let commandSellNBuyNSell = makeFirst(companyList[j], companyList[j + 1], calculationList);

        // 오버, 탐, 스팟넥에서만 소수 첫번째 자리까지 값을 사용할 수 있음 (ex. -6.5)
        if (isSingleDecimalPlace(calculationList[0]) && parts[0] != '오버' && parts[0] !== '탐' && parts[0] !== '스팟넥') {
          colorRowList.push([i, errorColor]);
        }

        // 날짜 정보와 합치기
        let sbResult = commandSellNBuyNSell[0] + '\n' + commandPeriod;
        let bsResult = commandSellNBuyNSell[1] + '\n' + commandPeriod;

        let resultColumn = createdCommandColumnStart + 2 * j;
        commandSheet.getRange(i + 1, resultColumn).setValue(sbResult);
        commandSheet.getRange(i + 1, resultColumn + 1).setValue(bsResult);

        // 결과값의 정상적인 값인지 확인
        if (!isValidResult(sbResult, bsResult)) {
          colorRowList.push([i, errorColor]);
          continue;
        }
      }
    } catch (err) {
      colorRowList.push([i, errorColor]);
    }
  }

  // 셀의 ROW 사이즈 자동 지정
  // autoResizeColumnToFitContents(commandSheet, commandColumn, false, 7.5);
  commandSheet.setColumnWidth(commandColumn, 450);
  for (let k = createdCommandColumnStart; k <= createdCommandColumnStart + 2 * maxCompanyCount - 3; k++) {
    autoResizeColumnToFitContents(commandSheet, k, true, 7.2);  
  }

  // 배경색 초기화 및 지정
  let range = commandSheet.getRange(2, commandColumn, commandSheet.getMaxRows(), 1);
  range.setBackground(null);

  // 잘못된 거래내역에 대해 색깔 표시
  console.log('colorRowList', colorRowList);
  colorRowList.forEach(function(value) {
      row = value[0];
      color = value[1];
      commandSheet.getRange(row + 1, commandColumn).setBackground(color);
  });
}

/**
 * 거래내역에서 불필요한 값에 대해 제거
 */
function removeFirstNumberIfExist(parts) {

  // '3. 9달 산업강 > 스탠김 -355 50 1380.90'  여기서 '3.' 이 부분에 대해 무시하기 위한 내용
  var numberMatches = parts[0].match(/^\d+\.$/);
  if (numberMatches) {
    return parts.slice(1);
  }
  return parts;
}

/**
 * parts안에서 회사 리스트와 계산 리스트 추출
 */
function extractCompanyListAndCalculationListInParts(parts) {

  let object = {
    companyList: [],
    calculationList: [],
    startCompanyIndexInParts: 0
  };
  let startCompanyToken = false;
  for (let i = 1; i < parts.length; i++) {
    // 숫자가 나오면 그 값 이후는 계산해야하는 값임
    if (isNumeric(parts[i])) {
      object.calculationList = parts.slice(i);
      break;
    }
    
    let cleanWord = extractCompany(parts[i]);
    // 회사는 알파벳과 한글만 포함되어야함
    if (isAlphaHangulOnly(cleanWord)) {
      object.companyList.push(cleanWord)
      if (!startCompanyToken) {
        startCompanyToken = true;
        object.startCompanyIndexInParts = i;
      }
    }
  }
  return object;
}

/**
 * 소숫점 이하 자릿수가 1인 경우 확인
 */
function isSingleDecimalPlace(numStr) {

  // 소수점 위치를 찾기
  let decimalIndex = numStr.indexOf('.');
  // 소수점이 없으면 false 반환 (정수인 경우)
  if (decimalIndex === -1) {
    return false;
  }
  // 소수점 이후 부분의 길이를 확인
  let decimalPlaces = numStr.length - decimalIndex - 1;
  // 소수점 이하 자릿수가 정확히 1인 경우 true 반환
  return decimalPlaces === 1;
}

/**
 * 결과물의 첫번째 line을 만드는 함수
 * ex. BNP PARIBAS(SL) S/B USD 100 MIO WITH Hana Sec(SL) AT 1389.50 & 1384.60 (@ -490)
 */
function makeFirst(companyLeft, companyRight, calculationList) {
  
  // Extract Without Family Name
  companyLeft = companyLeft.slice(0, -1);
  companyRight = companyRight.slice(0, -1);

  let diff = parseFloat(calculationList[0]) / 100; // -65, -8.5, ...
  let amount = calculationList[1]; // 10, 50, ... (MIO)
  var leftRate = parseFloat(calculationList[2]); // 1389.80, 1390.60, ...
  var rightRate = leftRate + diff;
  var formattedLeftRate = leftRate.toFixed(2);
  var formattedRightRate = rightRate.toFixed(Math.max(2, compareDecimalPlaces(diff, leftRate)));

  // Sell & Buy, Buy & Sell 결과 생성 (첫번째줄)
  var sb = date.toUpperCase() + " " + companyDict[companyLeft] + " S/B USD " + amount + " MIO WITH " + companyDict[companyRight] + " AT " + formattedLeftRate + " & " + formattedRightRate + " (@ " + calculationList[0] + ")";
  var bs = date.toUpperCase() + " " + companyDict[companyRight] + " B/S USD " + amount + " MIO WITH " + companyDict[companyLeft] + " AT " + formattedLeftRate + " & " + formattedRightRate + " (@ " + calculationList[0] + ")";

  return [sb, bs];
}

/**
 * 소숫점 기준으로 오른쪽 영역의 길이 중에 큰 값 return
 */
function compareDecimalPlaces(a, b) {

  let aString = a.toString();
  let indexA = aString.indexOf('.');
  let placeA = 0;
  if (indexA != -1) {
    placeA = aString.length - indexA - 1;
  } 
  var bString = b.toString();
  var indexB = bString.indexOf('.');
  let placeB = 0;
  if (indexB != -1) {
    placeB = bString.length - indexB - 1;
  } 
  return Math.max(placeA, placeB);
}

/**
 * 결과물의 두번째 line을 만들면서 엑셀의 예외 처리 함수
 * ex. VAL 26 JUN 2024 AND 26 AUG 2024
 */
function makeSecond(parts, i, startCompanyIndex) {
  
  let key = getPeriodKey(parts[0]); // 첫번째로 date 값을 여기서 함수 안에서 맵핑함
  if (key in periodDict) {
    if (periodDict[key].toLocaleLowerCase() === 'new york holiday') {
      colorRowList.push([i, oSlashNColor]);
    }
    return periodDict[key];
  } else {
    let newTypePeriod = createPeriodWords(parts.slice(0, startCompanyIndex));
    if (newTypePeriod === '' || !newTypePeriod.includes('VAL')) {
      colorRowList.push([i, errorColor]);
      return;
    } else {
      return newTypePeriod;
    }
  }
}

/**
 * 주어진 문자에 대해 회사만 추출
 * '(', ')' 문자는 제거 후 추출
 */
function extractCompany(company) {

  let match = company.match(/\(?([^()]+)\)?/);
  if (match) {
    return match[1];
  } else {
    return company;
  }
}

/**
 * 셀의 배경색 지정
 */
function setBackgroudColor(sheet, row, column, color) {

  if (sheet.getBackground() !== color) {
    sheet.getRange(row + 1, column).setBackground(color);
  }
}

/**
 * 10달, 9달, 3주, 오버, 스팟넥, ... 에 대한 KEY 값 생성
 */
function getPeriodKey(inputKey) {

  let periodType = inputKey.slice(-1);
  switch (periodType) {
    case "주":
      date = inputKey.slice(0, -1) + 'W';
      return inputKey.slice(0, -1) + "w";
    case "달":
      date = inputKey.slice(0, -1) + 'M';
      return inputKey.slice(0, -1) + "m";
    case "년":
      date = inputKey.slice(0, -1) + 'Y';
      return inputKey.slice(0, -1) + "y";
    default:
      if (inputKey === "탐") {
        date = 'TN';
        return "T/N";
      }
      else if (inputKey === "오버") {
        date = 'ON';
        return "O/N";
      } 
      else if (inputKey === "스팟넥") {
        date = 'SN'
        return "S/N";
      }
      return "";
  }
}

/**
 * 기간에 대한 검증 및 데이터 세팅
 */
function validateDate(period, context) {

  let result = true;
  const dateRangeRegex = /(\d{1,2}\/\d{1,2}~\d{1,2}\/\d{1,2})(\((\d+)[일|d]\)?[,]?([+-]?\d*)\))?/;
  let matches = period.match(dateRangeRegex);
  if (matches) {
    context.parts = matches[1].split('~');
    context.diffDay = parseInt(matches[3]);
    if (matches[4] === '') {
      context.delta = 0;
    } else {
      context.delta = parseInt(matches[4]);
    }
    result = true;
  } else {
    result = false;
  }
  return result;
}

/**
 * date는 24/6/30, 2024/6/30, 6/30 세가지 케이스에 해당
 */
function convertDateToSentece(date) {

  let date_ = date.split('/');
  if (date_.length === 3) { // 24/6/30, 2024/6/30
    date_[0] = date_[0].length == 2 ? "20"+ date_[0] : date_[0];
    date_[1] = date_[1].padStart(2, '0');
    return date_[2] + ' ' + monthDict[date_[1]] + ' ' + date_[0];
  } else if (date_.length === 2) { // 6/30
    date_[0] = date_[0].padStart(2, '0');
    return date_[1] + ' ' + monthDict[date_[0]] + ' ' + curYear;
  } else {
    return;
  }
}

/**
 * S/B, B/S 결과값에 대해 검증
 */
function isValidResult(sbResult, bsResult) {

  if (sbResult.includes('undefined') || bsResult.includes('undefined') || sbResult.includes('NaN') || bsResult.includes('NaN')) {
    return false;
  }
  return true;
}

/**
 * 컬럼 사이즈를 자동으로 조정
 */
function autoResizeColumnToFitContents(sheet, column, lineToken, fixelSize) {
  
  let data = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();
  let maxLength = 1;
  data.forEach(function(row) {
    let cellValue = row[0];
    if (lineToken) {
      // 줄 바꿈 문자를 기준으로 각 줄의 길이를 계산
      let lines = cellValue.split('\n');
      lines.forEach(function(line) {
        if (line.length > maxLength) {
          maxLength = line.length;
        }
      });
    } else {
      if (cellValue.length > maxLength) {
        maxLength = cellValue.length;
      }
    }
  });
  let newWidth = maxLength * fixelSize; // 대략적인 문자당 픽셀 너비를 설정합니다.
  sheet.setColumnWidth(column, newWidth);
}

/**
 * 날짜 검증을 위해 만든 함수로 dateString인 기준 날짜(ex. 2024/06/24)에서
 * dayToAdd 일 만큼 더해서 나온 값 리턴
 */
function addDaysToDate(dateString, daysToAdd) {

  // 주어진 날짜 문자열을 Date 객체로 변환
  const [year, month, day] = dateString.split('/').map(Number);
  const date = new Date(year, month - 1, day); // 월은 0부터 시작하므로 -1
  // 주어진 일수를 더함
  date.setDate(date.getDate() + daysToAdd);

  // 결과를 'YYYY/MM/DD' 형식의 문자열로 변환하여 반환
  const resultYear = date.getFullYear();
  const resultMonth = (date.getMonth() + 1).toString().padStart(2, '0'); // 월은 0부터 시작하므로 +1
  const resultDay = date.getDate().toString().padStart(2, '0');

  return [`${resultYear}`,`${resultMonth}/${resultDay}`];
}

/**
 * 뒤에 년도 결정
 */
function createEndYear(curYear, parts, periodKey) {

  let endYear = curYear;
  let startM = parseInt(parts[0].split('/')[0]);
  let startD = parseInt(parts[0].split('/')[1]);
  let endM = parseInt(parts[1].split('/')[0]);
  let endD = parseInt(parts[1].split('/')[1]);

  if (periodKey === '1y') {
    if (startM === 12 && endM === 1) {
      endYear = curYear + 2;
    } else { 
      endYear = curYear + 1;
    }
  } else {
    if (startM * 100 + startD > endM * 100 + endD) {
      endYear = curYear + 1;
    }
  }
  return endYear;
}

/**
 * 날짜 정보에 대해 동일한지 체크
 */
function equals(start, end) {

  let s = start.split('/');
  let e = end.split('/');

  if (parseInt(s[0]) * 100 + parseInt(s[1]) === parseInt(e[0]) * 100 + parseInt(e[1])) {
    return true;
  }
  return false;
}


/**
 * 날짜형식 체크
 */
function specialValidate(dateTime) {

  const dateTimeRegex = /(\d{2,4}\/)?(\d{1,2}\/\d{1,2})/;
  return  dateTime.match(dateTimeRegex);
}

/**
 * 'VAL 26 JUN 2024 AND 26 AUG 2024` 양식의 데이터 생성
 */
function createPeriodWords(specialCaseArray) {

  let specialPeriod = specialCaseArray.join('');
  if (specialPeriod.indexOf('(') != -1) {
    specialPeriod = specialPeriod.slice(0, specialPeriod.indexOf('('));
  }
  date = specialPeriod;
  if (specialPeriod.includes('~')) {
    let tempSpecialList = specialPeriod.split('~');
    let specialLeft = tempSpecialList[0];
    let specialRight = tempSpecialList[1];
    if (specialRight.indexOf('(') != -1) {
      specialRight = specialRight.slice(0, specialRight.indexOf('('));
    }
    let sl = '';
    let sr = '';
    let matchLeft = specialValidate(specialLeft)
    if (matchLeft) {
      sl = convertDateToSentece(matchLeft[0])
    } else {
      if (specialLeft in speicalPeriodDict) {
        sl = speicalPeriodDict[specialLeft];
      } else {
        return '';
      }
    }

    let matchRight = specialValidate(specialRight);
    if (matchRight) {
      sr = convertDateToSentece(matchRight[0])
    } else {
      if (specialRight in speicalPeriodDict) {
        sr = speicalPeriodDict[specialLeft];
      } else {
        return '';
      }
    }
    return 'VAL ' + sl + ' AND ' + sr;

  } else { // 1wx2w => left = 1w / right = 2w
    // 1wx2w => left = 1w / right = 2w
    if (specialPeriod.includes('x')) {
      let temp = specialPeriod.split('x');
      let leftBy = temp[0];
      let rightBy = temp[1];
      
      // 2x3 같은 케이스는 2mx3m과 동일
      if (isNumeric(leftBy) && isNumeric(rightBy)) {
        leftBy += 'm';
        rightBy += 'm';
      }

      // 기간 딕셔너리에 없으면 이상한거임
      if (!(leftBy in periodDict && rightBy in periodDict)) {
        return specialPeriod;
      }

      let tempLeft = periodDict[leftBy];
      let tempRight = periodDict[rightBy];
      tempLeft = tempLeft.split('AND');
      tempRight = tempRight.split('AND');
      return 'VAL ' + tempLeft[1].trim() + ' AND ' + tempRight[1].trim();
    }
  }
  return "";
}

/**
 * 숫자인지 판별
 */
function isNumeric(str) {

  if (typeof str != "string") return false; // 문자열 타입이 아닌 경우 false 반환
  return !isNaN(str) && // NaN이 아닌 경우
         !isNaN(parseFloat(str)) && // 숫자로 변환 가능한 경우
         isFinite(str); // 유한한 숫자인 경우
}

/**
 * 알파벳과 한글만 있는지 판별
 */
function isAlphaHangulOnly(str) {

  // 문자열이 비어 있으면 false 반환
  if (str.length === 0) return false;

  // 영문자와 한글 문자만 허용하는 정규표현식
  const alphaHangulRegex = /^[A-Za-z가-힣]+$/;
  return alphaHangulRegex.test(str);
}
