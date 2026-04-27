function calculateSpreadChange(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sName = sheet.getName();
  if (sName !== "날짜 계산기" && sName !== "날짜계산기") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col !== 6 && col !== 7) return;

  if (row === 26) {
    const fVal = parseFloat(sheet.getRange("F26").getValue()) || 0;
    const gVal = parseFloat(sheet.getRange("G26").getValue()) || 0;
    sheet.getRange("H26").setValue(fVal + gVal);
  }
  
  if (row === 27) {
    const fVal = parseFloat(sheet.getRange("F27").getValue()) || 0;
    const gVal = parseFloat(sheet.getRange("G27").getValue()) || 0;
    sheet.getRange("H27").setValue(fVal - gVal);
  }
}
