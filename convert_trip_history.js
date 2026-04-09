const path = require('path');
const ExcelJS = require('exceljs');

function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === 'number') {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate());
  }

  if (typeof value === 'string') {
    const normalized = value.trim().replace(/\./g, '-').replace(/\//g, '-');
    const compact = normalized.replace(/-/g, '');
    if (/^\d{8}$/.test(compact)) {
      const y = Number(compact.slice(0, 4));
      const m = Number(compact.slice(4, 6));
      const d = Number(compact.slice(6, 8));
      const parsed = new Date(y, m - 1, d);
      if (!Number.isNaN(parsed.getTime())) {
        return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
      }
    }
    const d = new Date(normalized);
    if (!Number.isNaN(d.getTime())) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }

  return null;
}

function daysInclusive(start, end) {
  const ms = end.getTime() - start.getTime();
  return Math.floor(ms / (1000 * 60 * 60 * 24)) + 1;
}

function getCellValue(row, idx) {
  const cell = row.getCell(idx);
  return cell && cell.value != null ? cell.value : '';
}

function getCellText(row, idx) {
  const v = getCellValue(row, idx);
  if (typeof v === 'object' && v && v.text) return String(v.text).trim();
  return String(v ?? '').trim();
}

async function main() {
  const inputArg = process.argv[2] || '출장이력.xlsx';
  const outputArg = process.argv[3] || '';

  const inputPath = path.resolve(process.cwd(), inputArg);
  const outputPath = outputArg
    ? path.resolve(process.cwd(), outputArg)
    : path.join(path.dirname(inputPath), `${path.basename(inputPath, path.extname(inputPath))}_정산변환.xlsx`);

  const inWb = new ExcelJS.Workbook();
  await inWb.xlsx.readFile(inputPath);
  const inWs = inWb.worksheets[0];
  if (!inWs) {
    throw new Error('입력 파일의 첫 번째 시트를 찾지 못했습니다.');
  }

  const headerRow = inWs.getRow(1);
  const headerMap = new Map();
  headerRow.eachCell((cell, colNumber) => {
    const key = String(cell.text || cell.value || '').trim();
    if (key) headerMap.set(key, colNumber);
  });

  const required = ['시작일', '종료일', '출장지', '출장자', '공무용차량여부'];
  let hasAllHeaders = true;
  for (const h of required) {
    if (!headerMap.has(h)) {
      hasAllHeaders = false;
      break;
    }
  }

  // If headers are missing, fallback to the user-provided fixed column order.
  const idxStart = hasAllHeaders ? headerMap.get('시작일') : 1;
  const idxEnd = hasAllHeaders ? headerMap.get('종료일') : 2;
  const idxDest = hasAllHeaders ? headerMap.get('출장지') : 3;
  const idxTraveler = hasAllHeaders ? headerMap.get('출장자') : 4;
  const idxVehicle = hasAllHeaders ? headerMap.get('공무용차량여부') : 7;

  const rows = [];

  for (let r = 2; r <= inWs.rowCount; r += 1) {
    const row = inWs.getRow(r);
    const traveler = getCellText(row, idxTraveler);
    if (!traveler) continue;

    const startDate = parseDate(getCellValue(row, idxStart));
    let endDate = parseDate(getCellValue(row, idxEnd));

    if (!startDate) continue;
    if (!endDate) endDate = startDate;

    let s = startDate;
    let e = endDate;
    if (e < s) {
      const t = s;
      s = e;
      e = t;
    }

    const tripDays = Math.max(1, daysInclusive(s, e));
    const mealUnit = 25000;
    const isGovVehicle = getCellText(row, idxVehicle) === '이용';
    const dailyUnit = isGovVehicle ? 12500 : 25000;

    rows.push({
      traveler,
      tripDate: s,
      origin: '세종',
      destination: getCellText(row, idxDest),
      transport: isGovVehicle ? '관용차량' : '',
      transportFare: 0,
      tripDays,
      mealUnit,
      mealAmount: mealUnit * tripDays,
      dailyUnit,
      dailyAmount: dailyUnit * tripDays,
      lodgingNights: 0,
      lodgingCost: 0,
    });
  }

  rows.sort((a, b) => {
    if (a.traveler !== b.traveler) return a.traveler.localeCompare(b.traveler, 'ko');
    return a.tripDate - b.tripDate;
  });

  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet('출장정산');

  outWs.columns = [
    { header: '출장자', key: 'traveler', width: 14 },
    { header: '출장일', key: 'tripDate', width: 12 },
    { header: '출발지', key: 'origin', width: 10 },
    { header: '도착지', key: 'destination', width: 18 },
    { header: '교통수단', key: 'transport', width: 12 },
    { header: '교통요금', key: 'transportFare', width: 12 },
    { header: '출장일수', key: 'tripDays', width: 10 },
    { header: '식비단가', key: 'mealUnit', width: 12 },
    { header: '식비금액', key: 'mealAmount', width: 12 },
    { header: '일비단가', key: 'dailyUnit', width: 12 },
    { header: '일비금액', key: 'dailyAmount', width: 12 },
    { header: '숙박일수', key: 'lodgingNights', width: 12 },
    { header: '숙박비', key: 'lodgingCost', width: 12 },
    { header: '청구액', key: 'claimAmount', width: 12 },
    { header: '지급액', key: 'paidAmount', width: 12 },
  ];

  for (const item of rows) {
    const row = outWs.addRow({
      traveler: item.traveler,
      tripDate: item.tripDate,
      origin: item.origin,
      destination: item.destination,
      transport: item.transport,
      transportFare: item.transportFare,
      tripDays: item.tripDays,
      mealUnit: item.mealUnit,
      mealAmount: item.mealAmount,
      dailyUnit: item.dailyUnit,
      dailyAmount: item.dailyAmount,
      lodgingNights: item.lodgingNights,
      lodgingCost: item.lodgingCost,
      claimAmount: 0,
      paidAmount: 0,
    });

    const n = row.number;
    row.getCell(14).value = { formula: `F${n}+I${n}+K${n}+M${n}` };
    row.getCell(15).value = { formula: `N${n}` };
  }

  if (rows.length > 0) {
    // Insert traveler subtotal rows from bottom to top so row indices stay valid.
    const groups = [];
    let groupStart = 2;
    let currentTraveler = outWs.getCell('A2').value;

    for (let r = 3; r <= outWs.rowCount; r += 1) {
      const traveler = outWs.getCell(`A${r}`).value;
      if (traveler !== currentTraveler) {
        groups.push({ traveler: String(currentTraveler || ''), start: groupStart, end: r - 1 });
        currentTraveler = traveler;
        groupStart = r;
      }
    }
    groups.push({ traveler: String(currentTraveler || ''), start: groupStart, end: outWs.rowCount });

    for (let i = groups.length - 1; i >= 0; i -= 1) {
      const g = groups[i];
      const insertAt = g.end + 1;
      outWs.spliceRows(insertAt, 0, []);
      const subRow = outWs.getRow(insertAt);
      subRow.getCell(1).value = `${g.traveler} 소계`;
      subRow.getCell(14).value = { formula: `SUM(N${g.start}:N${g.end})` };
      subRow.getCell(15).value = { formula: `SUM(O${g.start}:O${g.end})` };
      subRow.font = { bold: true };
      subRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF7F7F7' },
      };
      subRow.getCell(14).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF7F7F7' },
      };
      subRow.getCell(15).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF7F7F7' },
      };
    }

    const subtotalRows = [];
    for (let r = 2; r <= outWs.rowCount; r += 1) {
      const label = String(outWs.getCell(`A${r}`).value || '');
      if (label.endsWith('소계')) {
        subtotalRows.push(r);
      }
    }

    const totalRow = outWs.addRow([]);
    const t = totalRow.number;
    totalRow.getCell(1).value = '전체 합계';
    if (subtotalRows.length === 1) {
      totalRow.getCell(14).value = { formula: `N${subtotalRows[0]}` };
      totalRow.getCell(15).value = { formula: `O${subtotalRows[0]}` };
    } else {
      const claimRefs = subtotalRows.map((r) => `N${r}`).join(',');
      const paidRefs = subtotalRows.map((r) => `O${r}`).join(',');
      totalRow.getCell(14).value = { formula: `SUM(${claimRefs})` };
      totalRow.getCell(15).value = { formula: `SUM(${paidRefs})` };
    }
    totalRow.font = { bold: true };
    totalRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE8EEF8' },
    };
    totalRow.getCell(14).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE8EEF8' },
    };
    totalRow.getCell(15).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE8EEF8' },
    };
  }

  outWs.getRow(1).font = { bold: true };
  outWs.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  outWs.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE8EEF8' },
  };

  for (let r = 2; r <= outWs.rowCount; r += 1) {
    const label = String(outWs.getCell(`A${r}`).value || '');
    const isSubtotal = label.endsWith('소계');
    const isTotal = label === '전체 합계';
    const isDetail = label.length > 0 && !isSubtotal && !isTotal;

    // After subtotal insertion, row indices shift. Re-assign detail formulas by current row number.
    if (isDetail) {
      outWs.getCell(`N${r}`).value = { formula: `F${r}+I${r}+K${r}+M${r}` };
      outWs.getCell(`O${r}`).value = { formula: `N${r}` };
      outWs.getCell(`B${r}`).numFmt = 'yyyy-mm-dd';
    }

    ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'].forEach((col) => {
      outWs.getCell(`${col}${r}`).numFmt = '#,##0';
    });
  }

  // Rebuild subtotal formulas by current sheet layout to avoid shifted references.
  for (let r = 2; r <= outWs.rowCount; r += 1) {
    const label = String(outWs.getCell(`A${r}`).value || '');
    if (!label.endsWith('소계')) continue;

    const end = r - 1;
    let start = end;

    while (start > 2) {
      const prevLabel = String(outWs.getCell(`A${start - 1}`).value || '');
      const prevIsBreak = prevLabel.endsWith('소계') || prevLabel === '전체 합계' || prevLabel.length === 0;
      if (prevIsBreak) break;
      start -= 1;
    }

    outWs.getCell(`N${r}`).value = { formula: `SUM(N${start}:N${end})` };
    outWs.getCell(`O${r}`).value = { formula: `SUM(O${start}:O${end})` };
  }

  // Rebuild grand total formulas from actual subtotal rows.
  let totalRowIndex = 0;
  const subtotalRowsForTotal = [];
  for (let r = 2; r <= outWs.rowCount; r += 1) {
    const label = String(outWs.getCell(`A${r}`).value || '');
    if (label.endsWith('소계')) subtotalRowsForTotal.push(r);
    if (label === '전체 합계') totalRowIndex = r;
  }
  if (totalRowIndex > 0 && subtotalRowsForTotal.length > 0) {
    if (subtotalRowsForTotal.length === 1) {
      outWs.getCell(`N${totalRowIndex}`).value = { formula: `N${subtotalRowsForTotal[0]}` };
      outWs.getCell(`O${totalRowIndex}`).value = { formula: `O${subtotalRowsForTotal[0]}` };
    } else {
      const claimRefs = subtotalRowsForTotal.map((x) => `N${x}`).join(',');
      const paidRefs = subtotalRowsForTotal.map((x) => `O${x}`).join(',');
      outWs.getCell(`N${totalRowIndex}`).value = { formula: `SUM(${claimRefs})` };
      outWs.getCell(`O${totalRowIndex}`).value = { formula: `SUM(${paidRefs})` };
    }
  }

  await outWb.xlsx.writeFile(outputPath);

  console.log(`변환 완료: ${outputPath}`);
  console.log(`처리 건수: ${rows.length}`);
}

main().catch((err) => {
  console.error('오류:', err.message);
  process.exit(1);
});
