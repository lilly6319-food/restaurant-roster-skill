import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';

const runtimeModules = process.env.CODEX_NODE_MODULES
  ?? path.join(os.homedir(), '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'node', 'node_modules');
const runtimeRequire = createRequire(path.join(runtimeModules, '_restaurant_roster_loader.cjs'));
const { FileBlob, SpreadsheetFile } = runtimeRequire('@oai/artifact-tool');

const args = process.argv.slice(2);
const valueOf = flag => {
  const i = args.indexOf(flag);
  return i >= 0 ? args[i + 1] : null;
};
const valuesOf = flag => args.flatMap((v, i) => v === flag && args[i + 1] ? [args[i + 1]] : []);
const basePath = valueOf('--base');
const outputPath = valueOf('--output');
const monthText = valueOf('--month');
const weekPaths = valuesOf('--week');
const holidayTexts = valuesOf('--holiday');
const holidaysConfirmed = args.includes('--holidays-confirmed') || holidayTexts.length > 0;
const replace = args.includes('--replace');
const previewPath = valueOf('--preview');
const auditDir = valueOf('--audit-dir');

assert(basePath, '需要 --base <含週班表的xlsx>');
assert(outputPath, '需要 --output <輸出xlsx>');
assert(/^\d{4}-\d{2}$/.test(monthText ?? ''), '需要 --month YYYY-MM');

const [year, month] = monthText.split('-').map(Number);
assert(month >= 1 && month <= 12, '月份必須介於01到12');
const monthStart = new Date(Date.UTC(year, month - 1, 1));
const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
const uniqueHolidayTexts = [...new Set(holidayTexts)];
assert.equal(uniqueHolidayTexts.length, holidayTexts.length, '國定假日日期不可重複');
const holidayDates = uniqueHolidayTexts.map(text => {
  assert(/^\d{4}-\d{2}-\d{2}$/.test(text), `國定假日需使用 YYYY-MM-DD：${text}`);
  const date = new Date(`${text}T00:00:00Z`);
  assert(!Number.isNaN(date.valueOf()) && date.toISOString().slice(0, 10) === text, `無效國定假日：${text}`);
  assert(date.getUTCFullYear() === year && date.getUTCMonth() === month - 1, `國定假日不在目標月份：${text}`);
  return date;
});
const regularRestDays = Array.from({ length: daysInMonth }, (_, i) => new Date(Date.UTC(year, month - 1, i + 1)))
  .filter(date => [3, 6].includes(date.getUTCDay())).length;
const overlappingHolidays = holidayDates.filter(date => [3, 6].includes(date.getUTCDay()));
assert.equal(overlappingHolidays.length, 0, `國定假日與週三／週六重疊，須先確認是否另計：${overlappingHolidays.map(d => d.toISOString().slice(0, 10)).join('、')}`);
const monthlyRestSummary = holidaysConfirmed
  ? `應休${regularRestDays + holidayDates.length}天（含國定假日${holidayDates.length}天）`
  : '應休待確認（含國定假日待確認）';
const staff = [
  ['店長', 'Meryl'], ['正職', 'Tom'], ['正職', 'Emma'], ['工讀', 'Brad'],
  ['工讀', 'Keanu'], ['工讀', 'Denzel'], ['工讀', 'Will'], ['工讀', 'Leo'],
  ['工讀', 'Zendaya'], ['工讀', 'Morgan'], ['工讀', 'Scarlett'],
];
const staffRow = new Map(staff.map(([, name], i) => [name, 8 + i]));
const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
const helperStartColumn = 38; // AL；保存每日淨工時，避免由顯示班次推算工時
const C = {
  ink: '#243746', teal: '#245C58', pale: '#EAF3F0', orange: '#FCE4D1',
  amber: '#FFF2CC', green: '#E2F0D9', gray: '#ECEFF1', line: '#C9D3D8', blue: '#1D4ED8',
};

function excelDate(value) {
  if (value instanceof Date) return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
  if (typeof value === 'string' && value.trim()) {
    const d = new Date(value);
    if (!Number.isNaN(d.valueOf())) return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }
  return null;
}

function col(n) {
  let s = '';
  for (; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

function sameMonth(date) {
  return date && date.getUTCFullYear() === year && date.getUTCMonth() === month - 1;
}

function hoursFromRow(row) {
  const cached = Number(row[11]);
  if (Number.isFinite(cached)) return cached;
  const start = Number(row[6]);
  const end = Number(row[7]);
  if (!Number.isFinite(start) || !Number.isFinite(end)) return 0;
  const breakStart = Number(row[8]);
  const breakEnd = Number(row[9]);
  const rest = Number.isFinite(breakStart) && Number.isFinite(breakEnd) ? (breakEnd - breakStart) * 24 : 0;
  return (end - start) * 24 - rest;
}

function hourFromCell(value) {
  if (value instanceof Date) return value.getUTCHours() + value.getUTCMinutes() / 60;
  const n = Number(value);
  if (!Number.isFinite(n)) return null;
  return n >= 0 && n < 1 ? n * 24 : n;
}

function compactHour(value) {
  const raw = hourFromCell(value);
  assert(raw !== null, `無法解析班次時間：${value}`);
  const rounded = Math.round(raw * 2) / 2;
  assert(Math.abs(raw - rounded) < 1e-6, `整月工時表只支援整點或半點：${raw}`);
  return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
}

function shiftDisplayFromRow(row) {
  const hours = Math.round(hoursFromRow(row) * 10) / 10;
  if (hours <= 0) return { display: 0, hours: 0 };
  assert(row[6] !== null && row[6] !== '' && row[7] !== null && row[7] !== '', `有工時但缺少上下班時間：${row[2]}`);
  return { display: `${compactHour(row[6])}-${compactHour(row[7])}`, hours };
}

async function getSheetOrNull(workbook, name) {
  const listed = await workbook.inspect({ kind: 'sheet', include: 'id,name', maxChars: 5000 });
  if (!listed.ndjson.includes(`"name":"${name}"`)) return null;
  return workbook.worksheets.getItem(name);
}

function buildSheet(sheet) {
  sheet.showGridLines = false;
  sheet.tabColor = C.teal;
  sheet.getRange('A1:AJ25').format.font = { name: 'Microsoft JhengHei', size: 10, color: C.ink };
  sheet.getRange('A1:AJ25').format.verticalAlignment = 'center';
  sheet.getRange('A1:AJ25').format.rowHeight = 23;

  sheet.getRange('A1:AJ1').merge();
  sheet.getRange('A3:AJ3').merge();
  for (const area of ['A2:D2', 'E2:H2', 'I2:M2', 'N2:R2', 'A5:A7', 'B5:B7', 'AH5:AH7', 'AI5:AI7', 'AJ5:AJ7', 'A20:B20', 'A21:B21', 'A23:AJ25']) sheet.getRange(area).merge();

  sheet.getRange('A1').values = [['整月工時表']];
  sheet.getRange('A1:AJ1').format = { fill: C.teal, font: { name: 'Microsoft JhengHei', bold: true, color: '#FFFFFF', size: 16 }, rowHeight: 36 };
  sheet.getRange('A2:W2').format.borders = { preset: 'all', style: 'thin', color: C.line };
  sheet.getRange('A2:W2').format.horizontalAlignment = 'center';
  sheet.getRange('A2:W2').format.fill = C.pale;
  sheet.getRange('A2:W2').format.font = { name: 'Microsoft JhengHei', size: 10, bold: true, color: C.ink };
  sheet.getRange('A3').values = [['每日格顯示精簡班次（11:00～21:00＝11-21；11:30～21:30＝11.5-21.5）。空白＝尚未匯入；0＝已確認但未排班。月總仍按扣除休息後淨工時計算。']];
  sheet.getRange('A3:AJ3').format = { fill: C.pale, font: { name: 'Microsoft JhengHei', size: 10, color: C.ink }, rowHeight: 25 };

  sheet.getRange('A5').values = [['身分']];
  sheet.getRange('B5').values = [['姓名']];
  sheet.getRange('AH5').values = [['月總工時']];
  sheet.getRange('AI5').values = [['工作天']];
  sheet.getRange('AJ5').values = [['資料狀態']];
  sheet.getRange('A5:AJ7').format.horizontalAlignment = 'center';
  sheet.getRange('A5:AJ7').format.borders = { preset: 'all', style: 'thin', color: '#FFFFFF' };
  sheet.getRange('A5:AJ7').format.fill = '#23394A';
  sheet.getRange('A5:AJ7').format.font = { name: 'Microsoft JhengHei', size: 9, bold: true, color: '#FFFFFF' };
  sheet.getRange('C7:AG7').format.font = { name: 'Microsoft JhengHei', size: 8, bold: false, color: C.ink };

  sheet.getRange('A8:B18').values = staff;
  sheet.getRange('A8:AJ18').format.borders = { preset: 'all', style: 'thin', color: C.line };
  sheet.getRange('A8:B18').format.horizontalAlignment = 'left';
  sheet.getRange('C8:AJ18').format.horizontalAlignment = 'center';
  sheet.getRange('C8:AG18').format.font = { name: 'Microsoft JhengHei', size: 8, color: C.blue };
  sheet.getRange('AH8:AH18').format.numberFormat = '0.0';
  sheet.getRange('AI8:AI18').format.numberFormat = '0';
  sheet.getRange('AJ8:AJ18').conditionalFormats.add('containsText', { text: '待匯入', format: { fill: C.amber, font: { color: C.ink } } });
  sheet.getRange('AJ8:AJ18').conditionalFormats.add('containsText', { text: '部分匯入', format: { fill: C.amber, font: { color: C.ink } } });
  sheet.getRange('AJ8:AJ18').conditionalFormats.add('containsText', { text: '完整', format: { fill: C.green, font: { color: C.ink } } });

  sheet.getRange('A20').values = [['全店每日工時']];
  sheet.getRange('A21').values = [['每日排班人數']];
  sheet.getRange('A20:AJ21').format.borders = { preset: 'all', style: 'thin', color: C.line };
  sheet.getRange('A20:AJ21').format.fill = C.pale;
  sheet.getRange('A20:AJ21').format.font = { name: 'Microsoft JhengHei', size: 9, bold: true, color: C.ink };
  sheet.getRange('C20:AH20').format.numberFormat = '0.0';
  sheet.getRange('C21:AG21').format.numberFormat = '0';
  sheet.getRange('A23').values = [['每日格只顯示精簡上下班時間；月總工時、工作天與每日合計取自每週「班次明細」的預估淨工時。週班表更正時須重新匯入該週，不直接手改月表。本表不取代打卡、實際休息、假別認定或薪資核算。']];
  sheet.getRange('A23:AJ25').format = { wrapText: true, font: { name: 'Microsoft JhengHei', size: 10, color: C.ink }, rowHeight: 25 };

  sheet.getRange('A1:A25').format.columnWidth = 9;
  sheet.getRange('B1:B25').format.columnWidth = 10;
  sheet.getRange('C1:AG25').format.columnWidth = 6.4;
  sheet.getRange('AH1:AH25').format.columnWidth = 12;
  sheet.getRange('AI1:AI25').format.columnWidth = 9;
  sheet.getRange('AJ1:AJ25').format.columnWidth = 12;
  sheet.freezePanes.freezeRows(7);
  sheet.freezePanes.freezeColumns(2);
}

function configureCalculations(sheet) {
  for (const range of ['I2:M2', 'N2:R2', 'S2:W2']) sheet.getRange(range).merge();
  sheet.getRange('I2').values = [[monthlyRestSummary]];
  sheet.getRange('I2:M2').format = {
    borders: { preset: 'all', style: 'thin', color: C.line },
    horizontalAlignment: 'center',
    verticalAlignment: 'center',
    fill: C.pale,
    font: { name: 'Microsoft JhengHei', size: 10, bold: true, color: C.ink },
  };
  sheet.getRange('N2').formulas = [['="已匯入日數："&COUNTIF(C7:AG7,"已確認")']];
  sheet.getRange('S2').formulas = [[`=IF(COUNTIF(C7:AG7,"已確認")=${daysInMonth},"月表狀態：整月完成","月表狀態：待匯入")`]];
  sheet.getRange('S2:W2').conditionalFormats.deleteAll();
  sheet.getRange('S2:W2').conditionalFormats.add('containsText', { text: '待匯入', format: { fill: C.amber, font: { color: C.ink, bold: true } } });
  sheet.getRange('S2:W2').conditionalFormats.add('containsText', { text: '整月完成', format: { fill: C.green, font: { color: C.ink, bold: true } } });
  sheet.getRange('A3').values = [['每日格顯示精簡班次（11:00～21:00＝11-21；11:30～21:30＝11.5-21.5）。空白＝尚未匯入；0＝已確認但未排班。月總仍按扣除休息後淨工時計算。']];
  sheet.getRange('A23').values = [['每日格只顯示精簡上下班時間；月總工時、工作天與每日合計取自每週「班次明細」的預估淨工時。週班表更正時須重新匯入該週，不直接手改月表。本表不取代打卡、實際休息、假別認定或薪資核算。']];
  sheet.getRange('C8:AG18').format.font = { name: 'Microsoft JhengHei', size: 8, color: C.blue };
  sheet.getRange('C1:AG25').format.columnWidth = 6.4;
  sheet.getRange('AL1').values = [['內部淨工時計算區']];
  sheet.getRange('AL2:BP18').format.font = { name: 'Microsoft JhengHei', size: 8, color: '#AEB8BD' };
  sheet.getRange('AL1:BP18').format.columnWidth = 2.5;
  for (let d = 1; d <= 31; d++) {
    const helperLetter = col(helperStartColumn + d - 1);
    sheet.getRange(`${helperLetter}5`).values = [[d <= daysInMonth ? d : null]];
  }
  for (let r = 8; r <= 18; r++) {
    sheet.getRange(`AH${r}`).formulas = [[`=IF(COUNT(AL${r}:BP${r})=0,"",SUM(AL${r}:BP${r}))`]];
    sheet.getRange(`AI${r}`).formulas = [[`=IF(COUNT(AL${r}:BP${r})=0,"",COUNTIF(AL${r}:BP${r},">0"))`]];
    sheet.getRange(`AJ${r}`).formulas = [[`=IF(COUNT(AL${r}:BP${r})=0,"待匯入",IF(COUNT(AL${r}:BP${r})=COUNT(C$5:AG$5),"完整","部分匯入"))`]];
  }
  sheet.getRange('AH20').formulas = [['=IF(COUNT(AL8:BP18)=0,"",SUM(AH8:AH18))']];
  sheet.getRange('AJ20').formulas = [['=IF(COUNT(AL8:BP18)=0,"待匯入",IF(COUNT(AL8:BP18)=COUNT(C5:AG5)*ROWS(C8:C18),"完整","部分匯入"))']];
  for (let d = 1; d <= daysInMonth; d++) {
    const visibleLetter = col(d + 2);
    const helperLetter = col(helperStartColumn + d - 1);
    sheet.getRange(`${visibleLetter}20`).formulas = [[`=IF(COUNT(${helperLetter}8:${helperLetter}18)=0,"",SUM(${helperLetter}8:${helperLetter}18))`]];
    sheet.getRange(`${visibleLetter}21`).formulas = [[`=IF(COUNT(${helperLetter}8:${helperLetter}18)=0,"",COUNTIF(${helperLetter}8:${helperLetter}18,">0"))`]];
  }
}

function resetMonth(sheet) {
  sheet.getRange('A1').values = [[`整月工時表｜${year}年${month}月`]];
  sheet.getRange('A2').values = [[`月份：${year}年${month}月`]];
  sheet.getRange('E2').values = [[`本月天數：${daysInMonth}`]];

  const dates = [];
  const dayNames = [];
  const states = [];
  sheet.getRange('C8:AG18').clear({ applyTo: 'contents' });
  sheet.getRange('AL8:BP18').clear({ applyTo: 'contents' });
  sheet.getRange('C5:AG21').format.fill = '#FFFFFF';
  for (let d = 1; d <= 31; d++) {
    if (d <= daysInMonth) {
      const date = new Date(Date.UTC(year, month - 1, d));
      dates.push(date);
      dayNames.push(weekdays[date.getUTCDay()]);
      states.push('待匯入');
    } else {
      dates.push(null);
      dayNames.push('—');
      states.push('不適用');
    }
  }
  sheet.getRange('C5:AG5').values = [dates];
  sheet.getRange('C5:AG5').setNumberFormat('d');
  sheet.getRange('C6:AG6').values = [dayNames];
  sheet.getRange('C7:AG7').values = [states];
  for (let d = 1; d <= 31; d++) {
    const letter = col(d + 2);
    if (d > daysInMonth) {
      sheet.getRange(`${letter}5:${letter}21`).format.fill = C.gray;
      sheet.getRange(`${letter}20:${letter}21`).clear({ applyTo: 'contents' });
    } else {
      const date = dates[d - 1];
      const weekend = [0, 6].includes(date.getUTCDay());
      sheet.getRange(`${letter}5:${letter}6`).format.fill = weekend ? C.orange : '#23394A';
      sheet.getRange(`${letter}5:${letter}6`).format.font = { name: 'Microsoft JhengHei', size: 9, bold: true, color: weekend ? C.ink : '#FFFFFF' };
      sheet.getRange(`${letter}7`).format.fill = C.amber;
    }
  }
  configureCalculations(sheet);
}

async function extractWeek(weekPath) {
  const weekly = await SpreadsheetFile.importXlsx(await FileBlob.load(weekPath));
  const detail = weekly.worksheets.getItem('班次明細');
  const rows = detail.getRange('A6:L200').values;
  const days = new Map();
  for (const row of rows) {
    const date = excelDate(row[0]);
    const name = typeof row[2] === 'string' ? row[2].trim() : '';
    if (!sameMonth(date) || !name) continue;
    assert(staffRow.has(name), `月表沒有此人員：${name}`);
    const key = date.toISOString().slice(0, 10);
    if (!days.has(key)) days.set(key, new Map());
    assert(!days.get(key).has(name), `重複的人員日期：${key} ${name}`);
    days.get(key).set(name, shiftDisplayFromRow(row));
  }
  for (const [date, people] of days) {
    const missing = staff.map(([, name]) => name).filter(name => !people.has(name));
    assert.equal(missing.length, 0, `${path.basename(weekPath)} ${date} 缺少人員：${missing.join('、')}`);
  }
  return days;
}

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(basePath));
let sheet = await getSheetOrNull(workbook, '整月工時表');
if (!sheet) {
  sheet = workbook.worksheets.add('整月工時表');
  buildSheet(sheet);
  resetMonth(sheet);
} else {
  const expectedTitle = `整月工時表｜${year}年${month}月`;
  if (sheet.getRange('A1').values[0][0] !== expectedTitle) resetMonth(sheet);
  else configureCalculations(sheet);
}

for (const weekPath of weekPaths) {
  const days = await extractWeek(weekPath);
  for (const [dateText, people] of days) {
    const date = new Date(`${dateText}T00:00:00Z`);
    const day = date.getUTCDate();
    const letter = col(day + 2);
    const helperLetter = col(helperStartColumn + day - 1);
    const state = sheet.getRange(`${letter}7`).values[0][0];
    for (const [name, entry] of people) {
      const row = staffRow.get(name);
      const currentDisplay = sheet.getRange(`${letter}${row}`).values[0][0];
      const currentHours = sheet.getRange(`${helperLetter}${row}`).values[0][0];
      const differs = currentDisplay !== entry.display || Math.abs(Number(currentHours) - entry.hours) > 1e-7;
      if (state === '已確認' && currentDisplay !== null && currentDisplay !== '' && differs && !replace) {
        throw new Error(`${dateText} ${name} 已有 ${currentDisplay}／${currentHours} 小時；若確認要覆蓋，請加 --replace`);
      }
      sheet.getRange(`${letter}${row}`).values = [[entry.display]];
      sheet.getRange(`${helperLetter}${row}`).values = [[entry.hours]];
    }
    sheet.getRange(`${letter}7`).values = [['已確認']];
    sheet.getRange(`${letter}7`).format.fill = C.green;
  }
}

workbook.recalculate();
const check = await workbook.inspect({
  kind: 'table', range: "'整月工時表'!A1:AJ21", include: 'values,formulas',
  tableMaxRows: 21, tableMaxCols: 36, maxChars: 12000,
});
console.log(check.ndjson);
const errors = await workbook.inspect({
  kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 100 }, summary: 'monthly hours formula error scan', maxChars: 3000,
});
console.log(errors.ndjson);

if (previewPath) {
  await fs.mkdir(path.dirname(previewPath), { recursive: true });
  const preview = await workbook.render({ sheetName: '整月工時表', range: 'A1:AJ25', scale: 1, format: 'png' });
  await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
}

if (auditDir) {
  const auditRanges = {
    '週班表': 'A1:K39', '班次明細': 'A1:P82', '人力缺口': 'A1:L173',
    '工時與休假': 'A1:U21', '閱讀與待確認': 'A1:D27', '規則與技能': 'A1:C38',
    '整月工時表': 'A1:AJ25',
  };
  await fs.mkdir(auditDir, { recursive: true });
  for (const [sheetName, range] of Object.entries(auditRanges)) {
    const image = await workbook.render({ sheetName, range, scale: 1, format: 'png' });
    await fs.writeFile(path.join(auditDir, `${sheetName}.png`), new Uint8Array(await image.arrayBuffer()));
  }
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`SAVED ${outputPath}`);
