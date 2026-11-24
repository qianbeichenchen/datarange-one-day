import { createOrders, groupReport, amount, numericInput, roundMoney } from './business-model.js';

const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('⑤ 分组与联动');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 29;
[85, 140, 75, 75, 90, 70, 100].forEach((width, col) => sheet.setColumnWidth(col + 1, width));
sheet.setColumnWidth(8, 28);
[120, 75, 110].forEach((width, col) => sheet.setColumnWidth(col + 9, width));
const records = createOrders();
let region = '';
let report = groupReport(records, region);
let mergedRanges = buildMergedRanges();
const fields = ['id', 'name', 'region', 'category', 'price', 'qty'];
const headers = ['订单号', '商品', '地区', '品类', '单价', '数量', '金额'];

function buildMergedRanges() {
  const spans = [new S.Range(0, 0, 1, 4), new S.Range(0, 4, 1, 3)];
  report.display.forEach((line, index) => {
    if (!['band', 'header', 'data'].includes(line.kind)) spans.push(new S.Range(index, 0, 1, 6));
  });
  return spans;
}

class GroupProvider {
  getValue(row, col) {
    const line = report.display[row];
    if (!line) return null;
    if (line.kind === 'band') return col === 0 ? '订单资料' : col === 4 ? '报价 / 元' : null;
    if (line.kind === 'header') return headers[col] ?? null;
    if (line.kind === 'data') return col === 6 ? amount(line.item) : line.item[fields[col]] ?? null;
    if (col === 0) return `${line.kind === 'category' ? '    ' : ''}${line.label}（${line.count} 笔）`;
    return col === 6 ? line.amount : null;
  }
  setValue(row, col, value) {
    const line = report.display[row];
    if (line?.kind !== 'data' || ![4, 5].includes(col)) return true;
    const next = numericInput(value, { integer: col === 5 });
    if (next === null) { console.warn('单价为非负数，数量为非负整数。'); return true; }
    line.item[fields[col]] = next;
    console.log('分组明细写回：', { id: line.item.id, field: fields[col], next });
    queueMicrotask(refreshReport); // 同时重建分组汇总和右侧统计。
    return true;
  }
  getStyle(row, col) {
    const line = report.display[row];
    const kind = line?.kind;
    const style = new S.Style();
    style.vAlign = S.VerticalAlign.center;
    style.locked = kind !== 'data' || ![4, 5].includes(col);
    const colors = { band: '#24634d', header: '#357b61', region: '#cce4d7', category: '#e9f3ec', subtotal: '#f3f6e9', regionTotal: '#dcebdc', total: '#24634d' };
    style.backColor = colors[kind] || '#ffffff';
    if (['band', 'header', 'total'].includes(kind)) style.foreColor = '#ffffff';
    if (kind !== 'data') style.font = 'bold 12px Microsoft YaHei';
    if (row > 1 && [4, 6].includes(col)) style.formatter = '0.00';
    return style;
  }
  getSpans(row, col, rowCount, colCount) {
    // 所有合并坐标都是区域内部坐标，移动 DataRange 不用改变它们。
    // 合并列表只在布局变化时重建，不在每次绘制回调里重新计算。
    return mergedRanges.filter(span => span.intersect(row, col, rowCount, colCount));
  }
}

class SummaryProvider {
  getValue(row, col) {
    if (row === 0) return ['品类统计', '数量', '金额'][col] ?? null;
    const item = report.categories[row - 1];
    if (item) return [item.category, item.qty, item.amount][col] ?? null;
    if (row === report.categories.length + 1) return ['总计', report.total.qty, report.total.amount][col] ?? null;
    return null;
  }
  getStyle(row, col) {
    const style = new S.Style();
    style.locked = true;
    style.backColor = row === 0 ? '#245b75' : '#edf4f8';
    if (row === 0) style.foreColor = '#ffffff';
    if (row > 0 && col === 2) style.formatter = '0.00';
    return style;
  }
  setValue() { return true; }
}

const provider = new GroupProvider();
const summaryProvider = new SummaryProvider();
const dataRange = sheet.dataRanges.add('group-report', provider, new S.Range(2, 1, report.display.length, 7), { sticky: { top: 2 } });
const summaryRange = sheet.dataRanges.add('category-summary', summaryProvider, new S.Range(2, 9, report.categories.length + 2, 3));

function refreshReport() {
  report = groupReport(records, region);
  mergedRanges = buildMergedRanges();
  const old = dataRange.range();
  // 确保宿主有足够的行；预留空间，避免数据区域超出工作表。
  sheet.setRowCount(Math.max(100, old.row + report.display.length + 10));
  dataRange.range(new S.Range(old.row, old.col, report.display.length, 7));
  const summary = summaryRange.range();
  summaryRange.range(new S.Range(summary.row, summary.col, report.categories.length + 2, 3));
  dataRange.repaint();
  summaryRange.repaint();
}

document.getElementById('groups-controls').hidden = false;
document.getElementById('apply-group').onclick = () => {
  if (!sheet.endEdit()) return;
  region = document.getElementById('group-region').value;
  refreshReport();
};
document.getElementById('raise-price').onclick = () => {
  if (!sheet.endEdit()) return;
  for (const item of report.filtered) item.price = roundMoney(item.price * 1.1);
  refreshReport();
};
document.getElementById('move-report').onclick = () => {
  if (!sheet.endEdit()) return;
  const nextRow = dataRange.range().row === 2 ? 5 : 2;
  for (const range of [dataRange, summaryRange]) {
    const old = range.range();
    range.range(new S.Range(nextRow, old.col, old.rowCount, old.colCount));
  }
  refreshReport();
};
document.getElementById('inspect-groups').onclick = () => {
  console.table(records);
  console.table(report.display.map((line, row) => ({ row, kind: line.kind, id: line.item?.id, label: line.label, amount: line.kind === 'data' ? amount(line.item) : line.amount })));
  console.table(report.categories);
};
window.demo = { spread, sheet, records, provider, dataRange, summaryRange, get report() { return report; } };
console.log('第五课：只开放明细的单价、数量；小计、合计和右侧统计全部由同一份 records 计算。');
