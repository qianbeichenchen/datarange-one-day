// 第三课：把已学会的读写逻辑扩展成基础商品报价表。
const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1,
  allowUndo: false, // 自定义数据的撤销需要 changes + undo 钩子，见教程选学。
  allowUserDragDrop: false,
  allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('第三课：报价表');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 30;
sheet.setColumnWidth(1, 180);
for (let col = 2; col <= 6; col++) sheet.setColumnWidth(col, 120);

const rows = [
  { name: '键盘', price: 199, qty: 2 },
  { name: '鼠标', price: 89, qty: 3 },
  { name: '显示器', price: 1299, qty: 1 }
];
const fields = ['name', 'price', 'qty'];
const headers = ['商品', '单价', '数量', '金额'];

class ProductProvider {
  constructor(rows) { this.rows = rows; }

  getValue(row, col) {
    if (row === 0) return headers[col] ?? null;
    const item = this.rows[row - 1]; // 减去我们自己定义的 1 行表头。
    if (!item) return null;
    if (col === 3) return Math.round(item.price * item.qty * 100) / 100;
    return item[fields[col]] ?? null;
  }

  setValue(row, col, value) {
    const item = this.rows[row - 1];
    if (!item || col < 0 || col > 2) return true; // 表头/计算列不写回。
    const field = fields[col];
    let next;
    if (field === 'name') {
      next = String(value ?? '').trim();
      if (!next) { console.warn('商品名称不能为空。'); return true; }
    } else {
      next = Number(value);
      if (value === '' || value == null || !Number.isFinite(next) || next < 0 ||
          (field === 'qty' && !Number.isInteger(next))) {
        console.warn('单价需为非负数，数量需为非负整数。');
        return true;
      }
    }
    item[field] = next;
    console.log('setValue（区域坐标）', { row, col, field, value: next });
    // 等当前写入调用完成后重绘，使金额和高亮也立即更新。
    queueMicrotask(() => this.host.repaint()); // host 由 SpreadJS 注入，是 DataRange 实例。
    return true;
  }

  getStyle(row, col) {
    const style = new S.Style();
    style.vAlign = S.VerticalAlign.center;
    style.locked = row === 0 || col === 3;
    if (row === 0) {
      style.backColor = '#24634d';
      style.foreColor = '#ffffff';
      style.font = 'bold 13px Microsoft YaHei';
    } else {
      style.backColor = row % 2 === 0 ? '#f1f6f3' : '#ffffff';
      if (col === 1 || col === 3) style.formatter = '0.00';
      if (col === 2) style.formatter = '0';
      if (col === 3 && this.getValue(row, col) >= 1000) {
        style.backColor = '#fff0d6';
        style.foreColor = '#925700';
      }
    }
    return style;
  }
}

const provider = new ProductProvider(rows);
const dataRange = sheet.dataRanges.add(
  'products', provider,
  new S.Range(2, 1, rows.length + 1, headers.length),
  { sticky: { top: 1 } }
);

// 只有长度变化才调整 range；仅改值时只需 repaint。
function resizeDataRange() {
  const old = dataRange.range();
  dataRange.range(new S.Range(old.row, old.col, rows.length + 1, headers.length));
  dataRange.repaint();
}

document.getElementById('actions').hidden = false;
document.getElementById('add').onclick = () => {
  rows.push({ name: '新商品', price: 100, qty: 1 });
  resizeDataRange();
};
document.getElementById('remove').onclick = () => {
  if (rows.length === 0) return;
  rows.pop();
  resizeDataRange(); // 删除到 0 条时，仍保留一行表头。
};
document.getElementById('update').onclick = () => {
  if (!rows.length) return;
  rows[0].qty += 1;
  dataRange.repaint(); // 直接修改对象不经过 setValue，需主动通知重绘。
};
document.getElementById('move').onclick = () => {
  const old = dataRange.range();
  const atStart = old.row === 2 && old.col === 1;
  dataRange.range(new S.Range(atStart ? 5 : 2, atStart ? 3 : 1, old.rowCount, old.colCount));
  dataRange.repaint();
  console.log('区域已移动，provider 内部坐标不变：', dataRange.range());
};
document.getElementById('inspect').onclick = () => {
  console.table(rows);
  console.log('当前工作表区域：', dataRange.range());
};

window.demo = { spread, sheet, rows, provider, dataRange };
console.log('第三课已加载。F12 查看写回日志；刷新恢复初始数据。');
