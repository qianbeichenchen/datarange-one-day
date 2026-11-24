import { createOrders } from './demo-data.js';
import { orderView, amount, summarize, numericInput } from './business-model.js';

const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: true, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('④ 订单视图');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 32;
[90, 150, 90, 90, 110, 90, 120].forEach((width, col) => sheet.setColumnWidth(col + 1, width));

const records = createOrders();
const byId = new Map(records.map(item => [item.id, item]));
const undoManager = spread.undoManager();
const viewOptions = { keyword: '', region: '', sort: 'id' };
const fields = ['id', 'name', 'region', 'category', 'price', 'qty'];
const headers = ['订单号', '商品', '地区', '品类', '单价', '数量', '金额'];

class OrderProvider {
  constructor() {
    this.visible = orderView(records, viewOptions);
    this.total = summarize(this.visible);
  }
  getValue(row, col) {
    if (row === 0) return headers[col] ?? null;
    if (row === this.visible.length + 1) {
      if (col === 0) return '视图合计';
      if (col === 1) return `${this.total.count} 条记录`;
      return col === 5 ? this.total.qty : col === 6 ? this.total.amount : null;
    }
    const item = this.visible[row - 1]; // 显示下标，不是原始数组下标！
    return item ? (col === 6 ? amount(item) : item[fields[col]] ?? null) : null;
  }
  setValue(row, col, value, changes) {
    const item = this.visible[row - 1];
    if (!item || ![1, 4, 5].includes(col)) return true;
    const field = fields[col];
    const next = col === 1 ? String(value ?? '').trim() : numericInput(value, { integer: col === 5 });
    if (next === null || next === '') { console.warn('名称不能为空，单价为非负数，数量为非负整数。'); return true; }
    if (Object.is(item[field], next)) return true;
    // 官方机制：changes 由 SpreadJS 提供并关联到当前命令，无需自己维护历史栈。
    if (changes) changes.push({ type: 'value', id: item.id, field, oldValue: item[field] });
    item[field] = next;
    console.log('setValue：', { row, col, id: item.id, field, next, recorded: !!changes });
    // 只重算金额、合计并重绘；不改变显示顺序，保持原生命令的坐标映射。
    queueMicrotask(refreshValues);
    return true;
  }
  undo(change) {
    if (change.type !== 'value') return;
    const item = byId.get(change.id);
    if (!item) return;
    item[change.field] = change.oldValue;
    console.log('provider.undo：', change);
    // 恢复数据，不向 changes 再写入一次；重做由 SpreadJS 重放原命令。
    queueMicrotask(refreshValues);
  }
  getStyle(row, col) {
    const style = new S.Style();
    const isData = row > 0 && row <= this.visible.length;
    style.locked = !isData || ![1, 4, 5].includes(col);
    style.vAlign = S.VerticalAlign.center;
    style.backColor = row === 0 ? '#245b75' : !isData ? '#dceaf2' : row % 2 === 0 ? '#f3f7fa' : '#ffffff';
    if (row === 0) style.foreColor = '#ffffff';
    if (row !== 0 && (col === 4 || col === 6)) style.formatter = '0.00';
    return style;
  }
}

const provider = new OrderProvider();
const dataRange = sheet.dataRanges.add('order-view', provider, new S.Range(2, 1, provider.visible.length + 2, 7), { sticky: { top: 1 } });

function refreshView() {
  provider.visible = orderView(records, viewOptions);
  const old = dataRange.range();
  dataRange.range(new S.Range(old.row, old.col, provider.visible.length + 2, 7));
  refreshValues();
}

function refreshValues() {
  provider.total = summarize(provider.visible);
  dataRange.repaint();
  document.getElementById('undo-edit').disabled = !undoManager.canUndo();
  document.getElementById('redo-edit').disabled = !undoManager.canRedo();
}

document.getElementById('orders-controls').hidden = false;
document.getElementById('apply-view').onclick = () => {
  if (!sheet.endEdit()) return;
  viewOptions.keyword = document.getElementById('keyword').value;
  viewOptions.region = document.getElementById('order-region').value;
  viewOptions.sort = document.getElementById('order-sort').value;
  // 原生重做会重放单元格坐标；换布局前结束编辑并清空历史，避免写到其他记录。
  undoManager.clear();
  refreshView();
  console.log('已应用视图并清空原布局的撤销/重做记录。');
};
document.getElementById('undo-edit').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.undo();
  refreshValues();
};
document.getElementById('redo-edit').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.redo();
  refreshValues();
};
document.getElementById('inspect-orders').onclick = () => {
  console.table(records);
  console.table(provider.visible.map((item, index) => ({ displayRow: index + 1, id: item.id, amount: amount(item) })));
  console.log('当前视图：', viewOptions);
  console.log('原生撤销栈：', undoManager.getUndoStack());
  console.log('原生重做栈：', undoManager.getRedoStack());
};
refreshView();
window.demo = { spread, sheet, records, provider, dataRange, undoManager, viewOptions };
console.log('第四课：setValue(changes) + undo(change) 接入原生历史。支持页面按钮和表格焦点下 Ctrl+Z / Ctrl+Y。编辑不自动重排；应用视图会清空历史。');
