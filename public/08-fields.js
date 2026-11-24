import { createTasks } from './task-data.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <label><input id="show-owner" type="checkbox" checked style="width:auto"> 显示负责人</label>
  <button id="swap-fields">交换状态与负责人</button>
  <button id="inspect-fields">打印列映射与记录</button>`;

const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('08 动态字段');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 32;

const records = createTasks();
const fields = [
  { id: 'title', name: '任务', width: 220 },
  { id: 'status', name: '状态', width: 105 },
  { id: 'owner', name: '负责人', width: 120 }
];

// 字段定义保持不变，显示顺序与隐藏状态由视图维护。
const view = { columns: fields.map(field => field.id), showOwner: true };
let visibleFields = fields.slice();
const origin = { row: 2, col: 1 };

const provider = {
  getValue(row, col) {
    const field = visibleFields[col];
    if (!field) return null;
    return row === 0 ? field.name : records[row - 1]?.[field.id] ?? null;
  },
  setValue(row, col, value) {
    const record = records[row - 1];
    const field = visibleFields[col];
    if (!record || !field) return true;
    const next = String(value ?? '').trim();
    if (!next) {
      console.warn(`${field.name}不能为空。`);
      return true;
    }
    record[field.id] = next;
    console.log('写回任务字段', { recordId: record.id, viewCol: col, fieldId: field.id, value: next });
    queueMicrotask(() => dataRange.repaint());
    return true;
  },
  getStyle(row) {
    const style = new S.Style();
    style.locked = row === 0;
    style.vAlign = S.VerticalAlign.center;
    style.backColor = row === 0 ? '#245b75' : row % 2 ? '#ffffff' : '#f3f7fa';
    if (row === 0) style.foreColor = '#ffffff';
    return style;
  }
};

const dataRange = sheet.dataRanges.add('task-fields', provider,
  new S.Range(origin.row, origin.col, records.length + 1, visibleFields.length));

function applyColumns() {
  visibleFields = view.columns
    .filter(id => id !== 'owner' || view.showOwner)
    .map(id => fields.find(field => field.id === id));
  spread.suspendPaint();
  try {
    visibleFields.forEach((field, col) => sheet.setColumnWidth(origin.col + col, field.width));
    dataRange.range(new S.Range(origin.row, origin.col, records.length + 1, visibleFields.length));
    dataRange.repaint();
  } finally {
    spread.resumePaint();
  }
}

const ownerToggle = document.getElementById('show-owner');
const swapButton = document.getElementById('swap-fields');
ownerToggle.onchange = () => {
  if (!sheet.endEdit()) { ownerToggle.checked = view.showOwner; return; }
  view.showOwner = ownerToggle.checked;
  swapButton.disabled = !view.showOwner;
  applyColumns();
};
swapButton.onclick = () => {
  if (!sheet.endEdit()) return;
  const [title, second, third] = view.columns;
  view.columns = [title, third, second];
  applyColumns();
};
document.getElementById('inspect-fields').onclick = () => {
  console.table(visibleFields.map((field, col) => ({ col, fieldId: field.id, name: field.name })));
  console.table(records);
};

applyColumns();
window.demo = { spread, sheet, dataRange, provider, records, fields, view, get visibleFields() { return visibleFields; } };
console.log('第 08 课。交换状态与负责人列，再双击负责人单元格编辑。检查列号如何通过字段 ID 找到原值。');
