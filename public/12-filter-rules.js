import { createTasks, TASK_STATUSES, TASK_OWNERS } from './task-data.js';

const controls = document.getElementById('lesson-tools');
controls.innerHTML = `
  <label>状态 <select id="filter-status">${TASK_STATUSES.map(status => `<option>${status}</option>`).join('')}</select></label>
  <label>负责人 <select id="filter-owner">${TASK_OWNERS.map(owner => `<option>${owner}</option>`).join('')}</select></label>
  <label><input type="checkbox" id="include-urgent"> 或优先级为 1</label>
  <button id="apply-task-filter">应用组合筛选</button>
  <button id="clear-task-filter">显示全部</button>
  <button id="inspect-task-filter">打印条件与记录</button>
`;

const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('12 组合筛选');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 36;
[88, 240, 120, 120, 125].forEach((width, col) => sheet.setColumnWidth(col, width));

const records = createTasks();
const headers = ['编号', '任务标题', '状态', '负责人', '优先级'];
const fields = ['id', 'title', 'status', 'owner', 'priority'];
let condition = null;
let visible = records.slice();

function matches(record, node) {
  if (node.logic === 'and') return node.children.every(child => matches(record, child));
  if (node.logic === 'or') return node.children.some(child => matches(record, child));
  return record[node.field] === node.value;
}

class FilterProvider {
  getValue(row, col) {
    if (row === 0) return headers[col];
    if (!visible.length) return col === 1 ? '没有符合条件的任务' : null;
    return visible[row - 1]?.[fields[col]] ?? null;
  }

  setValue(row, col, value) {
    const record = visible[row - 1];
    const title = String(value ?? '').trim();
    if (record && col === 1 && title) {
      record.title = title;
      console.log('标题写回：', { displayRow: row, id: record.id, title });
    }
    return true;
  }

  getStyle(row, col) {
    const style = new S.Style();
    style.locked = row === 0 || !visible[row - 1] || col !== 1;
    style.vAlign = S.VerticalAlign.center;
    style.backColor = row === 0 ? '#24634d' : col === 1 && visible.length ? '#edf7f1' : '#ffffff';
    style.foreColor = row === 0 ? '#ffffff' : '#243346';
    if (!style.locked) style.formatter = '@';
    return style;
  }
}

const provider = new FilterProvider();
const dataRange = sheet.dataRanges.add('filter-rules', provider, new S.Range(0, 0, visible.length + 1, fields.length));
const status = document.getElementById('filter-status');
const owner = document.getElementById('filter-owner');
const urgent = document.getElementById('include-urgent');
status.value = '进行中';

function applyFilter(next) {
  if (!sheet.endEdit()) return;
  condition = next;
  // 显示列表保存原记录的引用，编辑不会误写 records[显示行号]。
  visible = condition ? records.filter(record => matches(record, condition)) : records.slice();
  dataRange.range(new S.Range(0, 0, Math.max(1, visible.length) + 1, fields.length));
  dataRange.repaint();
}

document.getElementById('apply-task-filter').onclick = () => {
  const and = {
    logic: 'and',
    children: [{ field: 'status', value: status.value }, { field: 'owner', value: owner.value }]
  };
  applyFilter(urgent.checked ? { logic: 'or', children: [and, { field: 'priority', value: 1 }] } : and);
};
document.getElementById('clear-task-filter').onclick = () => applyFilter(null);
document.getElementById('inspect-task-filter').onclick = () => {
  console.log('当前条件：', condition);
  console.table(visible);
  console.table(records);
};

window.demo = {
  spread, sheet, records, provider, dataRange,
  get visible() { return visible; },
  get condition() { return condition; }
};
