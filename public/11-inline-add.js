import { createTasks, TASK_STATUSES } from './demo-data.js';

const controls = document.getElementById('lesson-tools');
controls.innerHTML = `
  <label>新任务状态
    <select id="draft-status">${TASK_STATUSES.map(status => `<option>${status}</option>`).join('')}</select>
  </label>
  <button id="new-draft">添加记录</button>
  <button id="cancel-draft" disabled>取消草稿</button>
  <button id="inspect-draft">打印正式记录与草稿</button>
`;

const S = GC.Spread.Sheets;
const host = document.getElementById('ss');
const spread = new S.Workbook(host, {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('11 行内新增');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 37;
[90, 275, 125].forEach((width, col) => sheet.setColumnWidth(col, width));

const records = createTasks();
const headers = ['编号', '任务标题', '状态'];
const fields = ['id', 'title', 'status'];
let nextId = records.length + 1;
let draft = null;
let visibleRows = [];

function rebuildRows() {
  visibleRows = [
    { kind: 'header' },
    ...records.map(record => ({ kind: 'record', record })),
    ...(draft ? [{ kind: 'draft', record: draft }] : [{ kind: 'add' }])
  ];
}

class DraftProvider {
  getValue(row, col) {
    const line = visibleRows[row];
    if (line?.kind === 'header') return headers[col];
    if (line?.kind === 'add') return col === 1 ? '＋ 添加记录' : null;
    if (line?.kind === 'draft' && col === 0) return '草稿';
    return line?.record?.[fields[col]] ?? null;
  }

  getStyle(row, col) {
    const kind = visibleRows[row]?.kind;
    const style = new S.Style();
    style.locked = kind !== 'draft' || col !== 1;
    style.vAlign = S.VerticalAlign.center;
    style.backColor = kind === 'header' ? '#24634d' : kind === 'draft' ? '#fff4d9' : kind === 'add' ? '#e8f3ed' : '#ffffff';
    style.foreColor = kind === 'header' ? '#ffffff' : '#243346';
    if (kind === 'draft' && col === 1) style.formatter = '@';
    return style;
  }

  setValue(row, col, value) {
    const line = visibleRows[row];
    if (line?.kind !== 'draft' || line.record !== draft || col !== 1) return true;
    const title = String(value ?? '').trim();
    if (!title) return true;
    // 草稿直到提交这一刻才进入正式数组。保持回调期间的坐标不变。
    records.push({ ...draft, id: `T${String(nextId++).padStart(2, '0')}`, title });
    draft = null;
    queueMicrotask(refresh);
    return true;
  }

  onMouseDown(row, col, event) {
    if (event.button !== 0 || col !== 1 || visibleRows[row]?.kind !== 'add') return false;
    queueMicrotask(beginDraft);
    return true;
  }
}

rebuildRows();
const provider = new DraftProvider();
const dataRange = sheet.dataRanges.add('inline-task', provider, new S.Range(0, 0, visibleRows.length, fields.length));
const cancelButton = document.getElementById('cancel-draft');
const addButton = document.getElementById('new-draft');
const statusSelect = document.getElementById('draft-status');

function refresh() {
  rebuildRows();
  if (visibleRows.length > sheet.getRowCount()) sheet.setRowCount(visibleRows.length);
  dataRange.range(new S.Range(0, 0, visibleRows.length, fields.length));
  dataRange.repaint();
  cancelButton.disabled = draft === null;
  addButton.disabled = draft !== null;
  statusSelect.disabled = draft !== null;
}

function beginDraft() {
  if (draft || !sheet.endEdit()) return;
  draft = { title: '', status: statusSelect.value, owner: '', priority: 2, start: '', due: '' };
  refresh();
  const row = visibleRows.length - 1;
  sheet.setActiveCell(row, 1);
  sheet.showCell(row, 1, S.VerticalPosition.nearest, S.HorizontalPosition.nearest);
  spread.focus();
  sheet.startEdit();
}

function cancelDraft() {
  if (!draft) return;
  sheet.endEdit(true);
  draft = null;
  refresh();
}

addButton.onclick = beginDraft;
cancelButton.onclick = cancelDraft;
// 按下按钮时先取消，避免输入框因失焦自动提交标题。
window.addEventListener('pointerdown', event => {
  if (event.button !== 0 || event.target !== cancelButton || !draft) return;
  event.preventDefault();
  event.stopPropagation();
  cancelDraft();
}, true);
// 在工作表处理 Escape 前截获它，让取消编辑同时丢弃草稿记录。
window.addEventListener('keydown', event => {
  if (event.key !== 'Escape' || event.isComposing || !draft || !host.contains(event.target)) return;
  event.preventDefault();
  event.stopPropagation();
  cancelDraft();
}, true);
document.getElementById('inspect-draft').onclick = () => {
  console.table(records);
  console.log('草稿：', draft);
};
window.demo = { spread, sheet, records, provider, dataRange, get draft() { return draft; } };
