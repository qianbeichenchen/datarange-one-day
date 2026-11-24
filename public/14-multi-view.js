import { createTasks, TASK_STATUSES } from './task-data.js';
import { BoardProvider, renderBoard, showTaskDetail } from './board-layout.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <label>视图 <select id="task-view"><option value="table">表格</option><option value="board">看板</option></select></label>
  <label>状态筛选 <select id="view-status"><option value="">全部</option></select></label>
  <label id="view-owner-label"><input id="view-owner" type="checkbox"> 显示负责人</label>
  <button id="inspect-views" type="button">打印记录与视图设置</button>
`;
const viewSelect = document.getElementById('task-view');
const statusSelect = document.getElementById('view-status');
const ownerToggle = document.getElementById('view-owner');
const ownerLabel = document.getElementById('view-owner-label');
statusSelect.append(...TASK_STATUSES.map(status => new Option(status)));

const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, tabStripVisible: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('共享任务');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.isProtected = true;

// 两个视图读取同一个数组。视图选项各自保存。
const records = createTasks();
const views = {
  table: { status: '', showOwner: true },
  board: { status: '' }
};
let activeView = 'table';
let tableRows = [];
let tableFields = [];
const labels = { id: '编号', title: '任务', status: '状态', owner: '负责人' };
const board = new BoardProvider(records, id => showTaskDetail(records.find(task => task.id === id)));

const headerStyle = new S.Style();
headerStyle.backColor = '#eaf0f5';
headerStyle.font = 'bold 14px "Microsoft YaHei", sans-serif';
headerStyle.locked = true;
const readOnlyStyle = new S.Style();
readOnlyStyle.locked = true;
const statusStyle = new S.Style();
statusStyle.locked = false;
statusStyle.cellType = new S.CellTypes.ComboBox().items(TASK_STATUSES);

// 同一块 DataRange 按当前视图选择不同的坐标映射。
const provider = {
  getValue(row, col) {
    if (activeView === 'board') return board.getValue(row, col);
    const field = tableFields[col];
    if (row === 0) return labels[field] ?? null;
    return tableRows[row - 1]?.[field] ?? null;
  },
  getStyle(row, col) {
    if (activeView === 'board') return board.getStyle(row, col);
    if (row === 0) return headerStyle;
    return tableRows[row - 1] && tableFields[col] === 'status' ? statusStyle : readOnlyStyle;
  },
  setValue(row, col, value) {
    const task = tableRows[row - 1];
    if (activeView === 'table' && task && tableFields[col] === 'status' && TASK_STATUSES.includes(value)) {
      task.status = value;
      // 等本轮编辑提交完成后再筛选，避免编辑过程中行号指向其他任务。
      queueMicrotask(() => { if (activeView === 'table') render(); });
    }
    return true;
  },
  onMouseDown(row, col, event) {
    return activeView === 'board' && board.onMouseDown(row, col, event);
  }
};
const dataRange = sheet.dataRanges.add('shared-task-view', provider, new S.Range(0, 0, 1, 1));

function render() {
  if (activeView === 'board') {
    board.rebuild(views.board.status);
    renderBoard(sheet, dataRange, board);
    return;
  }
  const { status, showOwner } = views.table;
  tableRows = records.filter(task => !status || task.status === status);
  tableFields = showOwner ? ['id', 'title', 'status', 'owner'] : ['id', 'title', 'status'];

  sheet.suspendPaint();
  try {
    sheet.setRowCount(Math.max(8, tableRows.length + 1));
    sheet.setColumnCount(5);
    sheet.options.gridline = { showHorizontalGridline: true, showVerticalGridline: true };
    sheet.options.selectionBorderColor = '#3972a6';
    sheet.options.selectionBackColor = 'rgba(57, 114, 166, 0.12)';
    [68, 210, 118, 105, 64].forEach((width, col) => sheet.setColumnWidth(col, width));
    for (let row = 0; row < sheet.getRowCount(); row++) sheet.setRowHeight(row, 32);
    dataRange.range(new S.Range(0, 0, Math.max(2, tableRows.length + 1), tableFields.length));
    dataRange.repaint();
  } finally {
    sheet.resumePaint();
  }
}

function syncControls() {
  viewSelect.value = activeView;
  statusSelect.value = views[activeView].status;
  ownerToggle.checked = views.table.showOwner;
  ownerLabel.hidden = activeView !== 'table';
}

function changeViewSettings(update) {
  if (!sheet.endEdit()) {
    syncControls();
    return;
  }
  update();
  syncControls();
  render();
}

viewSelect.onchange = () => changeViewSettings(() => { activeView = viewSelect.value; });
statusSelect.onchange = () => changeViewSettings(() => { views[activeView].status = statusSelect.value; });
ownerToggle.onchange = () => changeViewSettings(() => { views.table.showOwner = ownerToggle.checked; });
document.getElementById('inspect-views').onclick = () => {
  console.table(records);
  console.log('表格与看板的独立设置', structuredClone(views));
};

syncControls();
render();
window.demo = { spread, sheet, records, provider, dataRange, views, board, get activeView() { return activeView; } };
