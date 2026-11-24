import { createTasks, TASK_MONTH, TASK_STATUS_COLORS } from './demo-data.js';
import { addDays, daysBetween } from './task-dates.js';

const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <label>练习阶段 <select id="gantt-stage">
    <option value="view">1. 查看任务条</option>
    <option value="move">2. 整体移动</option>
    <option value="resize">3. 移动或拉伸端点</option>
  </select></label>
  <button type="button" id="inspect-gantt">打印工期与映射</button>
`;
const stage = document.getElementById('gantt-stage');

const S = GC.Spread.Sheets;
const host = document.getElementById('ss');
const spread = new S.Workbook(host, { sheetCount: 1, tabStripVisible: false, allowUserDragDrop: false, allowUserDragFill: false });
const sheet = spread.getActiveSheet();
sheet.name('甘特时间轴');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.isProtected = true;

const records = createTasks();
const firstDay = addDays(`${TASK_MONTH}-01`, -2);
const dayCount = 35;
const dates = Array.from({ length: dayCount }, (_, offset) => addDays(firstDay, offset));
const dayIndex = new Map(dates.map((date, index) => [date, index]));
const rowCount = records.length + 1;
const colCount = 1 + dayCount;
const BAR_HEIGHT = 18;
const EDGE_WIDTH = 8;
let drag;
let dropCol = -1;

sheet.setRowCount(rowCount);
sheet.setColumnCount(colCount);
sheet.frozenRowCount(1);
sheet.frozenColumnCount(1);
sheet.setColumnWidth(0, 158);
sheet.setRowHeight(0, 36);
for (let col = 1; col < colCount; col++) sheet.setColumnWidth(col, 48);
for (let row = 1; row < rowCount; row++) sheet.setRowHeight(row, 45);

// 第一列是任务名称，日期列从 1 开始，一列代表一天。
function dateAtColumn(col) {
  return dates[col - 1];
}

function columnOfDate(date) {
  const index = dayIndex.get(date);
  return index === undefined ? -1 : index + 1;
}

class TimelineDayCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, w, h) {
    if (!ctx || !value) return;
    ctx.save();
    const weekday = new Date(`${value.date}T00:00:00Z`).getUTCDay();
    ctx.fillStyle = value.isTarget ? '#dcefe6' : [0, 6].includes(weekday) ? '#f2f5f8' : '#fff';
    ctx.fillRect(x, y, w, h);
    if (value.inBar) {
      const top = y + (h - BAR_HEIGHT) / 2;
      const { background, foreground } = TASK_STATUS_COLORS[value.status];
      ctx.fillStyle = background;
      ctx.fillRect(x, top, w, BAR_HEIGHT);
      ctx.fillStyle = foreground;
      ctx.fillRect(x, top + BAR_HEIGHT - 2, w, 2);
      if (stage.value === 'resize') {
        const handle = 4;
        if (value.isStart) ctx.fillRect(x + 1, top, handle, BAR_HEIGHT);
        if (value.isEnd) ctx.fillRect(x + w - handle - 1, top, handle, BAR_HEIGHT);
      }
    }
    ctx.restore();
  }

  format(value) { return value?.date ?? ''; }
}

const headerStyle = new S.Style();
headerStyle.backColor = '#eaf0f5';
headerStyle.foreColor = '#35516d';
headerStyle.font = 'bold 13px Microsoft YaHei';
headerStyle.hAlign = S.HorizontalAlign.center;
headerStyle.vAlign = S.VerticalAlign.center;
const titleStyle = new S.Style();
titleStyle.backColor = '#f7faf9';
titleStyle.foreColor = '#243346';
titleStyle.vAlign = S.VerticalAlign.center;
titleStyle.cellPadding = '0 6 0 8';
titleStyle.borderBottom = new S.LineBorder('#dce3eb', S.LineStyle.thin);
const dateStyle = new S.Style();
dateStyle.cellType = new TimelineDayCell();
dateStyle.borderBottom = new S.LineBorder('#dce3eb', S.LineStyle.thin);
dateStyle.borderRight = new S.LineBorder('#edf0f5', S.LineStyle.thin);

const provider = {
  getValue(row, col) {
    if (row < 0 || row >= rowCount || col < 0 || col >= colCount) return null;
    if (row === 0) return col === 0 ? '任务' : dateAtColumn(col).slice(5);
    const task = records[row - 1];
    if (col === 0) return task.title;
    const date = dateAtColumn(col);
    return {
      date,
      status: task.status,
      inBar: task.start <= date && date <= task.due,
      isStart: date === task.start,
      isEnd: date === task.due,
      isTarget: drag?.row === row && dropCol === col
    };
  },
  getStyle(row, col) {
    if (row === 0) return headerStyle;
    return col === 0 ? titleStyle : dateStyle;
  },
  onMouseDown(row, col, event) {
    if (event.button !== 0 || row < 1 || col < 1) return false;
    const task = records[row - 1];
    const cell = this.getValue(row, col);
    if (!cell.inBar) return false;
    if (stage.value === 'view') {
      console.log('任务条', { id: task.id, start: task.start, due: task.due, duration: daysBetween(task.start, task.due) + 1 });
      return true;
    }
    const rect = sheet.getCellRect(row, col);
    const bounds = host.getBoundingClientRect();
    const x = event.clientX - bounds.left - rect.x;
    const y = event.clientY - bounds.top - rect.y;
    if (Math.abs(y - rect.height / 2) > BAR_HEIGHT / 2) return false;

    // 按下时确定整条移动还是只改一端，整个拖动过程都沿用这个模式。
    let mode = 'move';
    if (stage.value === 'resize') {
      if (cell.isStart && x < EDGE_WIDTH) mode = 'start';
      if (cell.isEnd && x > rect.width - EDGE_WIDTH) mode = 'end';
    }
    drag = { id: task.id, row, col, mode, start: task.start, due: task.due, x: event.clientX };
    return true;
  }
};

const dataRange = sheet.dataRanges.add('task-gantt', provider, new S.Range(0, 0, rowCount, colCount));

function targetColumn(event) {
  const bounds = host.getBoundingClientRect();
  const x = event.clientX - bounds.left;
  const y = event.clientY - bounds.top;
  if (x < 0 || x >= bounds.width || y < 0 || y >= bounds.height) return -1;
  const hit = sheet.hitTest(x, y);
  return hit?.hitTestType === S.SheetArea.viewport && hit.row === drag.row && dateAtColumn(hit.col) ? hit.col : -1;
}

function movedEnough(event) {
  return Math.abs(event.clientX - drag.x) >= 5;
}

function cancelDrag() {
  drag = undefined;
  dropCol = -1;
  host.style.cursor = '';
  dataRange.repaint();
}

document.addEventListener('mousemove', event => {
  if (!drag) return;
  if (!(event.buttons & 1)) { cancelDrag(); return; }
  if (!movedEnough(event)) return;
  const col = targetColumn(event);
  host.style.cursor = col < 0 ? 'not-allowed' : drag.mode === 'move' ? 'grabbing' : 'ew-resize';
  if (col === dropCol) return;
  dropCol = col;
  dataRange.repaint();
});

document.addEventListener('mouseup', event => {
  if (!drag || event.button !== 0) return;
  const col = movedEnough(event) ? targetColumn(event) : -1;
  if (col > 0 && col !== drag.col) {
    // 一列就是一天。整体移动时起止日期加同一个差值，因此工期保持不变。
    const delta = col - drag.col;
    const start = drag.mode === 'end' ? drag.start : addDays(drag.start, delta);
    const due = drag.mode === 'start' ? drag.due : addDays(drag.due, delta);
    if (due < start) {
      console.warn('甘特拖动未提交：结束日期不能早于开始日期。', { start, due });
    } else if (columnOfDate(start) === -1 || columnOfDate(due) === -1) {
      console.warn('甘特拖动未提交：任务条超出练习时间轴。', { start, due });
    } else {
      const task = records.find(item => item.id === drag.id);
      task.start = start;
      task.due = due;
      console.log('甘特任务更新', { id: task.id, mode: drag.mode, start, due, duration: daysBetween(start, due) + 1 });
    }
  }
  cancelDrag();
});

document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && drag) cancelDrag();
});
window.addEventListener('blur', cancelDrag);
stage.onchange = cancelDrag;
document.getElementById('inspect-gantt').onclick = () => {
  console.table(records.map(({ id, title, start, due }) => ({ id, title, start, due, startColumn: columnOfDate(start), endColumn: columnOfDate(due), days: daysBetween(start, due) + 1 })));
};
window.demo = { spread, sheet, records, provider, dataRange, dateAtColumn, columnOfDate };
