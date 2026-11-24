import { createTasks, TASK_MONTH, TASK_STATUS_COLORS } from './demo-data.js';
import { addMonths, monthGrid } from './task-dates.js';

const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <button type="button" id="calendar-prev">上个月</button>
  <output id="calendar-month" aria-live="polite"></output>
  <button type="button" id="calendar-next">下个月</button>
  <label><input type="checkbox" id="calendar-drag">启用拖动改期</label>
  <button type="button" id="inspect-calendar">打印日期映射</button>
`;
const monthOutput = document.getElementById('calendar-month');
const dragEnabled = document.getElementById('calendar-drag');
dragEnabled.style.width = 'auto';

const S = GC.Spread.Sheets;
const host = document.getElementById('ss');
const spread = new S.Workbook(host, { sheetCount: 1, tabStripVisible: false, allowUserDragDrop: false, allowUserDragFill: false });
const sheet = spread.getActiveSheet();
sheet.name('月历');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.isProtected = true;
sheet.options.gridline = { showHorizontalGridline: false, showVerticalGridline: false };
sheet.setRowCount(7);
sheet.setColumnCount(7);
sheet.setRowHeight(0, 34);
sheet.frozenRowCount(1);
for (let row = 1; row < 7; row++) sheet.setRowHeight(row, 110);

const weekdays = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];
const records = createTasks();
let month = TASK_MONTH;
let days;
let tasksByDate;
let drag;
let dropDate;
const FIRST_TASK_TOP = 27;
const TASK_STEP = 20;
const TASK_LIMIT = 3;

function rebuildMonth() {
  days = monthGrid(month);
  tasksByDate = new Map();
  for (const task of records) {
    if (!tasksByDate.has(task.due)) tasksByDate.set(task.due, []);
    tasksByDate.get(task.due).push(task);
  }
  monthOutput.textContent = `${month.slice(0, 4)} 年 ${Number(month.slice(5))} 月`;
}

// 第一行是星期标题。其余位置由周序号、星期序号换算成一维日期索引。
function dateAt(row, col) {
  if (row < 1 || row > 6 || col < 0 || col > 6) return null;
  return days[(row - 1) * 7 + col];
}

class CalendarDayCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, w, h) {
    if (!ctx || !value) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    ctx.fillStyle = value.date === dropDate ? '#dcefe6' : value.inMonth ? '#fff' : '#f2f5f8';
    ctx.fillRect(x, y, w, h);
    ctx.fillStyle = value.inMonth ? '#243346' : '#687789';
    ctx.font = 'bold 13px Microsoft YaHei, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(value.inMonth ? String(Number(value.date.slice(8))) : value.date.slice(5), x + 8, y + 14);

    value.tasks.slice(0, TASK_LIMIT).forEach((task, index) => {
      const top = y + FIRST_TASK_TOP + index * TASK_STEP;
      const { background, foreground } = TASK_STATUS_COLORS[task.status];
      ctx.fillStyle = background;
      ctx.fillRect(x + 5, top, w - 10, TASK_STEP - 2);
      ctx.fillStyle = foreground;
      ctx.font = '12px Microsoft YaHei, sans-serif';
      ctx.fillText(task.title, x + 10, top + (TASK_STEP - 2) / 2, w - 20);
    });
    if (value.tasks.length > TASK_LIMIT) {
      ctx.fillStyle = '#315b77';
      ctx.font = '11px Microsoft YaHei, sans-serif';
      ctx.fillText(`另有 ${value.tasks.length - TASK_LIMIT} 项`, x + 8, y + h - 12);
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
const dayStyle = new S.Style();
dayStyle.cellType = new CalendarDayCell();
dayStyle.borderRight = new S.LineBorder('#dce3eb', S.LineStyle.thin);
dayStyle.borderBottom = new S.LineBorder('#dce3eb', S.LineStyle.thin);

const provider = {
  getValue(row, col) {
    if (row === 0) return weekdays[col];
    const date = dateAt(row, col);
    return date ? { date, inMonth: date.startsWith(month), tasks: tasksByDate.get(date) ?? [] } : null;
  },
  getStyle(row) { return row === 0 ? headerStyle : dayStyle; },
  onMouseDown(row, col, event) {
    if (event.button !== 0 || row === 0) return false;
    const day = this.getValue(row, col);
    if (!day) return false;
    const localY = event.clientY - host.getBoundingClientRect().top - sheet.getCellRect(row, col).y;
    const index = Math.floor((localY - FIRST_TASK_TOP) / TASK_STEP);
    const task = index >= 0 && index < TASK_LIMIT ? day.tasks[index] : undefined;
    if (task && dragEnabled.checked) {
      drag = { id: task.id, sourceDate: day.date, x: event.clientX, y: event.clientY };
    } else if (task) {
      console.log('点击日历任务', { ...task });
    } else {
      console.log('日期格', { row, col, date: day.date });
      console.table(day.tasks);
    }
    return true;
  }
};

rebuildMonth();
const dataRange = sheet.dataRanges.add('task-calendar', provider, new S.Range(0, 0, 7, 7));

function setColumnWidths() {
  const width = Math.max(96, Math.floor((host.clientWidth - 20) / 7));
  spread.suspendPaint();
  try {
    for (let col = 0; col < 7; col++) sheet.setColumnWidth(col, width);
  } finally { spread.resumePaint(); }
}
setColumnWidths();
new ResizeObserver(setColumnWidths).observe(host);

function changeMonth(delta) {
  cancelDrag();
  month = addMonths(month, delta);
  rebuildMonth();
  dataRange.repaint();
}

function dateUnderPointer(event) {
  const bounds = host.getBoundingClientRect();
  const x = event.clientX - bounds.left;
  const y = event.clientY - bounds.top;
  if (x < 0 || x >= bounds.width || y < 0 || y >= bounds.height) return null;
  const hit = sheet.hitTest(x, y);
  return hit?.hitTestType === S.SheetArea.viewport ? dateAt(hit.row, hit.col) : null;
}

function movedEnough(event) {
  return Math.hypot(event.clientX - drag.x, event.clientY - drag.y) >= 5;
}

function cancelDrag() {
  drag = undefined;
  dropDate = undefined;
  host.style.cursor = '';
  dataRange.repaint();
}

document.addEventListener('mousemove', event => {
  if (!drag) return;
  if (!(event.buttons & 1)) { cancelDrag(); return; }
  if (!movedEnough(event)) return;
  const date = dateUnderPointer(event);
  host.style.cursor = date ? 'grabbing' : 'not-allowed';
  if (date === dropDate) return;
  dropDate = date;
  dataRange.repaint();
});

document.addEventListener('mouseup', event => {
  if (!drag || event.button !== 0) return;
  const moved = movedEnough(event);
  const date = moved ? dateUnderPointer(event) : null;
  const task = records.find(item => item.id === drag.id);
  if (date && date < task.start) {
    console.warn('日历改期未提交：截止日期不能早于开始日期。', { id: task.id, start: task.start, due: date });
  } else if (date && date !== drag.sourceDate) {
    const before = task.due;
    task.due = date;
    rebuildMonth();
    console.log('日历任务改期', { id: task.id, before, due: task.due });
  } else if (!moved) {
    console.log('点击日历任务', { ...task });
  }
  cancelDrag();
});

document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && drag) cancelDrag();
});
window.addEventListener('blur', cancelDrag);
dragEnabled.onchange = cancelDrag;
document.getElementById('calendar-prev').onclick = () => changeMonth(-1);
document.getElementById('calendar-next').onclick = () => changeMonth(1);
document.getElementById('inspect-calendar').onclick = () => {
  console.table(days.map((date, index) => ({ row: Math.floor(index / 7) + 1, col: index % 7, date, tasks: tasksByDate.get(date)?.length ?? 0 })));
  console.table(records);
};
window.demo = { spread, sheet, records, provider, dataRange, dateAt, get month() { return month; } };
