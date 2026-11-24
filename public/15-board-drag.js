import { createTasks } from './demo-data.js';
import { BoardProvider, renderBoard, showTaskDetail } from './board-layout.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = '<button id="inspect-board-drag" type="button">打印任务状态</button>';
const host = document.getElementById('ss');
const spread = new S.Workbook(host, {
  sheetCount: 1, tabStripVisible: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('看板拖拽');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.isProtected = true;

const records = createTasks();
let drag = null;
let ghost = null;
const provider = new BoardProvider(records, (id, col, event) => {
  // 按下时只记任务 ID、来源列和指针起点。松开后才提交修改。
  drag = { id, sourceColumn: col, x: event.clientX, y: event.clientY, moved: false };
});
const dataRange = sheet.dataRanges.add('drag-board', provider, new S.Range(0, 0, 1, 1));

function targetColumnAt(event) {
  const bounds = host.getBoundingClientRect();
  const x = event.clientX - bounds.left;
  const y = event.clientY - bounds.top;
  if (x < 0 || x >= bounds.width || y < 0 || y >= bounds.height) return -1;
  const hit = sheet.hitTest(x, y);
  if (hit?.hitTestType !== S.SheetArea.viewport) return -1;

  // hitTest 返回工作表坐标，分组数组使用 DataRange 内部坐标。
  const range = dataRange.range();
  const row = hit.row - range.row;
  const col = hit.col - range.col;
  return row >= 0 && row < range.rowCount && col >= 0 && col < range.colCount ? col : -1;
}

function clearDrag() {
  drag = null;
  ghost?.remove();
  ghost = null;
  provider.targetColumn = -1;
  dataRange.repaint();
}

document.addEventListener('mousemove', event => {
  if (!drag) return;
  if (!(event.buttons & 1)) {
    clearDrag();
    return;
  }
  if (!drag.moved) {
    if (Math.hypot(event.clientX - drag.x, event.clientY - drag.y) < 6) return;
    drag.moved = true;
    ghost = document.createElement('span');
    ghost.textContent = records.find(task => task.id === drag.id).title;
    ghost.setAttribute('aria-hidden', 'true');
    Object.assign(ghost.style, {
      position: 'fixed', zIndex: '10', pointerEvents: 'none', padding: '7px 10px',
      border: '1px solid #b6cadf', borderRadius: '5px', background: '#ffffff', color: '#243346'
    });
    document.body.append(ghost);
  }
  ghost.style.left = `${event.clientX + 14}px`;
  ghost.style.top = `${event.clientY + 12}px`;
  const hitColumn = targetColumnAt(event);
  const destination = hitColumn === drag.sourceColumn ? -1 : hitColumn;
  if (destination !== provider.targetColumn) {
    provider.targetColumn = destination;
    dataRange.repaint();
  }
}, true);

document.addEventListener('mouseup', event => {
  if (!drag || event.button !== 0) return;
  const finished = drag;
  const targetColumn = targetColumnAt(event);
  clearDrag();
  const task = records.find(item => item.id === finished.id);
  if (!finished.moved) {
    if (targetColumn === finished.sourceColumn) showTaskDetail(task);
    return;
  }
  if (targetColumn < 0 || targetColumn === finished.sourceColumn) return;

  // 目标列表示业务状态。写回原记录后重新分组，卡片自然出现在新列。
  const previousStatus = task.status;
  task.status = provider.groups[targetColumn].status;
  provider.rebuild();
  renderBoard(sheet, dataRange, provider);
  console.log('看板拖拽', { id: task.id, from: previousStatus, to: task.status });
}, true);

document.addEventListener('keydown', event => {
  if (drag && event.key === 'Escape') {
    event.preventDefault();
    clearDrag();
  }
}, true);
window.addEventListener('blur', () => { if (drag) clearDrag(); });

document.getElementById('inspect-board-drag').onclick = () => console.table(records);
renderBoard(sheet, dataRange, provider);
window.demo = { spread, sheet, records, provider, dataRange };
