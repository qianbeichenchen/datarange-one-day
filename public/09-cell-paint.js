import { createTasks, TASK_STATUS_COLORS } from './task-data.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <label>绘制示例 <select id="paint-mode">
    <option value="status">绘制状态标签</option>
    <option value="cards">绘制整张卡片</option>
  </select></label>
  <button id="inspect-paint">打印单元格数据</button>
  <dialog id="paint-details" aria-labelledby="paint-title">
    <form method="dialog">
      <h2 id="paint-title"></h2>
      <p id="paint-summary"></p>
      <button>关闭详情</button>
    </form>
  </dialog>`;

const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('09 自定义绘制');
sheet.options.isProtected = true;
sheet.options.gridline = { showHorizontalGridline: false, showVerticalGridline: false };

const records = createTasks();
const origin = { row: 2, col: 1 };
const cardsPerRow = 2;
let mode = 'status';

function paintStatus(ctx, value, x, y, maxWidth) {
  const { background, foreground } = TASK_STATUS_COLORS[value] ?? { background: '#eaf0f5', foreground: '#35516d' };
  ctx.font = '12px Microsoft YaHei';
  const width = Math.min(ctx.measureText(value).width + 20, maxWidth);
  ctx.beginPath();
  ctx.roundRect(x, y, width, 22, 11);
  ctx.fillStyle = background;
  ctx.fill();
  ctx.fillStyle = foreground;
  ctx.textBaseline = 'middle';
  ctx.fillText(value, x + 10, y + 11, width - 16);
}

class StatusCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, width, height) {
    if (!value) return;
    ctx.save();
    paintStatus(ctx, value, x + 12, y + (height - 22) / 2, width - 24);
    ctx.restore();
  }
}

class TaskCardCell extends S.CellTypes.Base {
  paint(ctx, task, x, y, width, height) {
    if (!task) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, width, height);
    ctx.clip();

    ctx.beginPath();
    ctx.roundRect(x + 6, y + 6, width - 12, height - 12, 8);
    ctx.fillStyle = '#ffffff';
    ctx.fill();
    ctx.strokeStyle = '#c8d9e4';
    ctx.stroke();

    ctx.textBaseline = 'middle';
    ctx.font = 'bold 15px Microsoft YaHei';
    ctx.fillStyle = '#243346';
    ctx.fillText(task.title, x + 18, y + 27, width - 36);
    paintStatus(ctx, task.status, x + 18, y + 46, width - 36);
    ctx.fillStyle = '#35516d';
    ctx.fillText(`${task.owner} · ${task.due}`, x + 18, y + 85, width - 36);
    ctx.fillStyle = '#245b75';
    ctx.fillText('查看详情', x + 18, y + 108, width - 36);
    ctx.restore();
  }
  format(task) { return task?.title ?? ''; }
}

const statusCell = new StatusCell();
const cardCell = new TaskCardCell();
const details = document.getElementById('paint-details');

function openCard(row, col) {
  if (mode !== 'cards') return false;
  const task = records[row * cardsPerRow + col];
  if (!task) return false;
  document.getElementById('paint-title').textContent = task.title;
  document.getElementById('paint-summary').textContent = `${task.id} · ${task.status} · ${task.owner} · 截止 ${task.due}`;
  details.showModal();
  return true;
}

const provider = {
  getValue(row, col) {
    // 状态示例返回一个字符串，卡片示例返回一条完整记录。
    if (mode === 'cards') return records[row * cardsPerRow + col] ?? null;
    if (row === 0) return ['任务', '状态'][col] ?? null;
    return records[row - 1]?.[col === 0 ? 'title' : 'status'] ?? null;
  },
  getStyle(row, col) {
    const style = new S.Style();
    style.locked = true;
    style.vAlign = S.VerticalAlign.center;
    if (mode === 'cards') {
      style.backColor = '#f2f6fa';
      style.cellType = cardCell;
    } else if (row === 0) {
      style.backColor = '#245b75';
      style.foreColor = '#ffffff';
    } else {
      style.backColor = row % 2 === 0 ? '#f3f7fa' : '#ffffff';
      if (col === 1) style.cellType = statusCell;
    }
    return style;
  },
  onMouseDown(row, col, event) {
    if (event.button !== 0 || mode !== 'cards') return false;
    sheet.setActiveCell(origin.row + row, origin.col + col);
    return openCard(row, col);
  },
  onKeyDown(row, col, event) {
    if (event.key !== ' ' || !openCard(row, col)) return false;
    event.preventDefault();
    return true;
  }
};

const dataRange = sheet.dataRanges.add('task-paint', provider,
  new S.Range(origin.row, origin.col, records.length + 1, cardsPerRow));

function applyMode() {
  const cards = mode === 'cards';
  const rowCount = cards ? Math.ceil(records.length / cardsPerRow) : records.length + 1;
  spread.suspendPaint();
  try {
    sheet.setColumnWidth(origin.col, cards ? 240 : 225);
    sheet.setColumnWidth(origin.col + 1, cards ? 240 : 120);
    for (let row = 0; row <= records.length; row++) {
      sheet.setRowHeight(origin.row + row, cards && row < rowCount ? 130 : 36);
    }
    dataRange.range(new S.Range(origin.row, origin.col, rowCount, cardsPerRow));
    sheet.setActiveCell(origin.row, origin.col);
    dataRange.repaint();
  } finally {
    spread.resumePaint();
  }
}

const modeSelect = document.getElementById('paint-mode');
modeSelect.onchange = () => { mode = modeSelect.value; applyMode(); };
document.getElementById('inspect-paint').onclick = () => {
  console.log(mode === 'status' ? '状态格返回字符串' : '卡片格返回完整对象',
    mode === 'status' ? provider.getValue(1, 1) : provider.getValue(0, 0));
};

applyMode();
window.demo = { spread, sheet, dataRange, provider, records, get mode() { return mode; } };
console.log('第 09 课。先看状态标签，再切换到整张卡片。点击卡片，或选中后按空格，查看详情。');
