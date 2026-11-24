import { TASK_STATUSES, TASK_STATUS_COLORS } from './task-data.js';

const S = GC.Spread.Sheets;

class BoardCell extends S.CellTypes.Base {
  paint(ctx, value, x, y, width, height, style) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, width, height);
    ctx.clip();
    ctx.fillStyle = style.backColor;
    ctx.fillRect(x, y, width - 2, height);

    if (value?.kind === 'heading') {
      ctx.fillStyle = TASK_STATUS_COLORS[value.status].foreground;
      ctx.font = 'bold 14px "Microsoft YaHei", sans-serif';
      ctx.textBaseline = 'middle';
      ctx.fillText(`${value.status}  ·  ${value.count}`, x + 14, y + height / 2, width - 28);
    } else if (value) {
      // value 是一整条任务记录。任务的几个字段都画进同一个单元格。
      ctx.beginPath();
      ctx.roundRect(x + 8, y + 5, width - 18, height - 10, 8);
      ctx.fillStyle = '#ffffff';
      ctx.strokeStyle = '#dce3eb';
      ctx.fill();
      ctx.stroke();

      ctx.textBaseline = 'middle';
      ctx.fillStyle = '#243346';
      ctx.font = 'bold 14px "Microsoft YaHei", sans-serif';
      ctx.fillText(value.title, x + 19, y + 25, width - 38);
      ctx.fillStyle = '#53677a';
      ctx.font = '12px "Microsoft YaHei", sans-serif';
      ctx.fillText(`${value.id}  ·  ${value.owner || '未指派'}`, x + 19, y + 48, width - 38);
      ctx.fillText(value.due ? `到期 ${value.due}` : '未设到期日', x + 19, y + 65, width - 38);
    }
    ctx.restore();
  }

  format(value) {
    return value?.title ?? value?.status ?? '';
  }
}

export class BoardProvider {
  constructor(records, onCardMouseDown) {
    this.records = records;
    this.onCardMouseDown = onCardMouseDown;
    this.cellType = new BoardCell();
    this.targetColumn = -1;
    this.styles = new Map();
    this.rebuild();
  }

  rebuild(statusFilter = '') {
    // 列次序由状态定义决定。即使某状态没有任务，也保留一列。
    this.groups = TASK_STATUSES.filter(status => !statusFilter || status === statusFilter).map(status => ({
      status,
      tasks: this.records.filter(task => task.status === status)
    }));
    this.colCount = this.groups.length;
    this.rowCount = 2 + Math.max(0, ...this.groups.map(group => group.tasks.length));
  }

  taskAt(row, col) {
    // 第 0 行是列标题，第一张卡片在第 1 行。
    return this.groups[col]?.tasks[row - 1] ?? null;
  }

  getValue(row, col) {
    const group = this.groups[col];
    if (row === 0 && group) return { kind: 'heading', status: group.status, count: group.tasks.length };
    return this.taskAt(row, col);
  }

  getStyle(row, col) {
    const group = this.groups[col];
    if (!group) return undefined;
    const background = col === this.targetColumn
      ? '#dcebf9'
      : row === 0 ? TASK_STATUS_COLORS[group.status].background : '#f4f7fa';
    if (!this.styles.has(background)) {
      const style = new S.Style();
      style.locked = true;
      style.backColor = background;
      style.cellType = this.cellType;
      this.styles.set(background, style);
    }
    return this.styles.get(background);
  }

  onMouseDown(row, col, event) {
    if (event.button !== 0 || !this.onCardMouseDown) return false;
    const task = this.taskAt(row, col);
    if (!task) return false;
    this.onCardMouseDown(task.id, col, event);
    return true;
  }
}

export function renderBoard(sheet, dataRange, provider) {
  sheet.suspendPaint();
  try {
    // 此处的 DataRange 从工作表 (0, 0) 开始。
    sheet.setRowCount(Math.max(8, provider.rowCount));
    sheet.setColumnCount(Math.max(4, provider.colCount));
    for (let col = 0; col < sheet.getColumnCount(); col++) {
      sheet.setColumnWidth(col, col < provider.colCount ? 222 : 64);
    }
    for (let row = 0; row < sheet.getRowCount(); row++) {
      sheet.setRowHeight(row, row === 0 ? 36 : row < provider.rowCount ? 84 : 30);
    }
    sheet.options.gridline = { showHorizontalGridline: false, showVerticalGridline: false };
    sheet.options.selectionBorderColor = 'transparent';
    sheet.options.selectionBackColor = 'transparent';
    dataRange.range(new S.Range(0, 0, provider.rowCount, provider.colCount));
    dataRange.repaint();
  } finally {
    sheet.resumePaint();
  }
}

let dialog;

export function showTaskDetail(task) {
  if (!task) return;
  if (!dialog) {
    dialog = document.createElement('dialog');
    dialog.setAttribute('aria-label', '任务详情');
    dialog.innerHTML = '<h2></h2><p></p><p></p><form method="dialog"><button autofocus>关闭</button></form>';
    document.body.append(dialog);
  }
  dialog.querySelector('h2').textContent = task.title;
  const [summary, due] = dialog.querySelectorAll('p');
  summary.textContent = `${task.id} · ${task.status} · ${task.owner || '未指派'}`;
  due.textContent = task.due ? `到期日 ${task.due}` : '未设到期日';
  if (!dialog.open) dialog.showModal();
}
