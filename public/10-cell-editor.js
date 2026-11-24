import { createTasks } from './demo-data.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <button id="undo-date">撤销日期编辑</button>
  <button id="redo-date">重做日期编辑</button>
  <button id="inspect-dates">打印任务日期</button>`;

const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: true, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('10 自定义编辑器');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 34;
sheet.setColumnWidth(1, 220);
sheet.setColumnWidth(2, 180);

const records = createTasks();
const undoManager = spread.undoManager();

class DateCell extends S.CellTypes.Base {
  createEditorElement() {
    const root = document.createElement('div');
    const input = document.createElement('input');
    root.setAttribute('gcUIElement', 'gcEditingInput');
    input.setAttribute('gcUIElement', 'gcEditingInput');
    input.setAttribute('aria-label', '截止日期');
    input.type = 'date';
    input.min = '1900-01-01';
    input.max = '9999-12-31';
    input.style.cssText = 'width:100%;height:100%;padding:0 6px;border:0;border-radius:0;background:white;';
    root.append(input);
    return root;
  }
  setEditorValue(root, value) { root.firstElementChild.value = value ?? ''; }
  getEditorValue(root) { return root.firstElementChild.value; }
  focus(root) { root.firstElementChild.focus(); }
  updateEditor(root, style, cellRect) {
    root.style.width = `${cellRect.width}px`;
    root.style.height = `${cellRect.height}px`;
    return cellRect;
  }
  isReservedKey(event) {
    return event.target?.type === 'date' && event.key.startsWith('Arrow');
  }
  isEditingValueChanged(oldValue, newValue) {
    if (oldValue === newValue) return false;
    // 编辑器里的无效输入不产生原生命令。直接写入仍由 Provider 校验。
    if (typeof newValue === 'string' && isDateKey(newValue)) return true;
    console.warn('截止日期必须是有效的 YYYY-MM-DD。本次保留原值。');
    return false;
  }
}

function isDateKey(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || value < '1900-01-01') return false;
  const time = Date.parse(`${value}T00:00:00Z`);
  return Number.isFinite(time) && new Date(time).toISOString().slice(0, 10) === value;
}

const dateCell = new DateCell();
const provider = {
  getValue(row, col) {
    if (row === 0) return ['任务', '截止日期'][col] ?? null;
    return records[row - 1]?.[col === 0 ? 'title' : 'due'] ?? null;
  },
  getStyle(row, col) {
    const style = new S.Style();
    style.locked = row === 0 || col !== 1;
    style.vAlign = S.VerticalAlign.center;
    style.backColor = row === 0 ? '#245b75' : row % 2 ? '#ffffff' : '#f3f7fa';
    if (row === 0) style.foreColor = '#ffffff';
    if (!style.locked) {
      style.cellType = dateCell;
      style.formatter = '@'; // 业务值保留为 YYYY-MM-DD 字符串。
    }
    return style;
  },
  setValue(row, col, value, changes) {
    const record = records[row - 1];
    if (!record || col !== 1) return true;
    if (typeof value !== 'string' || !isDateKey(value)) {
      console.warn('截止日期必须是有效的 YYYY-MM-DD。本次保留原值。');
      queueMicrotask(refreshValues);
      return true;
    }
    if (record.due === value) return true;
    // 将旧值放入当前 SpreadJS 命令的变化记录中。
    if (changes) changes.push({ id: record.id, oldDue: record.due });
    record.due = value;
    console.log('日期写回', { id: record.id, due: record.due });
    queueMicrotask(refreshValues);
    return true;
  },
  undo({ id, oldDue }) {
    const record = records.find(task => task.id === id);
    if (record) record.due = oldDue;
    queueMicrotask(refreshValues);
  }
};

const dataRange = sheet.dataRanges.add('task-dates', provider, new S.Range(2, 1, records.length + 1, 2));
const undoButton = document.getElementById('undo-date');
const redoButton = document.getElementById('redo-date');

function refreshValues() {
  dataRange.repaint();
  undoButton.disabled = !undoManager.canUndo();
  redoButton.disabled = !undoManager.canRedo();
}

undoButton.onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.undo();
  refreshValues();
};
redoButton.onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.redo();
  refreshValues();
};
document.getElementById('inspect-dates').onclick = () => console.table(records.map(({ id, due }) => ({ id, due })));

refreshValues();
window.demo = { spread, sheet, provider, records, dataRange, undoManager };
console.log('第 10 课。双击截止日期，用日期控件编辑。通过原生撤销和重做观察业务日期回写。');
