import { createTasks, TASK_STATUSES, TASK_OWNERS } from './task-data.js';
import { BoardProvider, renderBoard, showTaskDetail } from './board-layout.js';

const S = GC.Spread.Sheets;
const tools = document.getElementById('lesson-tools');
tools.innerHTML = `
  <label>新增到 <select id="board-status"></select></label>
  <button id="board-add" type="button">添加任务</button>
  <button id="board-inspect" type="button">打印数据与分组</button>
`;
const statusSelect = document.getElementById('board-status');
statusSelect.append(...TASK_STATUSES.map(status => new Option(status)));

const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, tabStripVisible: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('状态看板');
sheet.options.rowHeaderVisible = false;
sheet.options.colHeaderVisible = false;
sheet.options.isProtected = true;

const records = createTasks();
let nextId = records.length + 1;
const provider = new BoardProvider(records, id => {
  // 点击得到记录 ID，详情按 ID 读取。列和行会因新增、改组而变化。
  showTaskDetail(records.find(task => task.id === id));
});
const dataRange = sheet.dataRanges.add('status-board', provider, new S.Range(0, 0, 1, 1));

document.getElementById('board-add').onclick = () => {
  const number = String(nextId++).padStart(2, '0');
  records.push({
    id: `T${number}`, title: `新任务 ${number}`, status: statusSelect.value,
    owner: TASK_OWNERS[0], priority: 2, start: '', due: ''
  });
  provider.rebuild();
  renderBoard(sheet, dataRange, provider);
};
document.getElementById('board-inspect').onclick = () => {
  console.table(records);
  console.table(provider.groups.map((group, col) => ({
    col, status: group.status, taskIds: group.tasks.map(task => task.id).join('、')
  })));
};

renderBoard(sheet, dataRange, provider);
window.demo = { spread, sheet, records, provider, dataRange };
