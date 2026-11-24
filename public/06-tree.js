import { createTreeTasks } from './demo-data.js';
import { treeProjection, numericInput } from './business-model.js';

const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('⑥ 项目工时树');
sheet.options.isProtected = true;
sheet.defaults.rowHeight = 32;
[320, 100, 115, 115, 110, 120].forEach((width, col) => sheet.setColumnWidth(col + 1, width));

const nodes = createTreeTasks();
const collapsed = new Set(); // 只记录稳定节点 ID，不记录“第几行被折叠”。
let tree = treeProjection(nodes, collapsed);
const headers = ['任务（双击父任务展开/折叠）', '负责人', '计划 / 人天', '实际 / 人天', '偏差 / 人天', '工时消耗率'];

class TaskTreeProvider {
  getValue(row, col) {
    if (row === 0) return headers[col] ?? null;
    const line = tree.visible[row - 1];
    if (!line) return null;
    const marker = line.branch ? (collapsed.has(line.node.id) ? '▶ ' : '▼ ') : '· ';
    if (col === 0) return `${'    '.repeat(line.depth)}${marker}${line.node.title}`;
    return [null, line.node.owner, line.plan, line.actual, line.variance, line.ratio][col] ?? null;
  }
  setValue(row, col, value) {
    const line = tree.visible[row - 1];
    if (!line || line.branch || ![1, 2, 3].includes(col)) return true;
    const next = col === 1 ? String(value ?? '').trim() : numericInput(value, { positive: col === 2 });
    if (next === null || next === '') { console.warn('负责人不能为空；计划必须大于 0，实际必须为非负数。'); return true; }
    const field = { 1: 'owner', 2: 'plan', 3: 'actual' }[col];
    line.node[field] = next;
    console.log('更新叶子任务：', { id: line.node.id, field, next });
    queueMicrotask(refreshTree);
    return true;
  }
  getStyle(row, col) {
    const line = tree.visible[row - 1];
    const style = new S.Style();
    style.vAlign = S.VerticalAlign.center;
    style.locked = !line || line.branch || ![1, 2, 3].includes(col);
    style.backColor = row === 0 ? '#594979' : line?.branch ? '#eae5f2' : '#ffffff';
    if (row === 0) style.foreColor = '#ffffff';
    if (line?.branch) style.font = 'bold 12px Microsoft YaHei';
    if (row > 0 && col >= 2) style.formatter = col === 5 ? '0.0%' : '0.0';
    if (line && col === 4 && line.variance > 0) { style.foreColor = '#b63232'; style.backColor = '#ffeded'; }
    return style;
  }
  onDoubleClick(row, col) {
    const line = tree.visible[row - 1];
    if (!line?.branch || col !== 0) return false; // 让叶子负责人/工时继续进入默认编辑。
    const id = line.node.id;
    if (collapsed.has(id)) collapsed.delete(id); else collapsed.add(id);
    queueMicrotask(refreshTree);
    return true; // 已处理父任务双击，阻止后续默认双击动作。
  }
}

const provider = new TaskTreeProvider();
const dataRange = sheet.dataRanges.add('project-tree', provider, new S.Range(2, 1, tree.visible.length + 1, 6), { sticky: { top: 1 } });

function refreshTree() {
  tree = treeProjection(nodes, collapsed);
  const old = dataRange.range();
  dataRange.range(new S.Range(old.row, old.col, tree.visible.length + 1, 6));
  dataRange.repaint();
}

document.getElementById('tree-controls').hidden = false;
document.getElementById('collapse-tree').onclick = () => {
  if (!sheet.endEdit()) return;
  for (const [id, children] of tree.children) if (id !== null && children.length) collapsed.add(id);
  refreshTree();
};
document.getElementById('expand-tree').onclick = () => {
  if (!sheet.endEdit()) return;
  collapsed.clear();
  refreshTree();
};
document.getElementById('move-tree').onclick = () => {
  if (!sheet.endEdit()) return;
  const old = dataRange.range();
  dataRange.range(new S.Range(old.row === 2 ? 5 : 2, old.col, old.rowCount, old.colCount));
  dataRange.repaint();
};
document.getElementById('inspect-tree').onclick = () => {
  console.table(nodes);
  console.table(tree.visible.map((line, index) => ({ row: index + 1, id: line.node.id, depth: line.depth, plan: line.plan, actual: line.actual })));
  console.log('折叠节点 ID：', [...collapsed]);
};
window.demo = { spread, sheet, nodes, collapsed, provider, dataRange, get tree() { return tree; } };
console.log('第六课：双击父任务标题展开/折叠；只编辑叶子任务负责人和工时。工时消耗率=实际/计划，不代表完成进度。');
