import { createOrders, amount, numericInput, CARD_ROWS, CARD_COLS, cardCellAt } from './business-model.js';

const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: true, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('⑦ 商品卡片');
sheet.options.isProtected = true;
sheet.options.gridline = { showHorizontalGridline: false, showVerticalGridline: false };

const records = createOrders();
const byId = new Map(records.map(item => [item.id, item]));
const undoManager = spread.undoManager();
let nextId = 9;
let perRow = 3;
let category = '';
let keyword = '';
let visible = records.slice();
let spans = [];

// 这里的 row、col 已是区域内部坐标；不要再次减去工作表起点。
function resolveCell(row, col) {
  const cell = cardCellAt(row, col, perRow, visible.length);
  return cell ? { ...cell, item: visible[cell.index] } : null;
}

class CardProvider {
  getValue(row, col) {
    if (!visible.length) return row === 0 && col === 0 ? '没有符合条件的商品' : null;
    const cell = resolveCell(row, col);
    if (!cell) return null; // 间距 / 最后一排空位，禁止回退到普通工作表存储。
    const { item, localRow, localCol } = cell;
    if (localRow === 0) return localCol === 0 ? item.name : null;
    if (localRow === 1) return localCol === 0 ? `${item.id} · ${item.region} / ${item.category}` : null;
    if (localCol === 0) return { 2: '单价 / 元', 3: '数量', 4: '金额 / 元' }[localRow] ?? null;
    return localRow === 2 ? item.price : localRow === 3 ? item.qty : amount(item);
  }
  getStyle(row, col) {
    const style = new S.Style();
    style.vAlign = S.VerticalAlign.center;
    style.locked = true;
    style.backColor = '#ffffff';
    const cell = resolveCell(row, col);
    if (!cell) return style;
    const { localRow, localCol, field } = cell;
    style.locked = !field;
    style.backColor = '#f2f6fa';
    style.foreColor = '#294257';
    style.cellPadding = '0 10 0 10';
    if (localRow === 0) {
      style.backColor = '#245b75';
      style.foreColor = '#ffffff';
      style.font = 'bold 15px Microsoft YaHei';
    } else if (localRow === 1) {
      style.foreColor = '#748497';
      style.font = '11px Microsoft YaHei';
    } else if (localCol === 1) {
      style.hAlign = S.HorizontalAlign.right;
      style.formatter = localRow === 3 ? '0' : '0.00';
      if (localRow === 4) {
        style.backColor = '#e0eee6';
        style.foreColor = '#24634d';
        style.font = 'bold 14px Microsoft YaHei';
      }
    }
    return style;
  }
  getSpans(row, col, rowCount, colCount) {
    return spans.filter(span => span.intersect(row, col, rowCount, colCount));
  }
  setValue(row, col, value, changes) {
    const cell = resolveCell(row, col);
    if (!cell?.field) return true;
    const { item, field } = cell;
    const next = field === 'name' ? String(value ?? '').trim() : numericInput(value, { integer: field === 'qty' });
    if (next === null || next === '') { console.warn('名称不能为空；单价为非负数，数量为非负整数。'); return true; }
    if (Object.is(item[field], next)) return true;
    if (changes) changes.push({ type: 'value', id: item.id, field, oldValue: item[field] });
    item[field] = next;
    console.log('卡片写回：', { row, col, cardIndex: cell.index, id: item.id, field, next });
    queueMicrotask(refreshValues);
    return true;
  }
  undo(change) {
    if (change.type !== 'value') return;
    const item = byId.get(change.id);
    if (!item) return;
    item[change.field] = change.oldValue;
    console.log('卡片 undo：', change);
    queueMicrotask(refreshValues);
  }
}

const provider = new CardProvider();
const dataRange = sheet.dataRanges.add('product-cards', provider, new S.Range(2, 1, 1, 1));

function refreshValues() {
  dataRange.repaint();
  document.getElementById('undo-card').disabled = !undoManager.canUndo();
  document.getElementById('redo-card').disabled = !undoManager.canRedo();
}

function rebuildLayout() {
  visible = records.filter(item => (!category || item.category === category) && item.name.includes(keyword));
  const origin = dataRange.range();
  const rowCount = Math.max(1, Math.ceil(visible.length / perRow) * CARD_ROWS - 1);
  const colCount = perRow * CARD_COLS - 1;
  // 合并与卡片数量有关，在布局变化时构建，不在每次 getValue / getStyle 中生成。
  spans = visible.flatMap((item, index) => {
    const row = Math.floor(index / perRow) * CARD_ROWS;
    const col = (index % perRow) * CARD_COLS;
    return [new S.Range(row, col, 1, 2), new S.Range(row + 1, col, 1, 2)];
  });
  if (!visible.length) spans = [new S.Range(0, 0, 1, colCount)];
  sheet.setRowCount(Math.max(100, origin.row + rowCount + 10));
  sheet.setColumnCount(Math.max(20, origin.col + colCount + 2));
  // 高度、宽度属于工作表坐标；换位置时重新应用，不会改变提供者内部坐标。
  for (let row = 0; row < rowCount; row++) {
    const localRow = row % CARD_ROWS;
    sheet.setRowHeight(origin.row + row, localRow === 0 ? 40 : localRow === 1 ? 28 : localRow === 5 ? 18 : 32);
  }
  for (let col = 0; col < colCount; col++) sheet.setColumnWidth(origin.col + col, [90, 140, 22][col % CARD_COLS]);
  dataRange.range(new S.Range(origin.row, origin.col, rowCount, colCount));
  refreshValues();
}

document.getElementById('apply-cards').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.clear();
  perRow = Number(document.getElementById('cards-per-row').value);
  category = document.getElementById('card-category').value;
  keyword = document.getElementById('card-keyword').value.trim();
  rebuildLayout();
};
document.getElementById('add-card').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.clear();
  const item = { id: `O${String(nextId++).padStart(3, '0')}`, name: '新商品', region: '华东', category: category || '配件', price: 100, qty: 1 };
  records.push(item);
  byId.set(item.id, item);
  // 新商品沿用当前已应用的品类；若不匹配关键词，保留筛选并在控制台解释。
  if (keyword && !item.name.includes(keyword)) console.log('新商品已添加，但不符合当前关键词；清空关键词后可见。', item);
  rebuildLayout();
};
document.getElementById('undo-card').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.undo();
  refreshValues();
};
document.getElementById('redo-card').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.redo();
  refreshValues();
};
document.getElementById('move-cards').onclick = () => {
  if (!sheet.endEdit()) return;
  undoManager.clear();
  const old = dataRange.range();
  dataRange.range(new S.Range(old.row === 2 ? 4 : 2, old.col, old.rowCount, old.colCount));
  rebuildLayout();
};
document.getElementById('inspect-cards').onclick = () => {
  console.table(records);
  console.table(visible.map((item, index) => ({ id: item.id, index, cardRow: Math.floor(index / perRow), cardCol: index % perRow, row: Math.floor(index / perRow) * CARD_ROWS, col: (index % perRow) * CARD_COLS })));
  console.log('当前范围：', dataRange.range(), '原生历史：', undoManager.getUndoStack());
};
rebuildLayout();
window.demo = { spread, sheet, records, provider, dataRange, undoManager, resolveCell, get visible() { return visible; }, get perRow() { return perRow; } };
console.log('第七课：一张卡片占 5×2 内容格，步长为 6×3；名称、单价、数量可编辑，元数据与金额只读。');
