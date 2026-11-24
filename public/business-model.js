// 纯业务代码：不依赖 GC，可独立运行测试。先读对应课程的 Provider，再深入本文件。
export const roundMoney = value => Math.round((value + Number.EPSILON) * 100) / 100;
export const amount = item => roundMoney(item.price * item.qty);
export function summarize(records) {
  return {
    count: records.length,
    qty: records.reduce((sum, item) => sum + item.qty, 0),
    amount: roundMoney(records.reduce((sum, item) => sum + amount(item), 0))
  };
}

export function orderView(records, { keyword = '', region = '', sort = 'id' } = {}) {
  const needle = keyword.trim().toLocaleLowerCase();
  // filter 得到新数组；sort 不会改变原始 records 的顺序。数组元素仍是原对象引用。
  return records.filter(item => (!region || item.region === region) && item.name.toLocaleLowerCase().includes(needle))
    .sort((a, b) => {
      const delta = sort === 'amount-desc' ? amount(b) - amount(a) : sort === 'amount-asc' ? amount(a) - amount(b) : 0;
      return delta || a.id.localeCompare(b.id);
    });
}

export function numericInput(value, { integer = false, positive = false } = {}) {
  if (value == null || (typeof value !== 'number' && typeof value !== 'string') || String(value).trim() === '') return null;
  const next = Number(value);
  if (!Number.isFinite(next) || next < 0 || (positive && next === 0) || (integer && !Number.isInteger(next))) return null;
  return next;
}

// 第七课：卡片 5 行内容 + 1 行间距，2 列内容 + 1 列间距。
export const CARD_ROWS = 6;
export const CARD_COLS = 3;
export function cardCellAt(row, col, perRow, count) {
  if (!Number.isInteger(perRow) || perRow < 1) throw new Error('每行卡片数必须为正整数');
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || col >= perRow * CARD_COLS) return null;
  const localRow = row % CARD_ROWS;
  const localCol = col % CARD_COLS;
  const index = Math.floor(row / CARD_ROWS) * perRow + Math.floor(col / CARD_COLS);
  // 间距和最后一排未使用的位置都不对应业务记录。
  if (localRow === CARD_ROWS - 1 || localCol === CARD_COLS - 1 || index >= count) return null;
  const field = localRow === 0 ? 'name' : localCol === 1 && localRow === 2 ? 'price' : localCol === 1 && localRow === 3 ? 'qty' : null;
  return { index, localRow, localCol, field };
}

function bucketBy(records, field) {
  const groups = new Map();
  for (const item of records) {
    if (!groups.has(item[field])) groups.set(item[field], []);
    groups.get(item[field]).push(item);
  }
  return groups;
}

export function groupReport(records, region = '') {
  const filtered = records.filter(item => !region || item.region === region);
  // 显示行描述：合成行与真实明细分开，不往业务数组塞“小计对象”。
  const display = [{ kind: 'band' }, { kind: 'header' }];
  for (const [regionName, regionItems] of bucketBy(filtered, 'region')) {
    display.push({ kind: 'region', label: regionName, ...summarize(regionItems) });
    for (const [category, items] of bucketBy(regionItems, 'category')) {
      display.push({ kind: 'category', label: category, ...summarize(items) });
      display.push(...items.map(item => ({ kind: 'data', item })));
      display.push({ kind: 'subtotal', label: `${category}小计`, ...summarize(items) });
    }
    display.push({ kind: 'regionTotal', label: `${regionName}合计`, ...summarize(regionItems) });
  }
  display.push({ kind: 'total', label: '当前范围总计', ...summarize(filtered) });
  const categories = [...bucketBy(filtered, 'category')].map(([category, items]) => ({ category, ...summarize(items) }));
  return { display, categories, total: summarize(filtered), filtered };
}

export function treeProjection(nodes, collapsed = new Set()) {
  const byId = new Map(nodes.map(node => [node.id, node]));
  if (byId.size !== nodes.length) throw new Error('任务 ID 重复');
  const children = new Map();
  for (const node of nodes) {
    if (node.parentId != null && !byId.has(node.parentId)) throw new Error(`任务 ${node.id} 的父节点不存在`);
    if (!children.has(node.parentId)) children.set(node.parentId, []);
    children.get(node.parentId).push(node);
  }

  // 第一次遍历所有节点汇总；折叠不会影响汇总口径。
  const totals = new Map();
  const visiting = new Set();
  function rollup(node) {
    if (totals.has(node.id)) return totals.get(node.id);
    if (visiting.has(node.id)) throw new Error('任务父子关系存在循环');
    visiting.add(node.id);
    const kids = children.get(node.id) || [];
    const plan = kids.length ? kids.reduce((sum, kid) => sum + rollup(kid).plan, 0) : node.plan;
    const actual = kids.length ? kids.reduce((sum, kid) => sum + rollup(kid).actual, 0) : node.actual;
    const result = { plan, actual, variance: roundMoney(actual - plan), ratio: plan > 0 ? actual / plan : 0 };
    totals.set(node.id, result);
    visiting.delete(node.id);
    return result;
  }
  for (const node of nodes) rollup(node); // 连无根的循环也要检测，不能只遍历根节点。

  // 第二次遍历只负责可见行。折叠只隐藏后代，不删除业务节点。
  const visible = [];
  function visit(node, depth) {
    const kids = children.get(node.id) || [];
    visible.push({ node, depth, branch: kids.length > 0, ...totals.get(node.id) });
    if (!collapsed.has(node.id)) for (const kid of kids) visit(kid, depth + 1);
  }
  for (const root of children.get(null) || []) visit(root, 0);
  return { visible, totals, children };
}
