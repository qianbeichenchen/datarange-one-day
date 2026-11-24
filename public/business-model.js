// 纯业务代码：不依赖 GC，可独立运行测试。先读对应课程的 Provider，再深入本文件。
export function createOrders() {
  return [
    { id: 'O001', name: '键盘', region: '华东', category: '配件', price: 199, qty: 2 },
    { id: 'O002', name: '鼠标', region: '华东', category: '配件', price: 89, qty: 3 },
    { id: 'O003', name: '显示器', region: '华东', category: '设备', price: 1299, qty: 1 },
    { id: 'O004', name: '键盘', region: '华南', category: '配件', price: 219, qty: 4 },
    { id: 'O005', name: '笔记本', region: '华南', category: '设备', price: 4999, qty: 2 },
    { id: 'O006', name: '扩展坞', region: '华北', category: '配件', price: 299, qty: 3 },
    { id: 'O007', name: '显示器', region: '华北', category: '设备', price: 1499, qty: 2 },
    { id: 'O008', name: '耳机', region: '华北', category: '配件', price: 159, qty: 5 }
  ];
}

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

export function createTasks() {
  return [
    { id: 'P', parentId: null, title: '实施项目', owner: '李经理' },
    { id: 'A', parentId: 'P', title: '需求阶段', owner: '王顾问' },
    { id: 'A1', parentId: 'A', title: '业务访谈', owner: '小周', plan: 3, actual: 2 },
    { id: 'A2', parentId: 'A', title: '需求说明书', owner: '小林', plan: 5, actual: 3 },
    { id: 'B', parentId: 'P', title: '开发阶段', owner: '赵工' },
    { id: 'B1', parentId: 'B', title: '前端开发', owner: '小陈' },
    { id: 'B11', parentId: 'B1', title: '表格界面', owner: '小陈', plan: 8, actual: 4 },
    { id: 'B12', parentId: 'B1', title: '数据提供者', owner: '小吴', plan: 5, actual: 2 },
    { id: 'B2', parentId: 'B', title: '服务端接口', owner: '小郑', plan: 10, actual: 6 },
    { id: 'C', parentId: 'P', title: '验收阶段', owner: '孙顾问' },
    { id: 'C1', parentId: 'C', title: '回归测试', owner: '小林', plan: 4, actual: 1 },
    { id: 'C2', parentId: 'C', title: '用户培训', owner: '小周', plan: 2, actual: 0 }
  ];
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
