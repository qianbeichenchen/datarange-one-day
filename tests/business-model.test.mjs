import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createOrders, orderView, summarize, groupReport, numericInput, createTasks, treeProjection, cardCellAt } from '../public/business-model.js';

test('卡片坐标映射跨排换行，跳过间距和末排空位；布局变化映射到正确记录', () => {
  assert.deepEqual(cardCellAt(0, 0, 3, 8), { index: 0, localRow: 0, localCol: 0, field: 'name' });
  assert.deepEqual(cardCellAt(3, 4, 3, 8), { index: 1, localRow: 3, localCol: 1, field: 'qty' });
  assert.equal(cardCellAt(6, 0, 3, 8).index, 3);
  assert.equal(cardCellAt(6, 0, 2, 8).index, 2);
  assert.equal(cardCellAt(5, 0, 3, 8), null);
  assert.equal(cardCellAt(0, 2, 3, 8), null);
  assert.equal(cardCellAt(12, 6, 3, 8), null);
  assert.equal(cardCellAt(0, 0, 3, 0), null);
  assert.equal(cardCellAt(-1, 0, 3, 8), null);
  assert.equal(cardCellAt(0, 9, 3, 8), null);
  assert.equal(cardCellAt(4, 1, 3, 8).field, null); // 金额只读
  assert.equal(cardCellAt(3, 0, 3, 8).field, null); // 标签只读
  const records = createOrders();
  const cell = cardCellAt(9, 4, 3, records.length);
  records[cell.index][cell.field] = 4;
  assert.equal(records.find(item => item.id === 'O005').qty, 4);
  assert.equal(records[0].qty, 2);
});

test('排序筛选是显示映射：不改变原始顺序，仍写回正确对象', () => {
  const records = createOrders();
  const view = orderView(records, { region: '华北', sort: 'amount-desc' });
  assert.deepEqual(view.map(item => item.id), ['O007', 'O006', 'O008']);
  view[0].qty = 7;
  assert.equal(records.find(item => item.id === 'O007').qty, 7);
  assert.equal(records[0].qty, 2);
  assert.deepEqual(records.map(item => item.id), ['O001', 'O002', 'O003', 'O004', 'O005', 'O006', 'O007', 'O008']);
  assert.equal(orderView(records, { keyword: '不存在' }).length, 0);
});

test('分组行、小计、品类摘要使用同一明细口径，汇总不重复计数', () => {
  const records = createOrders();
  assert.deepEqual(summarize(records), { count: 8, qty: 22, amount: 17528 });
  const report = groupReport(records);
  assert.equal(report.display.filter(line => line.kind === 'data').length, 8);
  assert.equal(report.display.find(line => line.kind === 'total').amount, 17528);
  assert.deepEqual(report.categories.map(item => [item.category, item.amount]), [['配件', 3233], ['设备', 14295]]);
  const first = report.display.find(line => line.kind === 'data');
  first.item.qty = 6;
  const updated = groupReport(records, '华东');
  assert.equal(updated.total.amount, 2760);
  assert.equal(updated.categories.find(item => item.category === '配件').amount, 1461);
  assert.equal(updated.display.find(line => line.kind === 'subtotal').amount, 1461);
  assert.equal(groupReport(records, '不存在').display.length, 3);
});

test('树汇总来自全部叶子，折叠不改变汇总且保留后代折叠状态', () => {
  const nodes = createTasks();
  const all = treeProjection(nodes);
  assert.equal(all.totals.get('P').plan, 37);
  assert.equal(all.totals.get('P').actual, 18);
  const collapsed = new Set(['P', 'B1']);
  const hidden = treeProjection(nodes, collapsed);
  assert.deepEqual(hidden.visible.map(line => line.node.id), ['P']);
  assert.equal(hidden.totals.get('P').actual, 18);
  collapsed.delete('P');
  const partial = treeProjection(nodes, collapsed);
  assert.ok(partial.visible.some(line => line.node.id === 'B1'));
  assert.ok(!partial.visible.some(line => line.node.id === 'B11'));
  nodes.find(node => node.id === 'A1').actual = 4;
  const edited = treeProjection(nodes, collapsed);
  assert.equal(edited.totals.get('A').actual, 7);
  assert.equal(edited.totals.get('P').actual, 20);
  assert.equal(nodes.find(node => node.id === 'P').actual, undefined);
});

test('树数据拒绝重复 ID、缺失父节点、循环', () => {
  assert.throws(() => treeProjection([{ id: 'x', parentId: null }, { id: 'x', parentId: null }]), /重复/);
  assert.throws(() => treeProjection([{ id: 'x', parentId: 'missing' }]), /不存在/);
  assert.throws(() => treeProjection([{ id: 'x', parentId: 'y' }, { id: 'y', parentId: 'x' }]), /循环/);
});

test('数值校验保留 0，拒绝空白、无穷、负数、小数数量和 0 计划', () => {
  assert.equal(numericInput('0', { integer: true }), 0);
  for (const value of ['', ' ', null, Infinity, -1, 'NaN', true]) assert.equal(numericInput(value), null);
  assert.equal(numericInput(1.5, { integer: true }), null);
  assert.equal(numericInput(0, { positive: true }), null);
  assert.equal(numericInput(1.5, { positive: true }), 1.5);
});
