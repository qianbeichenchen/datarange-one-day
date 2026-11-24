import { createProducts } from './demo-data.js';

// 第二课：对象数组 + 字段映射 + setValue。
const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), {
  sheetCount: 1, allowUndo: false, allowUserDragDrop: false, allowUserDragFill: false
});
const sheet = spread.getActiveSheet();
sheet.name('第二课：回写');
sheet.options.isProtected = true;
sheet.setColumnWidth(1, 180);
sheet.setColumnWidth(2, 110);
sheet.setColumnWidth(3, 110);

const rows = createProducts();
const fields = ['name', 'price', 'qty'];
const headers = ['商品', '单价', '数量'];

const provider = {
  getValue(row, col) {
    if (row === 0) return headers[col] ?? null;
    return rows[row - 1]?.[fields[col]] ?? null;
  },
  setValue(row, col, value) {
    if (row === 0 || !rows[row - 1]) return true;
    const field = fields[col];
    if (!field) return true;
    const next = field === 'name' ? String(value ?? '') : Number(value);
    if (field !== 'name' && (value === '' || value == null || !Number.isFinite(next) || next < 0)) {
      console.warn('请输入非负数；本次输入被忽略。');
      return true; // 已消费写入请求；false 不是“校验失败”，会继续写入工作表模型。
    }
    rows[row - 1][field] = next;
    console.log('写回业务对象：', { row, col, field, value: next });
    return true;
  },
  getStyle(row) {
    const style = new S.Style();
    style.locked = row === 0; // 配合 isProtected，保护表头、开放数据行。
    if (row === 0) { style.backColor = '#24634d'; style.foreColor = '#ffffff'; }
    return style;
  }
};

const dataRange = sheet.dataRanges.add('products', provider, new S.Range(2, 1, rows.length + 1, fields.length));
document.getElementById('inspect').onclick = () => console.table(rows);
window.demo = { spread, sheet, rows, provider, dataRange };
console.log('第二课已加载。双击 C4，把键盘单价改为 299，再打印数据。');
