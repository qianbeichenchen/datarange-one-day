// 第一课：先只看 getValue + dataRanges.add。
const S = GC.Spread.Sheets;
const spread = new S.Workbook(document.getElementById('ss'), { sheetCount: 1 });
const sheet = spread.getActiveSheet();
sheet.name('第一课：读取');
sheet.options.isProtected = true; // 本课只读；仅省略 setValue 并不等于禁止编辑。
sheet.setColumnWidth(1, 180);
sheet.setColumnWidth(2, 110);
sheet.setColumnWidth(3, 110);

// 第 0 行就是我们自己提供的表头，不是 SpreadJS 自动生成的。
const rows = [
  ['商品', '单价', '数量'],
  ['键盘', 199, 2],
  ['鼠标', 89, 3],
  ['显示器', 1299, 1]
];

const provider = {
  getValue(row, col) {
    // row/col 是区域内部坐标，左上角始终为 (0, 0)。
    // null 表示我们提供空值；undefined 会回退到工作表自身的数据。
    return rows[row]?.[col] ?? null;
  }
};

// B3：工作表索引 (2, 1)，高 4 行、宽 3 列，覆盖 B3:D6。
const dataRange = sheet.dataRanges.add('products', provider, new S.Range(2, 1, rows.length, 3));

document.getElementById('inspect').onclick = () => console.table(rows);
// 供你在 F12 控制台动手实验；不是 DataRange 必须的配置。
window.demo = { spread, sheet, rows, provider, dataRange };
console.log('第一课已加载。试试：demo.provider.getValue(1, 0)');
