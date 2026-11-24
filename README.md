# SpreadJS DataRange 课堂

17 节课从读取和编辑数据开始，逐步做到分组汇总、卡片、看板、日历和甘特图。左侧目录按 6 个主题分组，主题下列出课程，当前课目用绿色标记。课号对应本地网址的 `?step=` 参数和 JavaScript 文件编号。

## 本地运行

需要 Node.js（建议 22 或更新版本）。在本目录执行以下命令：

```powershell
npm install
npm start
```

打开 [本地课堂](http://localhost:5186)，用左侧菜单选择课程。按 F12 打开控制台，查看工具栏“打印”按钮输出的原始记录和映射。每课独立运行，刷新会恢复该课的样例数据。

依赖固定为 SpreadJS 核心包 19.2.3。需要配置授权时，将 `public/license.example.js` 复制为 `public/license.js` 并在本机填入许可。该文件由 Git 忽略。示例直接使用 JavaScript 数据，服务监听本机 `127.0.0.1:5186`。

## 课程目录

| 主题 | 打开课程 | 主要操作 | 源码 |
| --- | --- | --- | --- |
| 数据读写与区域 | [01 只读数据](http://localhost:5186/?step=1) | 读取二维数组，从工作表坐标换算到 DataRange 内部坐标。 | [`01-read.js`](./public/01-read.js) |
|  | [02 编辑回写](http://localhost:5186/?step=2) | 编辑单元格，将合法输入写回对象字段。 | [`02-edit.js`](./public/02-edit.js) |
|  | [03 基础报价表](http://localhost:5186/?step=3) | 计算金额、设置高亮、增删商品、移动区域。 | [`03-quote.js`](./public/03-quote.js) |
| 视图映射与汇总 | [04 订单视图与撤销](http://localhost:5186/?step=4) | 按商品和地区筛选、按金额排序，编辑正确的记录并撤销或重做。 | [`04-orders.js`](./public/04-orders.js) |
|  | [05 分组报表与联动](http://localhost:5186/?step=5) | 按地区和品类分组，查看小计、合计和联动汇总。 | [`05-groups.js`](./public/05-groups.js) |
|  | [06 项目工时树](http://localhost:5186/?step=6) | 展开或折叠任务树，编辑叶子工时并观察父级汇总。 | [`06-tree.js`](./public/06-tree.js) |
| 字段、单元格与卡片 | [07 多单元格卡片](http://localhost:5186/?step=7) | 用多个可编辑单元格组成一张卡片，切换每行两张或三张。 | [`07-cards.js`](./public/07-cards.js) |
|  | [08 动态字段](http://localhost:5186/?step=8) | 隐藏负责人或交换列顺序，再编辑验证按字段 ID 回写。 | [`08-fields.js`](./public/08-fields.js) |
|  | [09 自定义绘制与单格卡片](http://localhost:5186/?step=9) | 在状态标签和整格卡片间切换，点击卡片查看详情。 | [`09-cell-paint.js`](./public/09-cell-paint.js) |
|  | [10 自定义日期编辑器](http://localhost:5186/?step=10) | 用日期控件编辑截止日期，再用原生历史撤销和重做。 | [`10-cell-editor.js`](./public/10-cell-editor.js) |
| 表格业务交互 | [11 行内新增](http://localhost:5186/?step=11) | 创建草稿，输入标题正式新增，或按 Escape 取消。 | [`11-inline-add.js`](./public/11-inline-add.js) |
|  | [12 组合筛选](http://localhost:5186/?step=12) | 组合 AND 和 OR 条件，筛选后编辑并核对源记录。 | [`12-filter-rules.js`](./public/12-filter-rules.js) |
| 看板与多视图 | [13 看板布局](http://localhost:5186/?step=13) | 按状态分列，查看卡片详情，并在选定状态下新增任务。 | [`13-board.js`](./public/13-board.js) |
|  | [14 多视图共享数据](http://localhost:5186/?step=14) | 在表格修改任务状态并在看板查看，分别保存各视图的筛选设置。 | [`14-multi-view.js`](./public/14-multi-view.js) |
|  | [15 看板拖拽](http://localhost:5186/?step=15) | 跨列拖放卡片，按目标列更新任务状态。 | [`15-board-drag.js`](./public/15-board-drag.js) |
| 日期与时间轴 | [16 日历与拖动改期](http://localhost:5186/?step=16) | 翻月、查看同日多条任务，开启拖放后更新截止日期。 | [`16-calendar.js`](./public/16-calendar.js) |
|  | [17 甘特时间轴](http://localhost:5186/?step=17) | 查看任务条、整体移动保留工期、拉伸端点调整起止日期。 | [`17-gantt.js`](./public/17-gantt.js) |

分课教程：[01～03 核心用法](./docs/01-03-core.md)、[04～06 视图与汇总](./docs/04-06-views.md)、[07 多单元格卡片](./docs/07-cards.md)。第 08～17 课的操作步骤见下文。

## 样例数据与项目目录

初始样例集中在 [`public/demo-data.js`](./public/demo-data.js)。以下数字是每课刚打开时的样例数量，筛选、折叠和新增会改变实际显示的条目数。

| 课程 | 初始数据 |
| --- | --- |
| 01～03 | 3 件商品，取自订单样例的前三笔 |
| 04、05、07 | 8 笔订单 |
| 06 | 12 个工时树节点，含父任务与叶子任务 |
| 08～17 | 6 条任务，含状态、负责人、优先级和起止日期 |

切换课程或刷新页面会重新创建样例。第 14 课内部切换表格和看板时，两个视图共享同一份任务记录，各自保留筛选设置。

页面和课程源码在 `public/`，分课教程和配图在 `docs/`，业务模型测试在 `tests/`。根目录的 `package.json` 管理启动命令和依赖。`public/package.json` 为 Node.js 读取课程源码声明 ES 模块格式。

## 第 08～17 课操作教程

读这些课时，先找到 Provider 的 `getValue`，沿行列坐标找到任务、字段或日期，再观察一次操作如何改变对应数据。

### 08 动态字段

点击“交换状态与负责人”，双击换位后的负责人单元格，把负责人改为“王五”。点击“打印列映射与记录”，核对变化出现在同一条任务的 `owner` 字段。隐藏后重新显示负责人，刚才的值仍保留。

从 `applyColumns` 开始读。它把视图的字段 ID 变成 `visibleFields`，再更新区域宽度。`getValue` 和 `setValue` 都通过 `visibleFields[col].id` 定位字段。

### 09 自定义绘制与单格卡片

先看“绘制状态标签”，打印一个状态单元格，原始值是字符串。切换到“绘制整张卡片”，再打印一次，原始值变成任务对象。点击卡片可查看详情，选中卡片后也可按空格打开。

对照第 07 课试一次双击。第 07 课一张商品卡片占多个格子，可以直接编辑名称、单价和数量。第 09 课一个格子就是一张卡片，使用 `getStyle` 选择 CellType，再由它的 `paint` 绘制标题、状态和负责人。交互从格子坐标找到整条任务。

绘制时用 `ctx.save()` 和 `ctx.restore()` 隔离画布状态，并显式设置 `ctx.textAlign = 'left'`。这样 `fillText` 的 x 坐标就是文字左边缘，标题和状态会从卡片内部的留白处向右显示。第 13～15 课的看板遵循同样的方式。

### 10 自定义日期编辑器

双击“截止日期”，修改日期并按 Enter。依次点击撤销、重做，也可以在表格有焦点时按 Ctrl+Z、Ctrl+Y。试着清空日期再提交，原值会保留。

`DateCell` 的 `setEditorValue` 和 `getEditorValue` 负责传入旧值与取回输入。Provider 用 `setValue(..., changes)` 写回日期和记录旧值，`undo(change)` 按任务 ID 恢复旧值，重做由 SpreadJS 重放命令。

### 11 行内新增

选择新任务状态，再点击表格底部“添加记录”。输入标题但先按 Escape，打印数据确认正式记录没有新增。再添加一次，输入标题后按 Enter，观察新记录获得编号并继承刚才选择的状态。“取消草稿”也能放弃当前输入。

`rebuildRows` 生成表头、正式记录及草稿或新增入口。只有 `setValue` 收到有效标题时，草稿才加入 `records`。

### 12 组合筛选

先应用默认条件“进行中 AND 张三”，看到 T02。勾选“或优先级为 1”并重新应用，会多出 T04。编辑 T04 的任务标题，再点击“显示全部”，核对编辑后的标题仍对应 T04。

把控件值组织成 `(状态 AND 负责人) OR 优先级`，`matches` 递归判断条件节点。显示列表保存任务引用，编辑沿显示行找到源记录。

### 13 看板布局

选“进行中”，点击“添加任务”，再点击新增的卡片查看详情。打印数据与分组，核对新增记录的状态和所在列一致。

从 `board-layout.js` 的 `BoardProvider.getValue` 读起。列号选状态分组，行号选组内记录。每个卡片格把一条任务交给 CellType 绘制。

### 14 多视图共享数据

在表格中双击 T01 的状态，展开下拉框，选“进行中”并按 Enter。切到看板，确认卡片在“进行中”列。给表格和看板选不同的筛选条件，来回切换确认各自保留筛选。表格的负责人列显隐也会保留。

`records` 是共享数据，`views` 保存不同的筛选及列设置。`render` 按当前视图更新坐标映射和 DataRange 尺寸。只有切换本课内部视图时才共享这份记录。

### 15 看板拖拽

把“待开始”列的一张卡片拖到“进行中”列。打印任务状态，确认只更新了该卡片的状态。再试一次拖到原列、拖到区域外或按 Escape，记录会保持原样。

按下卡片时记住任务 ID，移动时用 `sheet.hitTest` 找目标列，放开后才提交状态修改并重建分组。

### 16 日历与拖动改期

查看 2026 年 6 月 6 日的两条任务，切换到上个月后再切回。勾选“启用拖动改期”，把其中一条任务标签拖到 6 月 8 日，核对只有这条任务的截止日期变化。

`dateAt` 将工作表行列映射到一个以周一开头的 42 天日历。`getValue` 用日期取出当天的任务，CellType 在一个日期格内绘制多条任务。拖放复用同一日期映射。

### 17 甘特时间轴

1. 在“查看任务条”阶段，找到 T01 对应的 6 月 1～3 日任务条。
2. 在“整体移动”阶段，从任务条中间向右移动两天，观察起止日期都加两天、工期不变。
3. 在“移动或拉伸端点”阶段，从深色端点拖动一端。打印日期验证只修改所拖动的一端。结束日期早于开始日期的结果会被拒绝。

任务决定行，日期决定列。`dateAtColumn` 和 `columnOfDate` 提供双向映射，拖动先算列差，再决定更新哪个日期。

## 验证

已在 Chrome 检查全部 17 个课程入口、当前课目高亮和导航滚动。浏览器验证覆盖自定义卡片绘制、看板拖拽、日历改期和甘特图移动等操作。

运行下面的命令可检查订单排序、分组汇总、树形汇总、卡片坐标和输入校验。这 6 项是业务模型测试，覆盖范围见[测试源码](./tests/business-model.test.mjs)。

```powershell
node --test tests/business-model.test.mjs
```

参考 [官方 DataRange 指南](https://demo.grapecity.com.cn/spreadjs/help/docs/features/core-concept/worksheet/worksheet-data-range) 和 [多维表格源码](https://github.com/qianbeichenchen/feishu-datarange-demo)。本课程的 GitHub 仓库为 [datarange-one-day](https://github.com/qianbeichenchen/datarange-one-day)。
