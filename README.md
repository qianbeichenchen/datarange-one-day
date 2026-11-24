# SpreadJS DataRange 一天入门

从 [一天学会 DataRange 核心用法.md](./一天学会%20DataRange%20核心用法.md) 开始，边读边运行三个独立阶段。

已学完基础三课，可继续阅读 [DataRange 进阶：从订单视图到树形报表](./DataRange%20进阶：从订单视图到树形报表.md)，运行 `/?step=4` 的三个复杂业务案例。

七课统一通过 `/?step=1`～`/?step=7` 访问，左侧按“基础入门 / 进阶场景”分组。源码按 `01`～`07` 编号；旧进阶链接会自动跳转到对应的新路径。

第七课新增 [卡片布局教程](./07-卡片布局教程.md)：用一块 DataRange 排列多张可编辑商品卡片，学习二维坐标映射、合并标题、原生历史、两列/三列切换与筛选。

## 运行

需要 Node.js（建议 22 或更新版本）。在本目录执行：

```powershell
npm install
Copy-Item public/license.example.js public/license.js
npm start
```

打开 http://localhost:5186 。默认第一课；左侧按难度分组，可切换全部七课。按 F12 打开控制台。

| 页面 | 源码 | 学习重点 |
| --- | --- | --- |
| `/?step=1` | `public/01-read.js` | getValue、add、两个坐标系 |
| `/?step=2` | `public/02-edit.js` | 对象数组、setValue、编辑回写 |
| `/?step=3` | `public/03-quote.js` | 计算金额、getStyle、repaint、range、增删行 |
| `/?step=4` | `public/04-orders.js` | 排序筛选、显示映射、changes + undo 接入原生撤销/重做 |
| `/?step=5` | `public/05-groups.js` | 地区/品类分组、合并表头、汇总、两个区域联动 |
| `/?step=6` | `public/06-tree.js` | 多层任务树、双击展开折叠、叶子编辑向上汇总 |
| `/?step=7` | `public/07-cards.js` | 卡片坐标映射、合并标题、编辑、原生历史、布局切换 |

进阶案例共享的纯业务逻辑在 `public/business-model.js`，可运行 `node --test tests/business-model.test.mjs` 验证。浏览器模块统一使用 `.js`，避免部分静态服务无法识别 `.mjs` 响应类型导致白屏。`public/package.json` 仅为 Node.js 测试声明这部分源码使用 ES module，不增加依赖。本机默认入口为 http://localhost:5186/?step=4 。

只依赖 SpreadJS 核心包；无后端数据服务。刷新或切换课程会重置本课数据。服务仅监听本机 127.0.0.1。

版本固定在 package.json / package-lock.json。首次克隆后先把 `public/license.example.js` 复制为 `public/license.js`。授权请在本机的 `public/license.js` 填入适用的许可证，该文件已被 Git 忽略；本仓库不包含许可证及 SpreadJS 专有库文件，依赖通过 npm 安装，未配置授权时受 SpreadJS 试用机制限制。

业务数据保存在 JavaScript 数组中。基础版关闭撤销；教程说明自定义 undo 的扩展方式。金额是 JavaScript 计算值，不是电子表格公式。所有学习诊断信息输出至浏览器控制台。

第四课按官方 DataRange 文档实现 `setValue(..., changes)` 与 `undo(change)`，启用 SpreadJS 原生撤销/重做，支持页面按钮和表格焦点下的 Ctrl+Z / Ctrl+Y。编辑期间保持行顺序；点击“应用视图”才重新筛选排序，并清空旧布局的历史。第五、六课不启用编辑历史。进阶案例之间的数据与历史相互独立。详见 [原生撤销重做验证](./原生撤销重做验证.md)。

第七课沿用原生 changes + undo。应用布局与筛选、新增商品、移动区域时清空旧历史；只编辑字段时保持卡片位置。
