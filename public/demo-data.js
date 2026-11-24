// 课程样例数据。每次调用工厂函数都会创建可独立编辑的记录。
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

export function createProducts() {
  // 基础课用前三笔订单里的商品、单价和数量。
  return createOrders().slice(0, 3).map(({ name, price, qty }) => ({ name, price, qty }));
}

export function createTreeTasks() {
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

export const TASK_STATUSES = ['待开始', '进行中', '已完成'];
export const TASK_OWNERS = ['张三', '李四', '王五'];
export const TASK_STATUS_COLORS = {
  待开始: { background: '#edf1f5', foreground: '#43556a' },
  进行中: { background: '#dbeafe', foreground: '#24569c' },
  已完成: { background: '#dff1e5', foreground: '#24634d' }
};
export const TASK_MONTH = '2026-06';

export function createTasks() {
  return [
    { id: 'T01', title: '梳理字段', status: '已完成', owner: '张三', priority: 2, start: '2026-06-01', due: '2026-06-03' },
    { id: 'T02', title: '设计任务表', status: '进行中', owner: '张三', priority: 1, start: '2026-06-02', due: '2026-06-06' },
    { id: 'T03', title: '整理演示数据', status: '进行中', owner: '李四', priority: 2, start: '2026-06-04', due: '2026-06-06' },
    { id: 'T04', title: '验证编辑回写', status: '待开始', owner: '王五', priority: 1, start: '2026-06-08', due: '2026-06-10' },
    { id: 'T05', title: '补写操作指南', status: '待开始', owner: '李四', priority: 3, start: '2026-06-10', due: '2026-06-12' },
    { id: 'T06', title: '准备项目演示', status: '待开始', owner: '王五', priority: 2, start: '2026-06-15', due: '2026-06-18' }
  ];
}
