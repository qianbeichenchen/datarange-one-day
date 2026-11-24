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
