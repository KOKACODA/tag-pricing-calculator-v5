# Quote History and Visibility Controls Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 增加可持久化的标准默认报价显隐开关，并交付包含客户信息、结果快照、详情查看、删除和参数重载的本地报价历史功能。

**Architecture:** 保持现有静态 HTML/CSS/JavaScript 和 localStorage 架构。可测试的编号、显隐判断和记录构造逻辑放在 DOM 初始化前的纯函数中；界面层复用现有报价结果、历史标签页和备份数据流。

**Tech Stack:** HTML5、CSS3、原生 JavaScript、localStorage、Node.js 内置测试、Git。

---

### Task 1：建立版本边界与设计记录

**Files:**

- Create: `docs/plans/2026-09-02-quote-history-design.md`
- Create: `docs/plans/2026-09-02-quote-history-implementation.md`

1. 在当前发布提交建立 `v7.8.1-before-quote-history` 注释标签。
2. 创建 `feature/quote-history-controls` 功能分支。
3. 写入设计与实施文档。
4. 提交文档，形成独立设计检查点。

### Task 2：先写纯逻辑失败测试

**Files:**

- Create: `tests/quote-history.test.mjs`
- Modify: `js/app.js`

1. 测试标准模式且偏好关闭时隐藏默认报价，直接系数模式不受影响。
2. 测试固定时间生成 `BJ-YYYYMMDD-HHMMSS` 编号。
3. 测试客户名为空时用编号作为标题，有客户名时使用修剪后的客户名。
4. 测试记录中的输入和结果为深拷贝，原对象后续变化不污染历史快照。
5. 运行测试，确认纯函数未实现前失败。

### Task 3：实现持久化默认报价开关

**Files:**

- Modify: `index.html`
- Modify: `css/style.css`
- Modify: `js/app.js`
- Test: `tests/quote-history.test.mjs`

1. 在参数输入标题控制组增加 `role="switch"` 按钮。
2. 增加 `defaultQuoteVisible` 本地偏好加载与保存。
3. 只在标准报价模式按偏好隐藏默认报价标签和三张价格卡片。
4. 在初始化、模式切换和重新计算后同步开关与结果显隐。
5. 运行测试并提交独立功能检查点。

### Task 4：实现报价记录保存与详情查看

**Files:**

- Modify: `index.html`
- Modify: `css/style.css`
- Modify: `js/app.js`
- Test: `tests/quote-history.test.mjs`

1. 在报价结果末尾增加客户名称、订单备注和保存按钮。
2. 保存当前输入参数与 `_lastResult` 深拷贝快照。
3. 将历史占位表改为正式列表，显示客户/编号、时间、模式、档位、成本和建议价摘要。
4. 增加历史详情模态框，展示纸张、工艺、吊绳、邮费、成本和各等级报价。
5. 保留旧占位记录的兼容显示。

### Task 5：实现删除与参数重新载入

**Files:**

- Modify: `js/app.js`
- Modify: `index.html`

1. 删除单条记录前请求确认，删除后重绘列表。
2. 载入前验证原报价表和纸张仍存在。
3. 恢复模式、档位、纸张/尺寸/工艺、吊绳和地区。
4. 切换回计算器并按当前价格重新计算。
5. 移除“添加示例记录”占位按钮及对应事件。

### Task 6：验证与发布版本

**Files:**

- Verify: `index.html`
- Verify: `css/style.css`
- Verify: `js/app.js`
- Verify: `tests/quote-history.test.mjs`

1. 运行 `node --test tests/*.test.mjs`。
2. 运行 `node --check js/app.js` 和 `node --check js/data.js`。
3. 启动本地静态服务器进行浏览器端完整流程测试。
4. 检查 Git 差异和工作区状态。
5. 提交功能、快进合并 `main` 并创建 `v7.8.2-quote-history` 回退标签。

