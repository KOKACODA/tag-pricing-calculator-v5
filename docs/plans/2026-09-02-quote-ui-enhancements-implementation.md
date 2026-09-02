# Quote UI Enhancements Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 完成成本合计强调、工艺/吊绳初始价上标、纸张模糊搜索和吊绳费用明细，并保留清晰的 Git 回退点。

**Architecture:** 保持现有静态 HTML/CSS/JavaScript 架构。把可复用的价格与匹配规则写成无 DOM 依赖的纯函数，渲染层调用这些函数生成界面；计算公式不做调整。

**Tech Stack:** HTML5、CSS3、原生 JavaScript、Node.js 内置测试、Git。

---

## Task 1：建立逻辑测试

**Files:**

- Create: `tests/ui-enhancements.test.mjs`
- Test: `tests/ui-enhancements.test.mjs`

1. 从 `js/data.js` 和 `js/app.js` 的 DOM 初始化前代码中加载待测纯函数。
2. 断言“白色方头”初始价格为 `4.5`，“米色棉绳”初始价格为 `9`。
3. 断言纸张简称、全称、部分关键字和多关键词均能匹配，错误关键字不能匹配。
4. 运行测试并确认在实现函数前失败。

## Task 2：实现初始价格与纸张匹配纯函数

**Files:**

- Modify: `js/app.js`
- Test: `tests/ui-enhancements.test.mjs`

1. 增加固定的初始价格档位 `1000`。
2. 增加安全读取并格式化初始价格的函数。
3. 增加纸张搜索文本归一化与分词包含匹配函数。
4. 运行逻辑测试并确认通过。

## Task 3：接入工艺和吊绳红色上标

**Files:**

- Modify: `js/app.js`
- Modify: `css/style.css`

1. 在吊绳选项模板中渲染 1000 档价格数字。
2. 在附加工艺选项模板中渲染同样的初始价格数字。
3. 为两类选项增加右上角红色上标样式，并为名称预留空间。

## Task 4：接入纸张搜索预填

**Files:**

- Modify: `js/app.js`
- Modify: `css/style.css`

1. 在纸张选项面板顶部加入搜索框与无结果提示。
2. 输入时筛选当前下拉面板内的纸张选项。
3. 支持回车选择首条匹配结果和 Escape 关闭面板。
4. 打开面板时重置搜索条件，选择纸张后继续触发原有重新渲染与报价计算。

## Task 5：增强报价结果

**Files:**

- Modify: `index.html`
- Modify: `js/app.js`
- Modify: `css/style.css`

1. 给成本合计行添加专用语义类并设置大号红色粗体样式。
2. 把吊绳费用改为“类型 × 数量 = 费用”的动态明细。
3. 保持直接报价模式和缺价提示的原行为。

## Task 6：完整验证与版本固化

**Files:**

- Verify: `index.html`
- Verify: `css/style.css`
- Verify: `js/app.js`
- Verify: `js/data.js`

1. 运行 `node --test tests/ui-enhancements.test.mjs`。
2. 运行 `node --check js/app.js` 和 `node --check js/data.js`。
3. 启动本地服务器进行浏览器交互测试，并检查控制台无错误。
4. 检查 Git 差异仅包含本次需求与对应文档/测试。
5. 提交功能版本并创建带注释的回退标签。
