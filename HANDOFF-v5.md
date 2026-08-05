# 项目转接文档 — KOKALabel报价系统 v5.0

> **此文件用于新对话窗口接手，请先完整阅读此文档，再阅读 `AGENT-INSTRUCTIONS.md`。**

---

## 一、项目基本信息

| 项目 | 值 |
|------|-----|
| 项目名 | KOKALabel报价系统 |
| 当前版本 | v5.0 |
| 架构 | 多文件（CSS / JS / HTML 分离） |
| GitHub 仓库 | `KOKACODA/tag-pricing-calculator-v5`（Public） |
| 线上地址 | https://tag-pricing-calculator-v5.pages.dev |
| 预览地址 | https://e71782e4.tag-pricing-calculator-v5.pages.dev |
| CF Account ID | a4fdb353836f801c65072d7810fbcc32 |
| 部署方式 | GitHub push → Cloudflare Pages 自动构建 |
| 旧版仓库 | `KOKACODA/tag-pricing-calculator`（v4.3，保留不动） |
| 用户称呼 | Master |

---

## 二、文件结构

```
.
├── index.html              # HTML 骨架（826 行）+ CSP + 脚本引用
├── css/
│   └── style.css           # 全部样式（1,481 行）
├── js/
│   ├── data.js             # 数据配置 + 存储函数（1,531 行）— 先加载
│   └── app.js              # 计算 + 渲染 + 交互 + 初始化（2,555 行）— 后加载
├── test.html               # 单元测试骨架（9 个测试用例）
├── robots.txt              # 爬虫禁止
├── .gitignore
├── CHANGELOG.md            # 更新日志
├── AGENT-INSTRUCTIONS.md   # Agent 强制规则
└── HANDOFF-v5.md           # 本文件
```

### 加载顺序（不可更改）

1. `css/style.css` → `<link rel="stylesheet">`
2. SheetJS CDN → `<script defer>`（备用加载）
3. `js/data.js` → `<script>`（定义全局配置和存储函数）
4. `js/app.js` → `<script>`（定义计算、渲染、交互逻辑）

---

## 三、v5.0 改动摘要（相对 v4.3）

### P0 安全修复

| 修复项 | 说明 |
|--------|------|
| XSS 漏洞 ×11 | 11 处 `innerHTML` 未转义位置全部加上 `escapeHtml()` |
| CSP 策略 | `<meta http-equiv="Content-Security-Policy">` 限制来源 |
| SheetJS 加载 | `defer` + `loadSheetJS()` 按需加载，不阻塞首屏 |
| 导入校验 | `validateImportedData()` 校验字段类型、长度、范围 |

### P1 性能优化

| 优化项 | 说明 |
|--------|------|
| 搜索防抖 | `debounce(renderPriceTable, 200)` |
| 滚动节流 | `throttle(closeAllPaperDropdowns, 100)` |
| SheetJS 懒加载 | 8 个 Excel 函数全部接入 `loadSheetJS()` + `.catch()` |

### P2 架构改进

| 改进项 | 说明 |
|--------|------|
| 多文件拆分 | CSS/JS/HTML 完全分离 |
| 单元测试骨架 | `test.html` 包含 9 个基础测试用例 |
| 严格模式 | data.js 和 app.js 均启用 `"use strict"` |

### 不变项

- 计算逻辑（`calculate()`、`calcBleedArea()`、`matchSpec()` 等）完全不变
- 数据配置（`PAPER_CONFIG`、`CRAFT_CONFIG` 等）完全不变
- Excel 模板导入导出格式完全不变，兼容新旧格式

---

## 四、数据架构

```
小组（Group） → 总报价表（PriceList） → Sheet 纸张报价表（Paper）
```

| 层级 | 代码 | 当前值 |
|------|------|--------|
| 小组 | `GROUP_META` | 1号小组 |
| 总报价表 | `PRICE_LIST_META` | 1号报价表 |
| Sheet 纸张 | `PAPER_CONFIG` | 9 张，可扩展至 30+ |

---

## 五、关键函数索引（app.js）

| 函数名 | 用途 |
|--------|------|
| `debounce()` / `throttle()` | P1 性能工具函数 |
| `loadSheetJS()` | SheetJS 按需加载 |
| `validateImportedData()` | P0 导入数据校验 |
| `parseDecimalPlaces()` | 安全解析小数位数 |
| `formatMoney()` | 格式化金额 |
| `calculate()` | 核心计算引擎 |
| `renderSheets()` | 纸张设置卡片渲染 |
| `renderPriceTable()` | 1号报价表页面渲染 |
| `syncPaperSelector()` | 同步下拉选择器 |
| `onCalculate()` | 收集输入并计算 |
| `escapeHtml()` | XSS 转义 |
| `parsePaperExcel()` | Excel 解析（标签匹配，兼容新旧格式） |
| `bindEvents()` | 全局事件绑定 |
| `safeInit()` | 启动初始化 |

---

## 六、新对话接手步骤

1. **读取 `AGENT-INSTRUCTIONS.md`** — 了解强制规则
2. **读取 `CHANGELOG.md`** — 了解完整版本历史
3. **git clone 仓库**：`git clone https://github.com/KOKACODA/tag-pricing-calculator-v5.git`
4. **修改代码后必须执行**：
   ```bash
   # 更新 CHANGELOG.md 顶部追加新版本记录
   git add -A
   git commit -m "vX.Y: 简述改动"
   git push origin main
   ```
5. **等待 60-90 秒后验证部署**：
   ```bash
   curl -s "https://tag-pricing-calculator-v5.pages.dev?cb=$(date +%s)" | grep '关键词'
   ```

---

## 七、已知技术债 / 待优化项

| 项目 | 优先级 | 说明 |
|------|--------|------|
| ES Modules 升级 | 中 | 当前用 `<script>` 全局共享，未来可升级为 `import/export` |
| 事件委托 | 中 | `sheetList` 等可用父级委托，避免每次渲染重绑 |
| 整数分计算 | 低 | 金额用浮点累加，可改为整数分消除精度误差 |
| 快照恢复功能 | 中 | 有创建/删除快照，缺"恢复到指定快照"按钮 |
| 历史记录假数据 | 中 | 需对接真实计算结果 |
| 多账号系统 | 低 | Cloudflare Pages Functions + D1 方案已搁置 |
| 文件进一步拆分 | 低 | app.js 2,555 行，可拆分为 calculator.js / ui.js / excel.js |

---

*文档生成时间：2026-08-06 | 版本：v5.0 | 生成者：TRAE Agent*
