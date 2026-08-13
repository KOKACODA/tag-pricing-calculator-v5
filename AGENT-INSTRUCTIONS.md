# Agent 指令（必读）

> **任何 AI Agent 接手本项目时，必须先阅读此文件，并严格遵守以下规则。**

## 项目概述

- **项目名**：KOKALabel报价系统
- **当前版本**：v6.0
- **架构**：多文件（CSS / JS / HTML 分离）
- **GitHub 仓库**：`KOKACODA/tag-pricing-calculator-v5`
- **线上地址**：https://tag-pricing-calculator-v5.pages.dev
- **旧版仓库**：`KOKACODA/tag-pricing-calculator`（v4.3，保留不动）

## 文件结构

```
.
├── index.html              # HTML 骨架 + 脚本引用
├── css/
│   └── style.css           # 全部样式
├── js/
│   ├── data.js             # 数据配置 + 存储函数（先加载）
│   └── app.js              # 计算逻辑 + 渲染 + 交互 + 初始化（后加载）
├── test.html               # 单元测试骨架
├── robots.txt              # 爬虫禁止
├── CHANGELOG.md            # 更新日志
├── AGENT-INSTRUCTIONS.md   # 本文件
└── HANDOFF-v5.md           # 项目交接文档
```

## 加载顺序

**必须按此顺序加载，不可更改：**
1. `css/style.css` — `<link rel="stylesheet">`
2. `https://cdn.sheetjs.com/xlsx-0.20.1/...` — `<script defer>`（备用）
3. `js/data.js` — `<script>`（定义全局配置和存储函数）
4. `js/app.js` — `<script>`（定义计算、渲染、交互逻辑）

## 数据架构

三级层级关系：**小组（Group） → 总报价表（PriceList） → Sheet 纸张报价表（Paper）**

- `GROUP_META`：小组元信息（当前硬编码 1号小组）
- `PRICE_LIST_META`：总报价表元信息（当前硬编码 1号报价表）
- `PAPER_CONFIG`：Sheet 纸张报价表数组（当前 9 张，可扩展至 30+）

## 强制规则

### 1. 每次修改必须更新 CHANGELOG.md

版本号规则：
- 安全修复 / Bug Fix：patch +1（如 v5.0 → v5.0.1）
- 新功能：minor +1（如 v5.0 → v5.1）
- 重大重构：major +1（如 v5.0 → v6.0）

### 2. 提交规范

```bash
git add -A
git commit -m "vX.Y: 简述改动内容"
git push origin main
```

### 3. 安全规则（P0 强化）

- **所有 `innerHTML` 中插入的变量必须用 `escapeHtml()` 转义**，无例外
- CSP meta 标签不可移除
- 导入数据必须经过 `validateImportedData()` 校验
- 不使用 `eval()` / `new Function()` / `document.write()`

### 4. 性能规则（P1 强化）

- 搜索类 `input` 事件必须使用 `debounce()`
- `scroll` / `resize` 事件必须使用 `throttle()`
- SheetJS 通过 `loadSheetJS()` 按需加载，不在 head 中预加载

### 5. 代码质量要求

- `"use strict"` 必须保留
- 不要引入 TRAE IDE 的 CSS 变量
- `select` 元素只绑定 `change` 事件
- 及时删除无用函数和冗余代码
- 避免不必要的对象复制或克隆
- 避免重复计算，使用适当的数据结构和算法

### 6. 称呼

用户称呼为 **Master**。
