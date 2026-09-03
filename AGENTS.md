# AGENTS.md — 接手本仓库前先读这一份

> 本文约 60 行，是留给任何新接手的开发者或 AI Agent 的「只读入口」。读完即可开工，不用通读整个代码库。

## 一句话定位

KOKALabel 报价系统：纯前端吊牌/标签印刷报价计算器。原生 HTML + CSS + JS，无框架、无后端，数据存浏览器 `localStorage`。线上：https://tag-pricing-calculator-v5.pages.dev（v8.9.0）。

## 分支真相（先确认，最容易踩坑）

- `v8` = 当前线上版本（v8.9.0），**一切改动在这里**。
- `main` = 旧谱系 v7.10 存档，不要在上面开发。
- 两条谱系无共同提交（独立 git 历史），勿做跨分支 merge。

## 核心数据流（一句话）

报价表（Excel 多 Sheet）→ 纸张 Paper（规格矩阵）→ `calculate()` → 成本/报价。
计算逻辑全在 `js/app.js`；价格数据全在 `js/data.js`。一个报价表 = 一个 Excel 文件，每张纸 = 一个 Sheet。

## 文件地图（含行数 —— 别把大文件整个读）

| 文件 | 行数 | 何时读 / 怎么读 |
|---|---|---|
| `js/data.js` | ~10800 | **90% 是静态价格数据。永不整读**。只看顶部 `DEFAULT_*` 配置区；改价用 grep 定位纸张 id 或简称 |
| `js/app.js` | ~4600 | 按函数读（见下方函数索引），不要整读 |
| `index.html` | ~1000 | 改页面结构 / 加对话框时读 |
| `css/style.css` | ~2900 | 改样式时才读，通常不动 |
| `js/vendor/xlsx.full.min.js` | 内置库 | **永不读**。SheetJS 已本地化，离线可用 |
| `AGENTS.md` | ~60 | 本文档，接手必读 |

## 常见改动怎么做（省 token 的关键）

- **改价格**：只动 `js/data.js` 顶部 `DEFAULT_PAPER_CONFIG` / `DEFAULT_CRAFT_CONFIG` / `DEFAULT_ROPE_CONFIG` / `SHIPPING_CONFIG` / `DEFAULT_CUSTOMER_LEVELS`。
- **改计算规则**：`js/app.js` → `calculate()` 及其调用的函数。
- **加页面/按钮**：`index.html` 结构 + `app.js` 对应 render / 事件绑定。
- **改导入导出**：`app.js` 里 `downloadPaperTemplate` / `exportPaperExcel` / `importPaperExcel` / `parseShippingExcel`。

## 函数索引（按需 grep）

- `calculate(inputs)` — 核心计算
- `getDirectCoeffsForTier` / `matchSpec` / `calcAreaCoefficient` — 系数与规格匹配
- `applyDefaultQuoteVisibility` — 报价区显隐控制
- `openShippingWeightDialog` / `openManualShippingDialog` — 邮费两档对话框
- `loadFromStorage` / `saveToStorage` / `migrate*` — 存储与版本迁移

## 验证命令

```bash
node --check js/app.js js/data.js   # 语法检查
node --test tests/*.test.mjs        # 单元测试
```

## 更深层内容（按需再读，一上手不要读）

- `docs/项目总结.md`（~250 行）— 设计新功能 / 要懂全部业务规则时再读
- `docs/main-branch-summary.md` — main 分支（v7.10 旧谱系）完整历史档案，做课题研究 / 追旧版本演化时读
- `docs/问题日志.md`、`docs/归档说明-v8.md` — 问题追踪 / 迁移来历
- `CHANGELOG.md` — 查历史用；近期改动可用 `git log --oneline -10` 快速扫

## 改完必做

- 更新 `CHANGELOG.md`（版本号递增，或加「维护」记录）
- 若文件地图 / 架构有变，同步更新本文件对应段落