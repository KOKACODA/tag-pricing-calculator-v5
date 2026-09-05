# KOKALabel 报价系统 v8.9.0 — 项目转手文档

> 本文档供新接手的开发者或 AI Agent 快速了解项目全貌与当前状态。
> 最后更新：2026-09-05 ｜ 当前版本：v9.4.2（`v8` 分支）
> 技术细节的权威来源是 `docs/项目总结.md`；全量历史 + 技术 + 转手方案见 `docs/项目历史与技术总档案.md`。本文档仅作快速上手索引。

---

## 一、项目概览

| 项目 | 说明 |
|---|---|
| 名称 | KOKALabel 报价系统 |
| 用途 | 吊牌 / 标签 / 不干胶印刷品报价（纸张 + 工艺 + 吊绳 + 邮费 + 客户等级系数） |
| 部署 | Cloudflare Pages（项目名 `tag-pricing-calculator-v5`） |
| 正式地址 | https://tag-pricing-calculator-v5.pages.dev（当前 v8.9.0） |
| GitHub | KOKACODA/tag-pricing-calculator-v5 |
| 技术栈 | 原生 HTML + CSS + JavaScript（无框架）；SheetJS 本地化于 `js/vendor/xlsx.full.min.js` |
| 存储 | 浏览器 `localStorage`（键前缀 `tagPricing_`，无后端、无数据库） |

---

## 二、分支与版本关系（重要）

| 分支 | 版本 | 状态 |
|---|---|---|
| `v8` | v9.4.2 | **当前线上版本** |
| `main` | v9.4.2 | 已与 `v8` 同步（v9.4.2 起），不再保留旧 v7.10 |

- v8 谱系来自另一台设备导出的压缩包（`8.5` / `8.6` / `8.8` 三个文件夹），根提交为 `v7.8 baseline`，与 `main` **无共同提交**，是两条独立 git 历史。
- 生产域名已切换至 v8 分支最新提交；`main` 仅作历史保留，请勿再基于它开发。
- 完整标签 `v7.8.1` ~ `v8.9.0`（含中间改动点）+ 快照标签 `archive-8.5/8.6/8.8` 已推送 GitHub。

---

## 三、文件结构（`v8` 分支）

```
tag-pricing-calculator-v5/
├── index.html                      # HTML 骨架 + CSP（约 1000 行）
├── css/style.css                   # 全部样式（约 2900 行）
├── js/
│   ├── data.js                     # 数据配置 + 存储 + 版本迁移（约 10800 行）
│   ├── app.js                      # 计算 + 渲染 + 交互 + 导入导出（约 4600 行）
│   └── vendor/xlsx.full.min.js     # SheetJS（本地化，离线可用）
├── tests/*.test.mjs                # Node 内置测试（vm 加载纯函数）
├── _headers / robots.txt           # Cloudflare 安全头 / noindex 屏蔽收录
├── .gitignore                      # 忽略 node_modules / .wrangler 等
├── CHANGELOG.md                    # 版本变更日志
└── docs/
    ├── HANDOFF-v8.md               # 本文档（转手 / 交接）
    ├── 项目总结.md                 # 完整项目总结（技术细节权威来源）
    ├── 归档说明-v8.md              # v8 谱系迁移归档
    ├── 问题日志.md                 # 已知问题与处理记录
    └── plans/                      # 设计文档
```

---

## 四、核心架构

### 三级数据层级

```
小组 Group（group1「1楼小组」、group2「3楼小组」）
 └─ 报价表 PriceList（「1楼」34 张纸、「3楼」10 张纸）
     └─ 纸张 Paper（每张纸 = 一个 Excel Sheet）
```

- 一个报价表 = 一个 Excel 文件（多 Sheet）。
- 吊绳 / 邮费 / 客户等级为全局共享配置，不随报价表切换。
- 纸张关键字段：`discount`（折扣，仅标准模式）、`directCoeff`（直接系数三行 `{tiers,max,min}`）、`batchDirect`（批量固定价 `{maxArea,prices}`）、`specs`（规格矩阵 `{code,maxArea,prices}`）。

### 三种计价路径

| 模式 | 成本 | 报价 |
|---|---|---|
| 直接系数（默认） | 纸张 + 工艺 | (纸张 + 工艺) × 直接系数 |
| 标准报价 | 纸张折后价 + 工艺 + 吊绳 + 邮费 | 成本 × 客户等级毛利系数 |
| 批量直接报价 | 固定价 + 工艺（面积 ≤ `batchDirect.maxArea` 且档位有价） | 直接用固定价，不打折不乘系数 |

- 出血面积 = (长+3)×(宽+3)；面积 > 10000 时面积系数 = 面积/10000。
- 强制面积映射：4000–4999→005，5000–5499→055，5500–6000→006（6000 边界强制 006）。
- 邮费 9 地区，档位 500/1000/2000/5000/10000；超量(>10000)弹重量输入 × 超量系数，中间档弹手动输入。
- 「修改后成本」红字**始终显示**，不受「默认报价」开关控制。
- 「吊牌成本合计」= 成本合计 − 邮费（标准模式）/ 成本合计（直接系数模式）。

---

## 五、开发与部署

### 本地运行

```bash
git clone https://github.com/KOKACODA/tag-pricing-calculator-v5.git
git checkout v8                     # 切到当前线上分支
python3 -m http.server 8080         # 或直接双击 index.html（离线可用）
```

### 验证命令

```bash
node --check js/app.js js/data.js   # 语法检查
node --test tests/*.test.mjs        # 单元测试
```

### 上线流程

- 平台为 Cloudflare Pages（项目名 `tag-pricing-calculator-v5`），生产域名已指向 v8 分支最新提交。
- SheetJS 已本地化，导出 / 导入 Excel 无需联网。
- 任何修改请同步更新 `CHANGELOG.md` 与本文档的「最后更新」信息。

---

## 六、维护注意事项

1. **GitHub 默认分支仍是 `main`（v7.10 旧版）**：受 GitHub Token 权限限制，未能自动切换默认分支。需在仓库 Settings → General → Default branch 手动改为 `v8`，避免新克隆默认拿到旧版。
2. **`account_id` 历史遗留**：早期提交的 `.wrangler/cache/pages.json` 及已删除的 `HANDOFF-v5.md` 含 Cloudflare account_id（`a4fdb…`，项目名 `tag-pricing-calculator-v5`）。当前树已清除，但 git 历史仍保留。该值为半公开标识（非访问令牌）；彻底清除需 `git filter-repo` 改写历史并会破坏全部标签，故未执行。
3. **文档分工**：技术细节以 `docs/项目总结.md` 为准；问题追踪见 `docs/问题日志.md`；迁移来历见 `docs/归档说明-v8.md`。
4. **数据迁移**：旧版 localStorage 用户首次加载会自动迁移到 v8 结构（`data.js` 内置幂等 `migrate*` 函数），无需手动重置。
5. **新增功能先看占位**：云同步 / 客户管理 / 订单管理 / 统计 / 价格趋势 / PDF 报价单目前仍为占位功能，未实现。

---

## 七、快速索引

- 计算逻辑入口：`js/app.js` → `calculate(inputs)`
- 数据唯一维护区：`js/data.js` 顶部 `DEFAULT_PAPER_CONFIG` / `DEFAULT_CRAFT_CONFIG` / `DEFAULT_ROPE_CONFIG` / `SHIPPING_CONFIG` / `DEFAULT_CUSTOMER_LEVELS`
- 邮费三档 + 对话框：`js/app.js` → 邮费段 / `openShippingWeightDialog` / `openManualShippingDialog`
- 显隐控制：`js/app.js` → `applyDefaultQuoteVisibility`
- 吊牌成本合计：`index.html` `#resTagCost` + `js/app.js` 计算后赋值
- 模板 / 导入导出：`js/app.js` → `downloadPaperTemplate` / `exportPaperExcel` / `importPaperExcel` / `parseShippingExcel` 等