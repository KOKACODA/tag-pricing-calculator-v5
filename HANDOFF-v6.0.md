# KOKALabel报价系统 v6.3 — 项目转手文档

> 本文档供新接手的开发者或 AI Agent 快速了解项目全貌。
> 最后更新：2026-08-14
> 版本：v6.3

---

## 一、项目概览

| 项目 | 说明 |
|---|---|
| 名称 | KOKALabel 报价系统 |
| 用途 | 标签/不干胶印刷品报价计算（纸张 + 工艺 + 吊绳 + 邮费 + 客户等级系数） |
| 部署 | Cloudflare Pages |
| 正式地址 | https://tag-pricing-calculator-v5.pages.dev |
| GitHub | KOKACODA/tag-pricing-calculator-v5 |
| 技术栈 | 原生 HTML + CSS + JavaScript（无框架），SheetJS 用于 Excel 导入导出 |
| 存储方案 | localStorage（无后端数据库） |

---

## 二、文件结构

```
tag-pricing-calculator-v5/
├── index.html          # HTML 骨架 + CSP 安全策略
├── css/
│   └── style.css       # 全部样式（~1800行）
├── js/
│   ├── data.js         # 数据配置层（价格/纸张/工艺/吊绳/邮费/客户等级/直接系数）
│   └── app.js          # 主程序（计算逻辑 + 渲染 + 交互 + 初始化，~3300行）
├── test.html           # 单元测试骨架
├── CHANGELOG.md        # 版本历史
├── HANDOFF-v6.0.md     # 本文档
├── v6.0-CHANGELOG.md   # v6.0 总结文档
├── v6.1-CHANGELOG.md   # v6.1 总结文档
├── v6.2-CHANGELOG.md   # v6.2 总结文档
└── v6.3-CHANGELOG.md   # v6.3 总结文档
```

---

## 三、核心架构

### 三级数据层级

```
小组（Group） → 总报价表（PriceList） → Sheet纸张报价表（Paper）
```

- **小组**：当前硬编码单小组（group1），后续可扩展多账号
- **总报价表**：一个小组的完整报价数据集合，对应一个 Excel 文件
- **Sheet 纸张报价表**：总报价表内的单张纸报价数据，对应 Excel 的一个 Sheet

### 两种计算模式

| 模式 | 计算方式 | 系数来源 |
|---|---|---|
| 直接系数计算（默认） | 成本 = 纸张 + 工艺（不计算吊绳/邮费） | 每 Sheet 专属 `directCoeff`（读取报价表表格三行：直接系数档位/最高倍数/最低倍数），无直接系数的 Sheet 按标准报价 |
| 标准报价 | 成本 = 纸张 + 工艺 + 吊绳 + 邮费 | CUSTOMER_LEVELS（毛利系数） |

### 直接系数（v6.2 起只读报价表表格）

直接系数只读取报价表（Excel）每个 Sheet 的「直接系数档位 / 最高倍数 / 最低倍数」三行数据，**不再支持个人主页全局设置**：

```javascript
// data.js 中每张纸的 directCoeff 字段（三行格式）
{
  "directCoeff": {
    "tiers": [500, 1000, 2000, 2500, 3000, 5000, 10000, 20000, 50000],
    "max":  [1.3, 1.3, 1.3, 1.3, 1.3, 1.2, 1.2, 1.2, 1.2],
    "min":  [1.25, 1.25, 1.25, 1.25, 1.25, 1.15, 1.15, 1.15, 1.15]
  }
}
```

- `tiers` = 直接系数档位（批量张数）
- `max` = 最高倍数（对应普通客户）
- `min` = 最低倍数（对应大客户）
- 中间等级（优质客户）按最高/最低等差插值
- 无直接系数的 Sheet：`directCoeff` 为 `null`，按标准报价计算（乘折扣系数）
- **有直接系数的 Sheet 不再计算折扣**（直接系数模式下使用原价）
- **702铜版纸** 已配置专属规则：1000-4000张 最高1.6/最低1.5；5000-10000张 最高1.5/最低1.45；20000张以上 最高1.45/最低1.4
- Excel 模板/导出/导入均支持三行格式；无直接系数的 Sheet 提供占位档位行（最高/最低留空），方便填写后重新导入

---

## 四、关键数据变量

### data.js 中的全局变量

| 变量 | 类型 | 说明 | 持久化键 |
|---|---|---|---|
| `GROUP_META` | const | 小组元信息 | 不持久化 |
| `PRICE_LISTS` | let | 报价表列表 | `priceLists` |
| `CURRENT_PRICE_LIST_ID` | let | 当前选中的报价表 ID | `currentPriceListId` |
| `PAPER_CONFIG` | let | 纸张配置数组 | `paperConfig` |
| `CRAFT_CONFIG` | let | 工艺配置对象 | `craftConfig` |
| `ROPE_CONFIG` | let | 吊绳配置数组 | `ropeConfig` |
| `SHIPPING_CONFIG` | let | 邮费配置数组 | `shippingConfig` |
| `CUSTOMER_LEVELS` | let | 毛利系数等级 | `customerLevels` |
| `DIRECT_COEFF_LEVELS` | let | 直接系数等级 | `directCoeffLevels` |
| `DIRECT_COEFF_TIER_RULES` | let | 直接系数档位规则 | `directCoeffTierRules` |

> 所有 `let` 变量通过 `loadFromStorage(key, default)` 加载，`saveToStorage(key, value)` 保存。
> localStorage 键前缀为 `tagPricing_`。

### app.js 中的关键函数

| 函数 | 说明 |
|---|---|
| `calculate(sheets, tier, isDirect)` | 核心计算函数，返回报价结果对象 |
| `getDirectCoeffsForTier(tier)` | 根据档位返回调整后的直接系数列表 |
| `matchSpec(paper, area)` | 按面积自动匹配规格代码 |
| `findSpecByDisplayCode(paper, code, area)` | 按手动选择的代码查找规格 |
| `calcAreaCoefficient(area)` | 计算面积系数（面积 > 10000 时） |
| `onCalculate()` | 触发计算并渲染结果 |
| `onCodeSwitch(sheetIdx, dir)` | 手动切换代码（上/下） |
| `switchCalcMode(mode)` | 切换标准/直接系数模式 |
| `renderDirectCoeffSettings()` | 渲染直接系数等级编辑器 |
| `renderDirectTierRules()` | 渲染档位规则编辑器 |
| `renderSheets()` | 渲染纸张设置卡片 |
| `resetToDefaults()` | 一键恢复全局默认 |

---

## 五、计算流程

### 标准报价模式

```
输入：纸张选择 + 尺寸 + 工艺 + 吊绳 + 邮费地区 + 批量档位
  ↓
1. 每张纸：面积 → 匹配规格代码 → 查价格表 → 纸张单价
2. 纸张总价 = Σ(单价 × 面积系数) × 折扣
3. 工艺费用 = Σ(每张工艺价格)
4. 吊绳费用 = 吊绳单价 × 数量
5. 邮费 = 按地区查表
6. 成本 = 纸张 + 工艺 + 吊绳 + 邮费
7. 报价 = 成本 × 毛利系数（3个等级）
```

### 直接系数模式

```
输入：纸张选择 + 尺寸 + 工艺 + 批量档位
  ↓
1. 每张纸：面积 → 匹配规格代码 → 查价格表 → 纸张单价
2. 纸张总价 = Σ(单价 × 面积系数) × 折扣
3. 工艺费用 = Σ(每张工艺价格)
4. 成本 = 纸张 + 工艺（不计算吊绳和邮费）
5. 直接系数 = getDirectCoeffsForTier(档位) → 按档位规则调整
6. 报价 = 成本 × 直接系数（3个等级，右上角显示系数徽章）
```

---

## 六、面积系数规则

当计算面积（含出血后）超过 10000mm² 时：

- 系数 = 面积 / 10000，四舍五入保留两位小数
- 价格 = 基础价格 × 面积系数
- 报价结果中显示面积系数行（平时隐藏）
- 不再弹出"联系上级"警告

---

## 七、代码手动切换（v5.8 新增）

报价结果表格中，代码列有上下箭头按钮：

- 点击上/下切换到该纸张的其他代码规格
- 手动切换后代码显示蓝色 + "手动"标签
- 切换纸张或修改尺寸时自动恢复自动匹配
- `sheetsState[i].manualCode` 存储手动选择的代码

---

## 八、部署流程

```bash
# 1. 提交代码到 GitHub
cd /data/user/work/tag-pricing-calculator-v5
git add -A
git commit -m "feat(vX.X): 功能描述"
git push origin main

# 2. 部署到 Cloudflare Pages
CLOUDFLARE_API_TOKEN="<your-token>" \
  npx wrangler pages deploy . --project-name=tag-pricing-calculator-v5

# 3. 版本号更新（三个文件同步）
# index.html: <title> + CSS/JS 引用版本参数 + 个人主页版本显示
# js/app.js: 文件头注释
# js/data.js: 文件头注释
```

---

## 九、备份策略

```bash
# 创建稳定版本标签（每次大版本前执行）
git tag vX.X-stable
git push origin vX.X-stable

# 现有备份
# - v5.2.1-stable（Git Tag）
# - v5.7-stable（Git Tag）
# - v5.8-stable（Git Tag）
# - v5.9-stable（Git Tag）
# - backup/v5.2.1（Git Branch）
```

---

## 十、注意事项

1. **版本号同步**：修改版本号时，三个文件（index.html、app.js、data.js）必须同步更新，包括 CSS/JS 引用的 `?v=X.X` 参数
2. **007/008 合并**：规格代码 007 和 008 在显示和手动切换时合并为"007/008"
3. **数据隔离**：纸张配置按报价表隔离，吊绳和邮费为全局共享
4. **CSP 策略**：`script-src` 仅允许 `self` + `cdn.sheetjs.com` + `unsafe-inline`
5. **Excel 导入**：导入超过 30 个 Sheet 时，使用 `paper_import_N` 模式生成唯一 ID 避免冲突
6. **直接系数模式**：不计算吊绳和邮费，隐藏相关 UI 元素，临时系数标签改为"临时直接系数"
7. **档位规则不校验重叠**：用户需自行确保多条规则的张数范围不重叠
8. **用户称呼**：用户要求称呼为"Master"

---

## 十一、Git 配置

```bash
git config user.email "koka@label.com"
git config user.name "KOKALabel"
```

---

## 十二、后续可优化方向

1. 档位规则重叠校验提醒
2. 直接系数模式的历史记录区分标记
3. 报价表组切换时自动切换吊绳/邮费配置（当前全局共享）
4. 移动端适配优化
5. 数据云端同步（替代 localStorage）
6. 多用户/多小组支持
