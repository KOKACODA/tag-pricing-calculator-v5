<div align="center">

# KOKALabel 报价系统 v5.0

**吊牌制造行业 · 纯前端报价计算工具**

多纸张报价 · 工艺核算 · 吊绳邮费 · 客户等级 · Excel 导入导出

</div>

---

## 线上体验

| 版本 | 地址 | 说明 |
|------|------|------|
| **v5.0（当前）** | [tag-pricing-calculator-v5.pages.dev](https://tag-pricing-calculator-v5.pages.dev) | 多文件架构 · 安全加固 · 性能优化 |
| v4.3（旧版） | [tag-pricing-calculator.pages.dev](https://tag-pricing-calculator.pages.dev) | 单文件架构 · 保留不动 |

---

## 功能一览

### 智能报价计算器
- 多纸张报价（默认 2 张，可增减），每张纸独立设置
- 纸张材质选择（9 张纸张报价表，可扩展至 30+）
- 尺寸输入（宽 × 长），支持单张 / 展开两种尺寸类型
- 附加工艺多选（双面过哑胶、烫金、UV 等）
- 批量档位快速切换（药丸按钮，一键切档重算）
- 吊绳选择 + 配送地区选择
- 三档客户等级报价卡片（毛利系数可自定义）
- 小数位数自定义（0-4 位，默认 2 位）

### 1号报价表查询
- 全部纸张报价表浏览（下拉选择器 + 翻页两种导航）
- 搜索过滤规格代码
- 工艺价格表同步展示
- Sheet > 10 张时自动编号前缀

### 个人主页
- 公司信息设置（名称、电话、默认档位/纸张/吊绳）
- 客户等级管理（增删改、毛利系数）
- 数据管理（全量导出/导入、本地备份/恢复、快照创建/删除）

### Excel 导入导出
- 纸张报价表导入导出（每个 Sheet = 一张纸）
- 吊绳报价、邮费报价独立导入导出
- 9 张模板表一键下载
- 导入解析按标签匹配，兼容新旧格式

---

## 数据架构

```
小组（Group） → 总报价表（PriceList） → Sheet 纸张报价表（Paper）
```

| 层级 | 说明 | 当前值 |
|------|------|--------|
| 小组 | 独立团队/账号 | 1号小组 |
| 总报价表 | 完整报价数据集合 = 1 个 Excel 文件 | 1号报价表 |
| Sheet 纸张 | 单张纸报价数据 = Excel 的 1 个 Sheet | 9 张（可扩展） |

每张纸含多个规格（code）× 多个档位（tier）的价格矩阵。

---

## 技术栈

| 层面 | 技术 |
|------|------|
| 前端 | 原生 HTML5 + CSS3（CSS 变量）+ 原生 JavaScript（ES6+，严格模式） |
| 表格处理 | SheetJS（xlsx 0.20.1，按需懒加载） |
| 存储 | localStorage（客户等级、应用配置、快照、历史） |
| 部署 | GitHub → Cloudflare Pages |
| 测试 | 原生测试骨架（9 个用例，`test.html`） |

---

## 项目结构

```
.
├── index.html              # HTML 骨架 + CSP 安全策略
├── css/
│   └── style.css           # 全部样式（响应式 · 桌面/iOS/安卓）
├── js/
│   ├── data.js             # 数据配置层 + 存储函数
│   └── app.js              # 计算逻辑 + 渲染 + 交互 + 初始化
├── test.html               # 单元测试骨架
├── robots.txt              # 爬虫禁止（测试阶段）
├── CHANGELOG.md            # 完整版本历史
├── AGENT-INSTRUCTIONS.md   # Agent 强制规则
└── HANDOFF-v5.md           # 项目交接文档
```

---

## v5.0 改动摘要

### P0 安全修复
- 11 处 `innerHTML` XSS 漏洞全部加 `escapeHtml()` 转义
- 新增 CSP（Content-Security-Policy）meta 标签
- 新增 `validateImportedData()` 导入数据校验（字段类型、长度、范围）
- SheetJS 改为 `defer` + 按需懒加载

### P1 性能优化
- 搜索输入加 200ms 防抖（`debounce()`）
- 滚动事件加 100ms 节流（`throttle()`）
- 8 个 Excel 函数全部接入 `loadSheetJS()` 懒加载

### P2 架构改进
- 6,285 行单文件拆分为 CSS / JS / HTML 多文件
- 新增单元测试骨架（9 个用例）
- 启用 `"use strict"` 严格模式

> 计算逻辑、数据配置、Excel 模板格式完全不变，与 v4.3 100% 兼容。

---

## 快速开始

```bash
# 克隆仓库
git clone https://github.com/KOKACODA/tag-pricing-calculator-v5.git

# 直接打开 index.html 即可使用（无需 npm install）
# 或用本地服务器：
cd tag-pricing-calculator-v5
python3 -m http.server 8080
# 浏览器打开 http://localhost:8080
```

运行测试：
```bash
# 浏览器打开 test.html
open test.html
```

---

## 版本历史

| 版本 | 日期 | 关键改动 |
|------|------|----------|
| **v5.0** | 2026-08-06 | 多文件拆分 · P0 安全修复 · P1 性能优化 · P2 测试骨架 |
| v4.3 | 2026-08-06 | 三级数据架构 · 下拉选择器 · Excel 架构信息 |
| v4.2.1 | 2026-08-05 | 报价表更名 · 公司名替换 |
| v4.2 | 2026-08-05 | 全端响应式 UI · 下拉视口边界检测 |
| v4.0 | 2026-08-05 | 删除 CSS 污染 · 修复 XSS/快照/事件绑定 |
| v3.4 | 2026-08-03 | 初始版本 |

完整历史见 [CHANGELOG.md](CHANGELOG.md)。

---

## 相关文档

- [项目交接文档](HANDOFF-v5.md) — 新对话/Agent 接手必读
- [Agent 强制规则](AGENT-INSTRUCTIONS.md) — 代码修改规范
- [全流程分析报告](https://github.com/KOKACODA/tag-pricing-calculator/blob/main/docs/tag-pricing-v4-analysis/tag-pricing-v4-analysis.html) — 架构/机制/优化方向

---

## 核心原则

- 源 Excel 数据 100% 不修改
- 空值不补 0，统一用「无该批量定价」占位
- 不臆造数据，未匹配档位用 `null` 表示
- 所有 `innerHTML` 插入用户数据处使用 `escapeHtml()` 转义

---

<div align="center">

**测试阶段** · 搜索引擎已屏蔽（`noindex` + `robots.txt`）

© 2026 KOKALabel · Powered by Cloudflare Pages

</div>
