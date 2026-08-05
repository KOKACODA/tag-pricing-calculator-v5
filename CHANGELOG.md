# 更新日志

所有 notable 变更均记录在此文件中。  
格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/)。

---

## [v5.0] - 2026-08-06

### 重大架构变更
- **多文件拆分**：从 6,285 行单文件 HTML 拆分为多文件结构
  - `css/style.css` — 全部样式（1,481 行）
  - `js/data.js` — 数据配置层 + 存储函数（1,531 行）
  - `js/app.js` — 计算逻辑 + 渲染 + 交互 + 初始化（2,555 行）
  - `index.html` — HTML 骨架 + 脚本引用（826 行）
  - `test.html` — 单元测试骨架（P2）

### P0 安全修复
- **XSS 漏洞修复**：11 处 `innerHTML` 未转义位置全部加上 `escapeHtml()`
  - `renderRopeRadios()` — 吊绳名称/ID 转义
  - `initOptions()` — 地区名称/ID 转义
  - `onCalculate()` — 警告信息转义
  - 报价卡片 — 客户等级名称转义
  - `renderPriceTable()` — 代码列转义
  - `updateDefaultRopeOptions()` / `updateDefaultPaperOptions()` — 下拉选项转义
  - `importLocalBackup()` — 地区重建、文件名转义
  - 纸张下拉选项 — data-paper 属性转义
- **CSP 策略**：`<meta http-equiv="Content-Security-Policy">` 限制 script/style/img/connect 来源
- **SheetJS SRI**：`defer` 加载 + `loadSheetJS()` 按需加载兜底
- **导入数据校验**：新增 `validateImportedData()` 函数
  - 字段类型检查、长度限制（200 字符）、系数范围校验
  - JSON 导入和本地备份导入均接入校验

### P1 性能优化
- **搜索防抖**：报价表搜索 `input` 事件加 200ms `debbounce()`
- **滚动节流**：`scroll` 事件加 100ms `throttle()`，减少下拉关闭频率
- **SheetJS 懒加载**：`loadSheetJS()` 按需加载，不阻塞首屏渲染
  - 8 个 Excel 导入/导出函数全部接入懒加载 + 错误处理
  - `defer` 备用加载确保用户交互前脚本已就绪
- 新增 `debounce()` / `throttle()` 工具函数

### P2 架构改进
- **单元测试骨架**：`test.html` 包含 9 个基础测试用例
  - `parseDecimalPlaces` 边界测试
  - `calcBleedArea` 面积计算测试
  - `matchSpec` 超限错误测试
  - `escapeHtml` XSS 转义测试
  - `hasExactTier` 档位匹配测试
  - `GROUP_META` / `PRICE_LIST_META` 架构元信息测试
  - `PAPER_CONFIG` 默认配置测试
- **"use strict"**：data.js 和 app.js 均启用严格模式
- **文件分离**：CSS/JS/HTML 完全分离，便于维护和版本控制

### 不变项
- 计算逻辑（`calculate()`、`calcBleedArea()`、`matchSpec()` 等）完全不变
- 数据配置（`PAPER_CONFIG`、`CRAFT_CONFIG` 等）完全不变
- Excel 模板导入导出格式完全不变，兼容新旧格式
- 三级架构（小组 → 总报价表 → Sheet 纸张报价表）不变

---

## [v4.3] - 2026-08-06

### 新增
- **三级数据架构明确**：在数据配置区新增架构层级注释和元信息常量
  - `GROUP_META`（小组：1号小组）→ `PRICE_LIST_META`（总报价表：1号报价表）→ Sheet 纸张报价表（PAPER_CONFIG）
  - 明确层级关系：小组 → 总报价表 → Sheet 纸张报价表（含规格 code × 档位价格）
  - 不改动任何现有数据和计算逻辑，仅增加架构概念定义
- **报价表查询页纸张选择器**：新增下拉选择器，可直接选择任意 Sheet 纸张查看报价表
  - Sheet 数量 > 10 时自动显示编号前缀（如"1. 350铜版纸"），方便快速定位
  - 保留原有上一张/下一张按钮，两种导航方式并存
- **计算器纸张下拉编号**：计算器页面的纸张材质下拉在 Sheet > 10 张时也显示编号前缀

### 变更
- **Excel 导出/模板新增架构层级信息**：每个 Sheet 的元数据区新增"所属小组"和"总报价表"两行
  - `paperToSheetRows()` 输出新增 2 行架构元信息
  - `downloadPaperTemplate()` 9 张模板表均新增架构元信息
  - `parsePaperExcel()` 改为按标签匹配读取（不再依赖固定行号），兼容新旧格式
  - 模板 Sheet 命名改为从 meta 中提取"简称"，不再依赖固定索引
- **报价表查询页分页信息优化**：`paperPageInfo` 显示格式从"1 / 9"改为"第 1 / 9 张"
