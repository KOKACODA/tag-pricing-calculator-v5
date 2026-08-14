# 更新日志

所有 notable 变更均记录在此文件中。  
格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/)。

---

## [v6.1] - 2026-08-14

### 新增
- **导入新报价表为默认数据库**：`DEFAULT_PAPER_CONFIG` 与 `DEFAULT_CRAFT_CONFIG` 全面更新为 2026-08-14 新版报价表数据（10 张 Sheet，含 602米白卡）
- **直接系数改为每个 Sheet 单独配置**：每张纸（Sheet）拥有独立的 `directCoeff` 字段
  - 702铜版纸 已配置专属档位规则：1000-4000张 最高1.6/最低1.5；5000-10000张 最高1.5/最低1.45；20000张以上 最高1.45/最低1.4
  - 其余 Sheet 提供 `directCoeff` 占位模板（空 `tierRules`），方便后续修改
- **Excel 模板新增「直接系数档位规则」行**：模板改为基于默认数据动态生成（10 张 Sheet，含 602米白卡），所有 Sheet 均含该行，格式 `最低张数-最高张数:最高系数/最低系数`，多条用 `;` 分隔；702铜版纸 已填好示例值
- **Excel 导入支持读取直接系数**：重新导入报价表时，若 Sheet 填写了「直接系数档位规则」则自动解析；未填写时匹配默认简称则继承默认配置，否则提供占位
- **Excel 导出包含直接系数**：导出时自动写入「直接系数档位规则」行，保证导入导出可往返

### 修改
- **直接系数模式不再触发纸张折后价**：直接系数模式下使用原价（不打折），报价结果隐藏「纸张折后价合计」行
- **首页默认打开直接系数计算**：`calcMode` 默认值改为 `direct`，启动时同步隐藏吊绳/邮费输入区域
- **模式切换 UI 调换**：切换按钮顺序改为「直接系数计算」在前、「标准报价」在后
- 所有文件版本号统一更新至 v6.1

---

## [v6.0] - 2026-08-13

### 新增
- **直接系数档位规则可编辑**：`DIRECT_COEFF_TIER_RULES` 从硬编码常量改为可持久化变量
- 个人主页直接系数设置区域新增「批量档位系数规则」编辑器
- 支持新增/删除/修改每条规则的张数范围和最高/最低系数
- 最大张数留空表示无上限
- 修改后实时保存 localStorage 并立即重算

### 修改
- 导出/导入备份包含 `directCoeffLevels` 和 `directCoeffTierRules`
- 「一键恢复全局默认」包含档位规则重置
- 标准报价和直接系数模式统一显示系数徽章（移除 `calcMode === "direct"` 条件）
- 所有文件版本号统一更新至 v6.0

---

## [v5.9] - 2026-08-13

### 新增
- **直接系数按批量档位自动调整**：新增 `DIRECT_COEFF_TIER_RULES` 配置
  - 1000-4000张: 最高1.6/最低1.5
  - 5000-10000张: 最高1.5/最低1.45
  - 20000张以上: 最高1.45/最低1.4
- 新增 `getDirectCoeffsForTier()` 按档位等差插值计算各级系数
- 直接系数模式报价卡片右上角显示系数徽章（×1.6 等）
- 临时系数卡片同样显示系数徽章

### 修改
- `calculate()` 函数直接系数模式改用 `getDirectCoeffsForTier(tier)` 获取系数
- 个人主页直接系数说明区域增加档位规则说明

---

## [v5.8] - 2026-08-12

### 新增
- **报价结果表格代码列手动切换**：上下箭头按钮切换代码规格
- 新增 `getPaperDisplayCodes()` 和 `findSpecByDisplayCode()` 辅助函数
- 新增 `onCodeSwitch()` 函数处理手动切换逻辑
- `sheetsState` 新增 `manualCode` 字段保留手动选择

### 修改
- `calculate()` 支持 `manualCode` 覆盖自动匹配的 spec
- 切换纸张或修改尺寸时自动清除 `manualCode` 恢复自动匹配
- 手动切换后代码显示蓝色 + 「手动」标签
- 007/008 合并处理保持一致

---

## [v5.7] - 2026-08-12

### 新增
- **直接系数计算模式**：参数输入旁新增模式切换器（标准报价 / 直接系数计算）
- 直接系数模式不计算吊绳和邮费，成本 = 纸张 + 工艺
- 个人主页新增独立直接系数等级设置（默认 1.5 / 1.45 / 1.4）
- 新增 `DEFAULT_DIRECT_COEFF_LEVELS` 和 `DIRECT_COEFF_LEVELS` 配置

### 修改
- 直接系数模式下隐藏吊绳类型、收货地区、吊绳费用、邮费行、邮费快速修改栏
- 临时毛利系数标签自动切换为「临时直接系数」
- 默认报价标签切换为「直接系数报价」

---

## [v5.6] - 2026-08-09

### 修复
- **关键Bug：`.card` / `.sheet-card` 的 `overflow: hidden` 截断下拉框**：移除容器的 overflow:hidden，这是纸张材质下拉框显示不全的根本原因
- **`.result-row` 溢出**：移除 overflow:hidden，防止长价格文本（含面积系数公式）被截断

### 响应式适配优化
- **全平台断点体系**：4 级断点覆盖所有设备
  - `≥1024px`（PC/iPad横屏）：增大间距和标题字号
  - `≤768px`（iPad竖屏）：纸张选项保持两列，表格触摸滚动
  - `≤480px`（手机）：导航字号自适应，面积系数行堆叠，价格卡片纵向
  - `≤360px`（超小屏）：表格紧凑显示，导航进一步缩小
- **组件 clamp() 流式缩放**：以下组件的固定 px 全部改为 clamp() 自适应
  - `.price-list-switcher`（报价表切换器）
  - `.area-coeff-row`（面积系数提示行）
  - `.custom-coeff-bar` / `.custom-coeff-input`（临时毛利系数栏）
  - `.shipping-override-bar` / `.shipping-override-input`（邮费快速修改栏）
  - `.price-section-label`（报价区分标签）
- **iOS 安全区域**：保留 `env(safe-area-inset-*)` 适配

---

## [v5.5.2] - 2026-08-09

### 修复
- **纸张下拉框无法滚动到底部**：`adjustDropdownPosition()` 重写为动态计算视口可用空间
  - 根据 trigger 位置实时设置 `max-height`，确保下拉框始终在视口范围内
  - 下方空间不足时自动翻转上方，并根据上方空间设置 max-height
  - 关闭下拉框时清理内联 maxHeight 样式
  - 滚轮边界判断阈值从 1px 增至 2px 防止亚像素计算问题

---

## [v5.5.1] - 2026-08-09

### 修复
- **纸张材质下拉框描述文字被截断**：移除 `.paper-desc` 的 `white-space:nowrap` / `overflow:hidden` / `text-overflow:ellipsis`
  - 改为 `word-break:break-all` + `line-height:1.4`，描述文字可换行完整显示
  - 同步移除响应式区域的截断样式
  - 仅修改纸张材质下拉框，不影响其他下拉框

---

## [v5.5] - 2026-08-09

### 重大新增
- **面积系数计算（超 10000mm²）**：面积超过 10000 时按面积系数计算价格
  - 新增 `calcAreaCoefficient()` 函数：面积 ÷ 10000，四舍五入保留两位小数
  - `matchSpec()` 超过 10000 时使用最大规格（code 100）+ 面积系数，不再报错
  - `calculate()` 中将系数应用到纸张价格（折扣前和折后价均乘以系数）
  - 举例：面积 10289 → 系数 1.03 → 80 × 1.03 = 82.4；面积 34289 → 系数 3.43 → 80 × 3.43 = 274.4

### UI 新增
- **面积系数提示行**：报价结果区新增黄色提示行，超 10000 时显示，平常隐藏
  - 显示当前面积系数值和触发原因说明
- **纸张价格显示格式更新**：触发面积系数时显示「¥ 基础价 × 系数 = ¥ 最终价」
- **个人主页新增面积系数说明卡片**：包含计算规则、举例、显示说明

### 变更
- **移除"请与上级联系"提示**：面积超 10000 不再弹出 toast，改为自动按系数计算
- **更新报价表查询页备注文案**：改为面积系数说明
- **纸张下拉框滚动修复**：添加 `overscroll-behavior:contain` + JS wheel 边界阻止

### 测试
- 更新 `test.html` 测试用例匹配新逻辑（`calcAreaCoefficient` + `matchSpec` 超限不再报错）

---

## [v5.4] - 2026-08-09

### 重大新增
- **报价表组功能**：支持多报价表管理，可切换不同报价表进行价格计算
  - 纸张按报价表隔离（通过 `priceListId` 字段），每个报价表有独立的纸张集合
  - 吊绳/邮费/客户等级为全局共享，不随报价表切换变化
  - 导航栏"1号报价表查询"重命名为"报价表组查询"
  - 报价表组查询页顶部新增报价表切换器（下拉选择 + 删除按钮）
  - 导入 Excel = 新增一个报价表，自动解析"总报价表"行作为报价表名称
  - 导出 Excel = 导出当前报价表的所有 Sheet
  - 可删除当前报价表（不允许删除最后一个）

### 数据架构变更
- `PRICE_LISTS` 数组替代单例 `PRICE_LIST_META`，支持增删报价表
- `CURRENT_PRICE_LIST_ID` 持久化到 localStorage，记录当前选中的报价表
- `PAPER_CONFIG` 每张纸增加 `priceListId` 字段，默认 9 张归 "priceList1"
- `ROPE_CONFIG` / `SHIPPING_CONFIG` 保持全局共享，移除 `priceListId` 字段
- 新增 `addPriceList(name)` / `deletePriceList(id)` 辅助函数
- `EPHEMERAL_KEYS` 为空数组，所有配置均持久化

### 导入导出改造
- `exportPaperExcel()` 仅导出当前报价表的纸张 Sheet
- `parsePaperExcel()` 解析"总报价表"行，返回 `priceListName`
- `importPaperExcel()` 导入时创建新报价表，追加纸张而非替换
- `exportFullData()` / `exportLocalBackup()` 包含 `priceLists` + `currentPriceListId`
- `importFullData()` / `importLocalBackup()` 还原完整报价表组结构
- `createSnapshot()` 快照包含报价表组结构
- 快照恢复同步还原报价表组、纸张、工艺配置

### 修复
- **ID 碰撞修复**：`parsePaperExcel()` 的 `usedIds` 现包含运行时 `PAPER_CONFIG` 所有 ID，防止多次导入生成相同 `paper_import_N`
- **空报价表守卫**：`renderSheets()` / `rebuildPaperUI()` 处理空报价表时不崩溃
- **硬编码修复**：`updateTierOptions()` 改为按当前报价表过滤；`updateDefaultPaperOptions()` 默认项显示当前报价表首张纸

---

## [v5.3] - 2026-08-09

### 修复
- **30+ Sheet 导入 ID 碰撞 Bug**：`parsePaperExcel()` 中不匹配默认简称的纸张 ID 生成逻辑存在碰撞风险
  - 旧逻辑：`id = "paper" + (sheetIndex + 1)`，与默认 `paper1`~`paper9` 冲突
  - 新逻辑：用 `usedIds` Set 去重，不匹配的纸张生成 `paper_import_N` 唯一 ID
  - 修复后 30+ Sheet 导入不再出现纸张价格不可达 / 工艺配置覆盖 / 计算错误
- **工艺 ID 生成**：改用完整 paperId 拼接（`craft_${id}_${n}`），兼容新 ID 格式

### 备份
- 创建 `backup/v5.2.1` 分支和 `v5.2.1-stable` Tag，作为报价表组功能改造前的稳定回退点

---

## [v5.2] - 2026-08-06

### 变更
- **默认客户等级毛利系数修正**：
  - 普通客户：1.3 → **1.2**
  - 优质客户：1.15（不变）
  - 大客户：1.05 → **1.1**
  - 同步更新 `DEFAULT_CUSTOMER_LEVELS`（data.js）和 HTML 静态默认值

### 新增
- **一键恢复全局默认设置**：个人主页新增「全局恢复」卡片
  - 清除浏览器中所有本地配置（纸张 / 工艺 / 吊绳 / 客户等级 / 公司信息 / 个人偏好）
  - 报价历史与本地快照不受影响
  - 操作前弹出确认框，确认后自动刷新页面
- **客户等级删除确认**：删除客户等级时增加确认弹出框
  - 显示被删除等级名称，防止误删
  - 提示可通过「恢复默认」恢复

### 修复
- **浏览器缓存问题**：JS/CSS 引用添加版本号 `?v=5.2`，强制浏览器加载最新文件
  - 修复旧版 `data.js` 被缓存导致全局恢复后系数仍为 1.3/1.15/1.05 的问题
- **v5.2 迁移逻辑**：自动检测 localStorage 中的旧默认系数（1.3/1.15/1.05）并迁移为新默认（1.2/1.15/1.1）
  - 仅迁移与旧默认完全一致的值，用户自定义系数不受影响

---

## [v5.1] - 2026-08-06

### 新增
- **临时毛利系数快速设置**：报价结果区新增临时系数输入框
  - 输入任意系数（如 1.18）后实时显示临时报价卡片（蓝色虚线边框区分）
  - 不影响默认的 3 个客户等级设置，仅当前会话临时显示
  - 空值或无效值时自动隐藏
- **邮费快速修改**：报价结果区新增邮费覆盖输入框
  - 输入修改后的邮费后，单独显示 3 个默认等级的报价卡片（绿色边框区分）
  - 基于「成本合计 - 原邮费 + 新邮费」重新计算，与原始报价区别清晰
  - 配带「清除」按钮，一键恢复
  - 与原始邮费报价区域视觉分离（标签 + 颜色区分）

### 变更
- 报价结果区增加「默认报价（原邮费）」标签，明确区分原始报价和修改邮费后报价

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
