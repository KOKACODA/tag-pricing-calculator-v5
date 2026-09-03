# KOKALabel 报价系统

**纯前端 · 多规格吊牌/标签报价计算工具**

面向吊牌/标签/不干胶印刷业务：纸张 + 工艺 + 吊绳 + 邮费 + 客户等级系数，多报价表管理，Excel 导入导出。

## 线上体验

- 最新版（v8.9.0）：https://tag-pricing-calculator-v5.pages.dev

## 功能一览

- **智能报价计算器**：多纸张、每张独立尺寸/材质/工艺，直接系数 / 标准 / 批量直接三种计价，邮费三档，吊牌成本合计
- **报价历史**：保存快照 + 输入参数，查看 / 删除 / 重新载入
- **报价表组查询**：价格表浏览 + 就地改价 + 纸张模糊搜索
- **个人主页**：报价设置 / 数据管理（备份恢复、JSON/Excel 导入导出、快照、重置）
- **离线可用**：SheetJS 已本地化，导出/导入 Excel 无需联网

## 技术栈

原生 HTML + CSS + JavaScript（无框架），SheetJS 本地化，`localStorage` 存储，Cloudflare Pages 部署。

## 分支说明

| 分支 | 版本 | 状态 |
|---|---|---|
| `v8` | v8.9.0 | **当前线上版本**，在此开发 |
| `main` | v7.10 | 旧谱系存档 |

> 两条谱系无共同提交，请勿跨分支 merge。

## 快速开始

```bash
git clone https://github.com/KOKACODA/tag-pricing-calculator-v5.git
git checkout v8
# 直接双击 index.html 即可使用，或用本地服务器：
python3 -m http.server 8080
```

## 项目结构

```
.
├── index.html                  # HTML 骨架 + CSP
├── css/style.css               # 全部样式
├── js/
│   ├── data.js                 # 数据配置 + 存储 + 版本迁移
│   ├── app.js                  # 计算 + 渲染 + 交互
│   └── vendor/xlsx.full.min.js # SheetJS（本地化）
├── tests/*.test.mjs            # Node 单元测试
├── _headers / robots.txt       # 安全头 / noindex
├── AGENTS.md                   # AI Agent 接手只读入口
├── CHANGELOG.md                # 变更日志
└── docs/
    ├── HANDOFF-v8.md           # 转手/交接文档
    ├── 项目总结.md             # 项目技术总结
    ├── main-branch-summary.md  # main 分支（v7.10 旧谱系）历史档案
    ├── 归档说明-v8.md          # v8 谱系迁移
    └── 问题日志.md             # 已知问题
```

## 验证

```bash
node --check js/app.js js/data.js   # 语法检查
node --test tests/*.test.mjs        # 单元测试
```

## 更多文档

- AI Agent 接手入口：[AGENTS.md](AGENTS.md)
- 转手/交接文档：[docs/HANDOFF-v8.md](docs/HANDOFF-v8.md)
- 技术细节：[docs/项目总结.md](docs/项目总结.md)
- main 分支历史档案：[docs/main-branch-summary.md](docs/main-branch-summary.md)
- 版本历史：[CHANGELOG.md](CHANGELOG.md)

---

测试阶段 · 搜索引擎已屏蔽（`noindex` + `robots.txt`）· © 2026 KOKALabel