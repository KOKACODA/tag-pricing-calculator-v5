import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const testDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(testDir, "..");

function buildStatsApi(history) {
  const dataSource = fs.readFileSync(path.join(projectRoot, "js", "data.js"), "utf8");
  const appSource = fs.readFileSync(path.join(projectRoot, "js", "app.js"), "utf8");
  const boundary = appSource.indexOf("function renderHistory() {");

  assert.notEqual(boundary, -1, "应能定位 renderHistory 边界（统计函数位于其之前）");

  const store = new Map([["tagPricing_history", JSON.stringify(history)]]);
  const context = {
    localStorage: {
      getItem: key => (store.has(key) ? store.get(key) : null),
      setItem: (key, value) => { store.set(key, String(value)); }
    },
    sessionStorage: {
      getItem: () => null,
      setItem: () => {}
    },
    console: { info: () => {}, warn: () => {}, error: () => {} },
    document: {
      getElementById: () => null,
      querySelector: () => null,
      querySelectorAll: () => [],
      addEventListener: () => {},
      removeEventListener: () => {}
    },
    window: {
      addEventListener: () => {},
      removeEventListener: () => {},
      confirm: () => true,
      innerWidth: 1200,
      innerHeight: 800
    },
    escapeHtml: value => String(value ?? ""),
    setTimeout,
    clearTimeout,
    Promise,
    Date,
    Math
  };
  vm.createContext(context);

  const source = `${dataSource}\n${appSource.slice(0, boundary)}\n` +
    "globalThis.__statsApi = { parseStatSize, collectRecordSizes, computeStats, renderStatBarChart, renderProfitBars, computeProfitBrackets };";
  vm.runInContext(source, context, { filename: "stats.bundle.js" });
  return context.__statsApi;
}

test("parseStatSize 只接受正数尺寸", () => {
  const api = buildStatsApi([]);
  assert.equal(api.parseStatSize("55"), 55);
  assert.equal(api.parseStatSize(30), 30);
  assert.equal(api.parseStatSize("0"), 0);
  assert.equal(api.parseStatSize("abc"), 0);
  assert.equal(api.parseStatSize(null), 0);
});

test("collectRecordSizes 从 inputs.sheets 提取宽×长，缺失时回退 snapshot.sheetDetails", () => {
  const api = buildStatsApi([]);

  const withSheets = { inputs: { sheets: [{ width: "55", length: "30" }] } };
  const withFallback = { inputs: {}, snapshot: { sheetDetails: [{ width: "80", length: "50" }] } };
  const empty = { inputs: { sheets: [{ width: "", length: "30" }] } };

  assert.equal(JSON.stringify(api.collectRecordSizes(withSheets)), '["55×30"]');
  assert.equal(JSON.stringify(api.collectRecordSizes(withFallback)), '["80×50"]');
  assert.equal(JSON.stringify(api.collectRecordSizes(empty)), '[]');
});

test("computeStats 正确聚合月度次数、热销尺寸与利润分布", () => {
  const history = [
    { id: "h1", createdAt: "2026-09-02T10:00:00.000Z", inputs: { sheets: [{ width: "55", length: "30" }] }, snapshot: { cost: 72, pricesByLevel: [{ price: 86.4 }] } },
    { id: "h2", createdAt: "2026-09-05T10:00:00.000Z", inputs: { sheets: [{ width: "55", length: "30" }] }, snapshot: { cost: 100, pricesByLevel: [{ price: 150 }] } },
    { id: "h3", createdAt: "2026-08-20T10:00:00.000Z", inputs: { sheets: [{ width: "80", length: "50" }] }, snapshot: { cost: 200, pricesByLevel: [{ price: 180 }] } }
  ];
  const stats = buildStatsApi(history).computeStats();

  assert.equal(stats.totalCount, 3);
  assert.equal(JSON.stringify(stats.months), '[{"label":"2026-08","count":1},{"label":"2026-09","count":2}]');
  assert.equal(JSON.stringify(stats.sizes), '[{"label":"55×30","count":2},{"label":"80×50","count":1}]');

  const profits = JSON.parse(JSON.stringify(stats.profits.map(p => Math.round(p.profit * 100) / 100)));
  assert.equal(JSON.stringify(profits), '[14.4,50,-20]');
});

test("computeStats 跳过不含有效成本/报价的记录，避免污染利润", () => {
  const history = [
    { id: "h1", createdAt: "2026-09-02T10:00:00.000Z", inputs: { sheets: [] }, snapshot: { cost: 72, pricesByLevel: [{ price: 86.4 }] } },
    { id: "h2", createdAt: "2026-09-02T10:00:00.000Z", inputs: { sheets: [] }, snapshot: { cost: null, pricesByLevel: [{ price: 100 }] } },
    { id: "h3", createdAt: "2026-09-02T10:00:00.000Z", inputs: { sheets: [] }, snapshot: {} }
  ];
  const stats = buildStatsApi(history).computeStats();
  assert.equal(stats.profits.length, 1);
  assert.ok(Math.abs(stats.profits[0].profit - 14.4) < 1e-9);
});

test("renderStatBarChart 输出 HTML 并对最大值给出 100% 宽度", () => {
  const api = buildStatsApi([]);
  const html = api.renderStatBarChart([{ label: "55×30", count: 2 }, { label: "80×50", count: 1 }], "次");
  assert.match(html, /55×30/);
  assert.match(html, /2 次/);
  assert.match(html, /width:100%/);
  assert.match(html, /width:50%/);
});

test("renderProfitBars 按利润区间分桶", () => {
  const api = buildStatsApi([]);
  const profits = [{ profit: -10 }, { profit: 30 }, { profit: 60 }, { profit: 150 }, { profit: 300 }];
  const html = api.renderProfitBars(profits);
  assert.match(html, /亏损/);
  assert.match(html, /200 元以上/);
});

test("computeProfitBrackets 按区间正确分桶", () => {
  const api = buildStatsApi([]);
  const profits = [{ profit: -5 }, { profit: 0 }, { profit: 49 }, { profit: 50 }, { profit: 150 }, { profit: 300 }];
  const brackets = JSON.parse(JSON.stringify(api.computeProfitBrackets(profits)));
  assert.equal(JSON.stringify(brackets), JSON.stringify([
    { label: "亏损", count: 1 },
    { label: "0~50 元", count: 2 },
    { label: "50~100 元", count: 1 },
    { label: "100~200 元", count: 1 },
    { label: "200 元以上", count: 1 }
  ]));
});