import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const testDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(testDir, "..");

function loadPureHelpers() {
  const dataSource = fs.readFileSync(path.join(projectRoot, "js", "data.js"), "utf8");
  const appSource = fs.readFileSync(path.join(projectRoot, "js", "app.js"), "utf8");
  const renderBoundary = appSource.indexOf("const els =");

  assert.notEqual(renderBoundary, -1, "应能定位 DOM 渲染层边界");

  const context = {
    localStorage: {
      getItem: () => null,
      setItem: () => {}
    },
    console: {
      info: () => {},
      warn: () => {},
      error: () => {}
    },
    setTimeout,
    clearTimeout,
    Promise,
    Date,
    Math
  };
  vm.createContext(context);

  const source = `${dataSource}\n${appSource.slice(0, renderBoundary)}\n` +
    "globalThis.__uiTestApi = { " +
    "getInitialOptionPrice, paperMatchesSearch, DEFAULT_ROPE_CONFIG, DEFAULT_PAPER_CONFIG " +
    "};";
  vm.runInContext(source, context, { filename: "ui-enhancements.bundle.js" });
  return context.__uiTestApi;
}

test("吊绳上标始终读取 1000 档初始价格", () => {
  const api = loadPureHelpers();
  const whiteSquare = api.DEFAULT_ROPE_CONFIG.find(item => item.name === "白色方头");
  const beigeCotton = api.DEFAULT_ROPE_CONFIG.find(item => item.name === "米色棉绳");

  assert.equal(api.getInitialOptionPrice(whiteSquare.prices), "4.5");
  assert.equal(api.getInitialOptionPrice(beigeCotton.prices), "9");
  assert.notEqual(api.getInitialOptionPrice(whiteSquare.prices), String(whiteSquare.prices[5000]));
});

test("纸张搜索支持简称、全称、部分文字和多关键词", () => {
  const api = loadPureHelpers();
  const paper = api.DEFAULT_PAPER_CONFIG.find(item => item.shortName === "702铜版纸");

  assert.ok(paper, "默认数据应包含 702铜版纸");
  assert.equal(api.paperMatchesSearch(paper, "702"), true);
  assert.equal(api.paperMatchesSearch(paper, "A级铜版"), true);
  assert.equal(api.paperMatchesSearch(paper, "702 0.85"), true);
  assert.equal(api.paperMatchesSearch(paper, "不存在的材质"), false);
  assert.equal(api.paperMatchesSearch(paper, ""), true);
});

