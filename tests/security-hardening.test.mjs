import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const testDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(testDir, "..");

// 安全相关纯函数位于 app.js 的 DOM 渲染层（const els）之后，
// 无法用现有测试切片的纯逻辑段直接覆盖，此处按函数名精确抽取安全函数片段单独求值。
function loadSecurityHelpers() {
  const appSource = fs.readFileSync(path.join(projectRoot, "js", "app.js"), "utf8");
  const start = appSource.indexOf("function escapeHtml(text) {");
  const end = appSource.indexOf("function importLocalBackup(file) {");

  assert.notEqual(start, -1, "应能定位 escapeHtml");
  assert.notEqual(end, -1, "应能定位 importLocalBackup（安全片段结束边界）");
  assert.ok(end > start, "安全函数片段应在导入逻辑之前");

  const context = {
    localStorage: { getItem: () => null, setItem: () => {} },
    console: { info() {}, warn() {}, error() {} },
    setTimeout,
    clearTimeout,
    Promise,
    Date,
    Math
  };
  vm.createContext(context);

  const source = appSource.slice(start, end) +
    "\nglobalThis.__sec = { escapeHtml, isNonNegFinite, safeExcelText, validateImportedData, MAX_IMPORT_FILE_SIZE };";
  vm.runInContext(source, context, { filename: "security-hardening.bundle.js" });
  return context.__sec;
}

function validData() {
  return {
    kind: "local-backup",
    priceLists: [{ id: "pl1", name: "1号报价表" }],
    customerLevels: [{ id: "l1", name: "普通客户", coefficient: 1.2 }],
    paperConfig: [{
      id: "p1", name: "铜版纸", shortName: "702铜版纸", discount: 0.85,
      specs: [{ code: "55x30", maxArea: 1, prices: { "1000": 4.5 } }],
      directCoeff: { tiers: [1000], max: [1.5], min: [0.5] },
      batchDirect: { maxArea: 100, prices: { "1000": 5 } }
    }],
    ropeConfig: [{ id: "r1", name: "白色方头", prices: { "1000": 4.5 } }],
    shippingConfig: [{ id: "s1", name: "广东省内", basePrices: { "1000": 10 }, overTierCoeff: 0.5 }],
    craftConfig: { "p1": [{ id: "c1", name: "烫金", prices: { "1000": 1 } }] }
  };
}

const sec = loadSecurityHelpers();

test("escapeHtml 转义全部五个 HTML 特殊字符", () => {
  assert.equal(sec.escapeHtml(null), "");
  assert.equal(sec.escapeHtml(undefined), "");
  assert.equal(sec.escapeHtml('<script>alert("x")</script>'),
    "&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;");
  assert.equal(sec.escapeHtml("a & b"), "a &amp; b");
  assert.equal(sec.escapeHtml("'单引号'"), "&#39;单引号&#39;");
  assert.equal(sec.escapeHtml('"双引号"'), "&quot;双引号&quot;");
});

test("isNonNegFinite 仅接受非负有限数", () => {
  assert.equal(sec.isNonNegFinite(0), true);
  assert.equal(sec.isNonNegFinite(4.5), true);
  assert.equal(sec.isNonNegFinite(-1), false);
  assert.equal(sec.isNonNegFinite(-0.001), false);
  assert.equal(sec.isNonNegFinite(NaN), false);
  assert.equal(sec.isNonNegFinite(Infinity), false);
  assert.equal(sec.isNonNegFinite(-Infinity), false);
  assert.equal(sec.isNonNegFinite("5"), false);
  assert.equal(sec.isNonNegFinite(null), false);
  assert.equal(sec.isNonNegFinite(undefined), false);
});

test("safeExcelText 对公式注入前缀追加单引号", () => {
  assert.equal(sec.safeExcelText("=cmd|' /C calc'!A0"), "'=cmd|' /C calc'!A0");
  assert.equal(sec.safeExcelText("+1+2"), "'+1+2");
  assert.equal(sec.safeExcelText("-1"), "'-1");
  assert.equal(sec.safeExcelText("@SUM(A1)"), "'@SUM(A1)");
  assert.equal(sec.safeExcelText("\t制表开头"), "'\t制表开头");
  assert.equal(sec.safeExcelText("正常文本"), "正常文本");
  assert.equal(sec.safeExcelText(""), "");
  assert.equal(sec.safeExcelText(null), "");
});

test("validateImportedData 拒绝无效或损坏数据", () => {
  assert.equal(sec.validateImportedData(validData()), true);

  assert.throws(() => sec.validateImportedData(null), /数据格式无效/);
  assert.throws(() => sec.validateImportedData([]), /数据格式无效/);
  assert.throws(() => sec.validateImportedData("str"), /数据格式无效/);
  assert.throws(() => sec.validateImportedData({ kind: "paper-excel" }, "local-backup"), /文件类型不匹配/);
});

test("validateImportedData 深校验拒绝负数与 NaN", () => {
  const badDiscount = validData();
  badDiscount.paperConfig[0].discount = -1;
  assert.throws(() => sec.validateImportedData(badDiscount), /折扣系数 数值无效/);

  const nanPrice = validData();
  nanPrice.ropeConfig[0].prices["1000"] = NaN;
  assert.throws(() => sec.validateImportedData(nanPrice), /数值无效/);

  const negCoeff = validData();
  negCoeff.customerLevels[0].coefficient = 0.5;
  assert.throws(() => sec.validateImportedData(negCoeff), /客户等级系数 数值无效/);
});

test("validateImportedData 拦截超长字符串与超量数据", () => {
  const longName = validData();
  longName.priceLists[0].name = "x".repeat(500);
  assert.throws(() => sec.validateImportedData(longName), /最大长度限制/);

  const manyLevels = validData();
  manyLevels.customerLevels = Array.from({ length: 6000 }, (_, i) => ({ id: "l" + i, name: "等级", coefficient: 1 }));
  assert.throws(() => sec.validateImportedData(manyLevels), /数量超出限制/);
});

test("MAX_IMPORT_FILE_SIZE 为 50MB", () => {
  assert.equal(sec.MAX_IMPORT_FILE_SIZE, 50 * 1024 * 1024);
});