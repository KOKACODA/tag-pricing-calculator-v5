import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const testDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(testDir, "..");

// SHA-256 与 canonicalJson 位于 data.js；signExport / verifyImportIntegrity 位于
// app.js 的 DOM 渲染层之后，故按函数名精确抽取片段并拼上 data.js 一并求值。
function loadIntegrityHelpers() {
  const dataSource = fs.readFileSync(path.join(projectRoot, "js", "data.js"), "utf8");
  const appSource = fs.readFileSync(path.join(projectRoot, "js", "app.js"), "utf8");
  const start = appSource.indexOf("function signExport(payload, kind) {");
  const end = appSource.indexOf("// -------------------- 数据管理：标签页 --------------------");

  assert.notEqual(start, -1, "应能定位 signExport");
  assert.notEqual(end, -1, "应能定位 signExport 片段结束边界");
  assert.ok(end > start, "signExport 片段应在数据管理标签页之前");

  const context = {
    localStorage: { getItem: () => null, setItem: () => {} },
    console: { info() {}, warn() {}, error() {} },
    setTimeout,
    clearTimeout,
    Promise,
    Date,
    Math,
    TextEncoder,
    Uint8Array,
    Uint32Array,
    DataView
  };
  vm.createContext(context);

  const source = `${dataSource}\n${appSource.slice(start, end)}\n` +
    "globalThis.__integrityApi = { sha256Hex, canonicalJson, signExport, verifyImportIntegrity, ACCESS_ENABLED, ACCESS_SALT };";
  vm.runInContext(source, context, { filename: "integrity.bundle.js" });
  return context.__integrityApi;
}

const api = loadIntegrityHelpers();

test("sha256Hex 与标准测试向量一致", () => {
  assert.equal(api.sha256Hex(""),
    "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
  assert.equal(api.sha256Hex("abc"),
    "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
  assert.equal(api.sha256Hex("The quick brown fox jumps over the lazy dog"),
    "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592");
  assert.equal(api.sha256Hex("中文口令123").length, 64);
});

test("canonicalJson 键排序、数组保序、忽略 undefined", () => {
  assert.equal(api.canonicalJson({ b: 1, a: 2 }), '{"a":2,"b":1}');
  assert.equal(api.canonicalJson({ a: [3, 2, 1] }), '{"a":[3,2,1]}');
  assert.equal(api.canonicalJson({ a: { c: 1, b: 2 } }), '{"a":{"b":2,"c":1}}');
  assert.equal(api.canonicalJson({ a: 1, b: undefined }), '{"a":1}');
  // 键插入顺序不影响结果
  assert.equal(api.canonicalJson({ a: 1, b: 2 }), api.canonicalJson({ b: 2, a: 1 }));
});

test("signExport 写入 64 位 sha256 完整性字段", () => {
  const signed = api.signExport({ a: 1, b: [2, 3] }, "full-config");
  assert.equal(signed.meta.algorithm, "sha256");
  assert.equal(signed.meta.kind, "full-config");
  assert.match(signed.meta.integrity, /^[0-9a-f]{64}$/);
  assert.deepEqual(signed.b, [2, 3]);
});

test("verifyImportIntegrity 正确区分 ok / legacy / tampered", () => {
  const signed = api.signExport({ a: 1, b: [2, 3] }, "local-backup");
  assert.equal(api.verifyImportIntegrity(signed), "ok");

  const tampered = JSON.parse(JSON.stringify(signed));
  tampered.a = 999;
  assert.equal(api.verifyImportIntegrity(tampered), "tampered");

  const tamperedNested = JSON.parse(JSON.stringify(signed));
  tamperedNested.b[0] = 999;
  assert.equal(api.verifyImportIntegrity(tamperedNested), "tampered");

  assert.equal(api.verifyImportIntegrity({ a: 1 }), "legacy");
  assert.equal(api.verifyImportIntegrity(null), "legacy");
  assert.equal(api.verifyImportIntegrity({ meta: { algorithm: "md5", integrity: "x" } }), "legacy");
});

test("访问门默认关闭且带盐值", () => {
  assert.equal(api.ACCESS_ENABLED, false);
  assert.equal(typeof api.ACCESS_SALT, "string");
  assert.ok(api.ACCESS_SALT.length > 0);
});