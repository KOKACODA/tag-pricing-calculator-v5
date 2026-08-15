// ============================================================
// KOKALabel报价系统 v6.5 - 主程序（计算 + 渲染 + 交互 + 初始化）
// ============================================================
"use strict";


// -------------------- P1 性能工具函数 --------------------
/**
 * 防抖：延迟执行，适合搜索输入。
 */
function debounce(fn, delay) {
  let timer = null;
  return function(...args) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

/**
 * 节流：固定频率执行，适合滚动/resize。
 */
function throttle(fn, delay) {
  let lastCall = 0;
  let timer = null;
  return function(...args) {
    const now = Date.now();
    const remaining = delay - (now - lastCall);
    if (remaining <= 0) {
      if (timer) { clearTimeout(timer); timer = null; }
      lastCall = now;
      fn.apply(this, args);
    } else if (!timer) {
      timer = setTimeout(() => {
        lastCall = Date.now();
        timer = null;
        fn.apply(this, args);
      }, remaining);
    }
  };
}

/**
 * P1.4: SheetJS 懒加载 — 仅在需要时动态加载 CDN 脚本。
 */
let _sheetjsLoaded = false;
let _sheetjsLoading = null;

function loadSheetJS() {
  if (_sheetjsLoaded) return Promise.resolve();
  if (_sheetjsLoading) return _sheetjsLoading;
  _sheetjsLoading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
    script.onload = () => { _sheetjsLoaded = true; _sheetjsLoading = null; resolve(); };
    script.onerror = () => { _sheetjsLoading = null; reject(new Error("SheetJS CDN 加载失败")); };
    document.head.appendChild(script);
  });
  return _sheetjsLoading;
}

// ============================================================
// ===================== 计算逻辑层（非必要请勿修改） =====================
// ============================================================

/**
 * 安全解析小数位数：0 有效，NaN/空值回退 2，范围 0-4。
 */
function parseDecimalPlaces(val) {
  const n = parseInt(val, 10);
  return isNaN(n) ? 2 : Math.max(0, Math.min(4, n));
}

/**
 * 金额格式化：按 APP_PROFILE.decimalPlaces 保留小数。
 */
function formatMoney(value) {
  if (typeof value !== "number" || isNaN(value)) return "-";
  return value.toFixed(parseDecimalPlaces(APP_PROFILE.decimalPlaces));
}

/**
 * 计算含出血的单张面积。
 */
function calcBleedArea(length, width) {
  return (length + 3) * (width + 3);
}

/**
 * 计算面积系数：面积超过 10000 时，系数 = 面积 / 10000，四舍五入保留两位小数。
 */
function calcAreaCoefficient(area) {
  if (area <= 10000) return 1;
  return Math.round((area / 10000) * 100) / 100;
}

/**
 * 判断纸张（Sheet）是否配置了直接系数（读取报价表表格字段）。
 * v6.2 起直接系数只读取表格：directCoeff 为 { tiers, max, min } 数组格式，无配置为 null。
 */
function paperHasDirectCoeff(paper) {
  const cfg = (paper && paper.directCoeff) || null;
  if (!cfg) return false;
  // 三行均非空且长度一致才算有效；档位与最高/最低数量不一致视为无效（提示切换标准报价）
  return Array.isArray(cfg.tiers) && cfg.tiers.length > 0 &&
    Array.isArray(cfg.max) && cfg.max.length === cfg.tiers.length &&
    Array.isArray(cfg.min) && cfg.min.length === cfg.tiers.length;
}

/**
 * 根据批量档位获取直接系数等级列表（读取报价表表格字段）。
 * 表格提供「直接系数档位 / 最高倍数 / 最低倍数」三行数据：
 *   - 找到所选批量档位对应的档位列（取小于等于所选档位的最大档位）
 *   - 最高倍数 → 普通客户（最高等级），最低倍数 → 大客户（最低等级），中间等级等差插值
 * 纸张无直接系数时返回 null（调用方按标准报价处理）。
 */
function getDirectCoeffsForTier(tier, paper) {
  if (!paperHasDirectCoeff(paper)) return null;
  const cfg = paper.directCoeff;
  const tiers = cfg.tiers;
  const maxArr = cfg.max;
  const minArr = cfg.min;

  // 找到匹配档位：取小于等于所选档位的最大档位；若所选档位小于最小档位，用最小档位
  let idx = -1;
  for (let i = 0; i < tiers.length; i++) {
    if (tier >= tiers[i]) idx = i;
  }
  if (idx < 0) idx = 0;
  const max = maxArr[idx];
  const min = minArr[idx];
  if (max == null || min == null) return null;

  const n = DIRECT_COEFF_LEVELS.length;
  if (n <= 1) return [{ ...DIRECT_COEFF_LEVELS[0], coefficient: max }];

  // 等差插值：max → min
  const step = (max - min) / (n - 1);
  return DIRECT_COEFF_LEVELS.map((l, i) => ({
    ...l,
    coefficient: Math.round((max - step * i) * 100) / 100
  }));
}

/**
 * 获取纸张的可选代码列表（007/008 合并为一项），按面积从小到大排列。
 */
function getPaperDisplayCodes(paper) {
  const seen = new Set();
  const codes = [];
  for (const spec of paper.specs) {
    const displayCode = (spec.code === "007" || spec.code === "008") ? "007/008" : spec.code;
    if (!seen.has(displayCode)) {
      seen.add(displayCode);
      codes.push(displayCode);
    }
  }
  return codes;
}

/**
 * 按显示代码查找规格（支持 007/008 合并），面积超出时附加面积系数。
 */
function findSpecByDisplayCode(paper, displayCode, area) {
  const spec = paper.specs.find(s => {
    const dc = (s.code === "007" || s.code === "008") ? "007/008" : s.code;
    return dc === displayCode;
  });
  if (!spec) return null;
  const result = { ...spec };
  if (result.code === "007" || result.code === "008") result.code = "007/008";
  if (area > result.maxArea) {
    result.areaCoefficient = calcAreaCoefficient(area);
  }
  return result;
}

/**
 * 根据有效面积向上匹配尺寸规格。
 * 面积超过 10000 时使用最大规格（code "100"）并附带面积系数；007 与 008 合并显示为 007/008。
 */
function matchSpec(paper, area) {
  const candidates = paper.specs
    .filter(s => s.maxArea >= area)
    .sort((a, b) => a.maxArea - b.maxArea);
  if (!candidates.length) {
    // 面积超过最大规格，使用最后一个规格并计算面积系数
    const lastSpec = paper.specs[paper.specs.length - 1];
    if (lastSpec && area > lastSpec.maxArea) {
      const coeff = calcAreaCoefficient(area);
      if (lastSpec.code === "007" || lastSpec.code === "008") {
        return { ...lastSpec, code: "007/008", areaCoefficient: coeff };
      }
      return { ...lastSpec, areaCoefficient: coeff };
    }
    return { error: true, message: "未找到匹配规格" };
  }
  const spec = candidates[0];
  // 007 与 008 视为同一规格
  if (spec.code === "007" || spec.code === "008") {
    return { ...spec, code: "007/008" };
  }
  return spec;
}

/**
 * 判断价格对象是否包含指定档位的精确价格。
 */
function hasExactTier(pricesObj, tier) {
  if (!pricesObj) return false;
  const v = pricesObj[tier];
  return v != null && !isNaN(Number(v));
}

/**
 * 标准计价流程：支持多纸张、每张纸独立吊牌展开尺寸、统一批量档位。
 */
function calculate(inputs) {
  const { sheetCount, tier, sheets, ropeId, regionId, mode } = inputs;
  const isDirect = mode === "direct";

  // 直接系数模式不需要 ropeId 和 regionId
  if (!sheetCount || !tier || !sheets) {
    return null;
  }
  if (sheetCount <= 0 || tier <= 0) {
    return null;
  }
  if (!Array.isArray(sheets) || sheets.length !== sheetCount) {
    return null;
  }

  // 标准模式需要 ropeId 和 regionId
  const rope = isDirect ? null : ROPE_CONFIG.find(r => r.id === ropeId);
  const region = isDirect ? null : SHIPPING_CONFIG.find(s => s.id === regionId);
  if (!isDirect && (!rope || !region)) return null;

  // 逐张纸计算：每张纸有独立的宽、长、尺寸类型
  let paperTotal = 0;
  let paperOriginalTotal = 0; // 折扣前纸张价合计
  let craftTotal = 0;
  const warnings = [];
  const sheetDetails = [];
  // 是否有任一纸张/工艺在所选档位没有定价（用于显示"无该批量定价"占位）
  let hasMissingTier = false;
  // 是否有任一纸张面积超过 10000 触发面积系数
  let hasAreaCoefficient = false;
  // 直接系数模式：以第一张纸的 Sheet 专属系数配置为准
  const primaryPaper = sheets.length
    ? getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === sheets[0].paperId)
    : null;

  for (const sheet of sheets) {
    const { width, length, sizeType, paperId, craftIds } = sheet;
    if (!width || !length || width <= 0 || length <= 0) return null;

    const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === paperId);
    if (!paper) return null;

    // 单张含出血面积（单张尺寸 / 展开尺寸目前均按输入长宽直接计算）
    const singleArea = calcBleedArea(length, width);

    // 代码选择：优先使用手动指定的代码，否则按面积自动匹配
    let spec;
    if (sheet.manualCode) {
      spec = findSpecByDisplayCode(paper, sheet.manualCode, singleArea);
    }
    if (!spec) {
      spec = matchSpec(paper, singleArea);
    }
    if (spec && spec.error) {
      return { error: spec.message };
    }

    const areaCoeff = spec.areaCoefficient || 1;
    if (areaCoeff > 1) hasAreaCoefficient = true;

    // 纸张基础价（未乘面积系数）
    const baseOriginalPrice = hasExactTier(spec.prices, tier)
      ? Number(spec.prices[tier])
      : null;
    // v6.5：标准报价模式纸张乘 discount（打折），直接系数模式纸张用原价（不打折）
    // 折扣只在标准报价触发：纸张原价 × paper.discount = 折后价，再 × 客户等级系数 = 最终报价
    // 直接系数模式：纸张原价（不打折），× 直接系数（或客户等级系数回退）= 最终报价
    const baseUnitPrice = hasExactTier(spec.prices, tier)
      ? Number(spec.prices[tier]) * (isDirect ? 1 : paper.discount)
      : null;
    // 纸张最终价（乘面积系数后）
    const paperOriginalPrice = baseOriginalPrice != null ? baseOriginalPrice * areaCoeff : null;
    const paperUnitPrice = baseUnitPrice != null ? baseUnitPrice * areaCoeff : null;

    if (paperUnitPrice == null) {
      hasMissingTier = true;
      warnings.push(`「${paper.shortName || paper.name}」无 ${tier} 张批量定价`);
    } else {
      paperTotal += paperUnitPrice;
      paperOriginalTotal += paperOriginalPrice;
    }

    // 工艺费用：每个工艺独立检查档位，无值则跳过并提示
    const crafts = CRAFT_CONFIG[sheet.paperId] || [];
    const sheetCraftDetails = [];
    if (craftIds && craftIds.length) {
      for (const cid of craftIds) {
        const craft = crafts.find(c => c.id === cid);
        if (!craft) continue;
        const cPrice = hasExactTier(craft.prices, tier)
          ? Number(craft.prices[tier])
          : null;
        if (cPrice == null) {
          hasMissingTier = true;
          warnings.push(`「${paper.shortName || paper.name}」工艺「${craft.name}」无 ${tier} 张批量定价`);
          sheetCraftDetails.push({ id: craft.id, name: craft.name, price: null, missing: true });
        } else {
          craftTotal += cPrice;
          sheetCraftDetails.push({ id: craft.id, name: craft.name, price: cPrice, missing: false });
        }
      }
    }

    sheetDetails.push({
      paperId: paper.id,
      paperName: paper.shortName || paper.name,
      width,
      length,
      sizeType,
      area: singleArea,
      code: spec.code,
      unitPrice: paperUnitPrice,
      originalUnitPrice: paperOriginalPrice,
      baseUnitPrice: baseUnitPrice,
      baseOriginalUnitPrice: baseOriginalPrice,
      areaCoefficient: areaCoeff,
      missing: paperUnitPrice == null,
      crafts: sheetCraftDetails
    });
  }

  // 吊绳费用（直接系数模式跳过）
  let ropePrice = null;
  if (!isDirect) {
    ropePrice = hasExactTier(rope.prices, tier)
      ? Number(rope.prices[tier])
      : null;
    if (ropePrice == null) {
      hasMissingTier = true;
      warnings.push(`吊绳「${rope.name}」无 ${tier} 张批量定价`);
    }
  }

  // 邮费（直接系数模式跳过）
  let shippingPrice = null;
  if (!isDirect) {
    shippingPrice = hasExactTier(region.basePrices, tier)
      ? Number(region.basePrices[tier])
      : null;
    if (shippingPrice != null && sheetDetails.length && sheetDetails[0].area < region.smallAreaThreshold) {
      shippingPrice *= region.discount;
    } else if (shippingPrice == null) {
      hasMissingTier = true;
      warnings.push(`地区「${region.name}」无 ${tier} 张批量定价`);
    }
  }

  // 成本合计：直接系数模式 = 纸张 + 工艺；标准模式 = 纸张 + 工艺 + 吊绳 + 邮费
  let cost, costKnown, costIncomplete;
  if (isDirect) {
    cost = paperTotal + craftTotal;
    costKnown = true;
    costIncomplete = false;
  } else {
    costKnown = [paperTotal, craftTotal, ropePrice, shippingPrice].every(v => v != null);
    cost = costKnown
      ? (paperTotal + craftTotal + ropePrice + shippingPrice)
      : (paperTotal + craftTotal + (ropePrice || 0) + (shippingPrice || 0));
    costIncomplete = !costKnown;
  }

  // 报价系数：直接系数模式按档位从表格读取（无直接系数的纸张回退到客户等级），标准模式用 CUSTOMER_LEVELS
  const coeffLevels = isDirect
    ? (getDirectCoeffsForTier(tier, primaryPaper) || CUSTOMER_LEVELS)
    : CUSTOMER_LEVELS;
  const pricesByLevel = coeffLevels.map(level => ({
    levelId: level.id,
    levelName: level.name,
    coefficient: level.coefficient,
    price: cost * level.coefficient
  }));

  return {
    sheetDetails,
    tier,
    paperTotal,
    paperOriginalTotal,
    craftTotal,
    ropePrice,
    shippingPrice,
    cost,
    costIncomplete,
    hasMissingTier,
    hasAreaCoefficient,
    pricesByLevel,
    warnings
  };
}

// ============================================================
// ===================== 交互渲染层（非必要请勿修改） =====================
// ============================================================

const els = {
  navBtns: document.querySelectorAll(".nav-btn"),
  pages: document.querySelectorAll(".page"),
  tier: document.getElementById("tier"),
  sheetCount: document.getElementById("sheetCount"),
  sheetList: document.getElementById("sheetList"),
  rope: document.getElementById("rope"),
  region: document.getElementById("region"),
  priceCards: document.getElementById("priceCards"),
  resSheetTable: document.getElementById("resSheetTable"),
  resTier: document.getElementById("resTier"),
  resPaperPrice: document.getElementById("resPaperPrice"),
  resPaperOriginalPrice: document.getElementById("resPaperOriginalPrice"),
  areaCoefficientRow: document.getElementById("areaCoefficientRow"),
  resAreaCoefficient: document.getElementById("resAreaCoefficient"),
  resAreaCoeffDesc: document.getElementById("resAreaCoeffDesc"),
  tierQuickSwitch: document.getElementById("tierQuickSwitch"),
  tierQuickBtns: document.getElementById("tierQuickBtns"),
  resCraftPrice: document.getElementById("resCraftPrice"),
  resRopePrice: document.getElementById("resRopePrice"),
  resShippingPrice: document.getElementById("resShippingPrice"),
  resCost: document.getElementById("resCost"),
  resWarnings: document.getElementById("resWarnings"),
  // 临时毛利系数 & 邮费快速修改
  customCoeffBar: document.getElementById("customCoeffBar"),
  customCoeffInput: document.getElementById("customCoeffInput"),
  customPriceCard: document.getElementById("customPriceCard"),
  shippingOverrideBar: document.getElementById("shippingOverrideBar"),
  shippingOverrideInput: document.getElementById("shippingOverrideInput"),
  shippingOverrideClear: document.getElementById("shippingOverrideClear"),
  shippingOverrideLabel: document.getElementById("shippingOverrideLabel"),
  shippingOverrideCards: document.getElementById("shippingOverrideCards"),
  searchInput: document.getElementById("searchInput"),
  priceTable: document.getElementById("priceTable"),
  craftTable: document.getElementById("craftTable"),
  paperDiscount: document.getElementById("paperDiscount"),
  paperDirectCoeff: document.getElementById("paperDirectCoeff"),
  tableMeta: document.getElementById("tableMeta"),
  prevPaper: document.getElementById("prevPaper"),
  nextPaper: document.getElementById("nextPaper"),
  paperPageInfo: document.getElementById("paperPageInfo"),
  paperSelector: document.getElementById("paperSelector"),
  priceListSelector: document.getElementById("priceListSelector"),
  deletePriceListBtn: document.getElementById("deletePriceListBtn"),
  paperNotes: document.getElementById("paperNotes"),
  toast: document.getElementById("toast"),
  // 扩展页元素
  levelSettings: document.getElementById("levelSettings"),
  addLevelBtn: document.getElementById("addLevelBtn"),
  resetLevelBtn: document.getElementById("resetLevelBtn"),
  companyName: document.getElementById("companyName"),
  companyPhone: document.getElementById("companyPhone"),
  defaultTier: document.getElementById("defaultTier"),
  defaultRope: document.getElementById("defaultRope"),
  defaultPaper: document.getElementById("defaultPaper"),
  decimalPlaces: document.getElementById("decimalPlaces"),
  saveProfileBtn: document.getElementById("saveProfileBtn"),
  exportProfileBtn: document.getElementById("exportProfileBtn"),
  tabs: document.querySelectorAll(".tab"),
  tabPanels: document.querySelectorAll(".tab-panel"),
  exportDataBtn: document.getElementById("exportDataBtn"),
  importDataBtn: document.getElementById("importDataBtn"),
  resetToDefaultsBtn: document.getElementById("resetToDefaultsBtn"),
  exportLocalBackupBtn: document.getElementById("exportLocalBackupBtn"),
  importLocalBackupBtn: document.getElementById("importLocalBackupBtn"),
  importLocalBackupFile: document.getElementById("importLocalBackupFile"),
  localBackupStatus: document.getElementById("localBackupStatus"),
  importFile: document.getElementById("importFile"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  importExcelBtn: document.getElementById("importExcelBtn"),
  importExcelFile: document.getElementById("importExcelFile"),
  downloadTemplateBtn: document.getElementById("downloadTemplateBtn"),
  excelImportStatus: document.getElementById("excelImportStatus"),
  exportRopeExcelBtn: document.getElementById("exportRopeExcelBtn"),
  importRopeExcelBtn: document.getElementById("importRopeExcelBtn"),
  importRopeExcelFile: document.getElementById("importRopeExcelFile"),
  downloadRopeTemplateBtn: document.getElementById("downloadRopeTemplateBtn"),
  ropeExcelImportStatus: document.getElementById("ropeExcelImportStatus"),
  exportShippingExcelBtn: document.getElementById("exportShippingExcelBtn"),
  importShippingExcelBtn: document.getElementById("importShippingExcelBtn"),
  importShippingExcelFile: document.getElementById("importShippingExcelFile"),
  downloadShippingTemplateBtn: document.getElementById("downloadShippingTemplateBtn"),
  shippingExcelImportStatus: document.getElementById("shippingExcelImportStatus"),
  snapshotTable: document.getElementById("snapshotTable"),
  snapshotEmpty: document.getElementById("snapshotEmpty"),
  createSnapshotBtn: document.getElementById("createSnapshotBtn"),
  clearSnapshotBtn: document.getElementById("clearSnapshotBtn"),
  historyTable: document.getElementById("historyTable"),
  historyEmpty: document.getElementById("historyEmpty"),
  addHistoryPlaceholderBtn: document.getElementById("addHistoryPlaceholderBtn"),
  clearHistoryBtn: document.getElementById("clearHistoryBtn"),
  // 模式切换
  modeBtns: document.querySelectorAll(".mode-btn"),
  // 报价结果中需要模式控制的行
  resRopePriceRow: null, // 稍后在 onCalculate 中动态获取
  resShippingPriceRow: null,
  // 默认报价标签
  defaultPriceLabel: document.querySelector("#priceCards") ? document.querySelector("#priceCards").previousElementSibling : null
};

let currentPaperIndex = 2;
let toastTimer = null;
let sheetsState = [];
// 计算模式：默认直接系数计算（v6.1 起），持久化到 localStorage
let calcMode = loadFromStorage("currentCalcMode", "direct"); // "standard" | "direct"

// -------------------- 初始化下拉选项 --------------------
// 根据 APP_PROFILE.defaultRope 生成吊绳单选项 HTML（顶层作用域，供 initOptions 和 rebuildRopeUI 共用）
function renderRopeRadios() {
  const fallback = "rope1"; // 默认「普通吊绳」
  const desired = APP_PROFILE.defaultRope || fallback;
  const ropes = ROPE_CONFIG;
  const hasDesired = ropes.some(r => r.id === desired);
  const checkedId = hasDesired ? desired : (ropes.find(r => r.id === fallback) ? fallback : (ropes[0] && ropes[0].id));
  return ropes.map(r => `
    <label class="rope-item">
      <input type="radio" name="rope" value="${escapeHtml(r.id)}"${r.id === checkedId ? " checked" : ""} />
      <span>${escapeHtml(r.name)}</span>
    </label>
  `).join("");
}

function initOptions() {
  // 批量档位：基于第一张纸的规格价格 keys（取并集）
  updateTierOptions(true);

  // 吊绳（三列单选）
  els.rope.innerHTML = renderRopeRadios();

  // 地区：默认选中第一项（广东省内），确保页面加载后即可计算
  if (els.region) {
    const shippingOptions = SHIPPING_CONFIG;
    els.region.innerHTML = '<option value="">请选择地区</option>' +
      shippingOptions.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
    if (shippingOptions.length > 0) {
      els.region.value = shippingOptions[0].id;
    }
  }
}

/**
 * 更新批量档位下拉选项。
 * @param {boolean} isInit 是否为初始化，若是则默认选中第一项。
 */
function updateTierOptions(isInit) {
  // 批量档位取当前报价表所有规格的价格档位并集，避免切换纸张时档位缺失
  const tierSet = new Set();
  for (const paper of getPapersByPriceList(CURRENT_PRICE_LIST_ID)) {
    for (const spec of paper.specs) {
      Object.keys(spec.prices).forEach(t => tierSet.add(Number(t)));
    }
  }
  const tiers = Array.from(tierSet).sort((a, b) => a - b);

  const currentTier = els.tier.value;
  els.tier.innerHTML = '<option value="">请选择批量档位</option>' +
    tiers.map(t => `<option value="${t}">${t} 张</option>`).join("");

  if (isInit) {
    if (tiers.includes(1000)) {
      els.tier.value = 1000;
    } else if (tiers.length) {
      els.tier.value = tiers[0];
    }
  } else if (currentTier) {
    // 尝试保留用户选择意图，若不可用则选最接近档位
    if (tiers.includes(Number(currentTier))) {
      els.tier.value = currentTier;
    } else {
      const fallback = tiers.find(t => t >= Number(currentTier));
      els.tier.value = fallback !== undefined ? fallback : (tiers[tiers.length - 1] || "");
    }
  }
}

// -------------------- 纸张设置卡片渲染 --------------------
function renderSheets() {
  const count = parseInt(els.sheetCount.value, 10) || 2;
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  const defaultPaperId = currentPapers[0]?.id || "";
  const newSheets = [];

  for (let i = 0; i < count; i++) {
    const old = sheetsState[i];
    newSheets.push({
      paperId: old && old.paperId ? old.paperId : defaultPaperId,
      craftIds: old && old.craftIds ? old.craftIds.slice() : [],
      width: old && old.width ? old.width : "55",
      length: old && old.length ? old.length : "30",
      sizeType: old && old.sizeType ? old.sizeType : "single",
      manualCode: old && old.manualCode != null ? old.manualCode : null
    });
  }
  sheetsState = newSheets;

  els.sheetList.innerHTML = "";
  sheetsState.forEach((sheet, index) => {
    const card = document.createElement("div");
    card.className = "sheet-card";

    const currentPaper = currentPapers.find(p => p.id === sheet.paperId) || currentPapers[0];
    // Sheet 数量 > 10 时显示编号前缀，方便快速定位
    const showPaperIndex = currentPapers.length > 10;
    const paperOptions = currentPapers.map((p, pIdx) => `
      <div class="paper-option${p.id === sheet.paperId ? " active" : ""}" data-sheet="${index}" data-paper="${escapeHtml(p.id)}">
        <span class="paper-name">${showPaperIndex ? (pIdx + 1) + ". " : ""}${escapeHtml(p.shortName || p.name)}</span>
        <span class="paper-desc">${escapeHtml(p.name)}</span>
      </div>
    `).join("");

    const crafts = CRAFT_CONFIG[sheet.paperId] || [];
    const craftHtml = crafts.length
      ? crafts.map(c => `
          <label class="craft-item">
            <input type="checkbox" value="${c.id}" data-sheet="${index}" data-craft="${c.id}"${sheet.craftIds.includes(c.id) ? " checked" : ""} />
            <span>${escapeHtml(c.name)}</span>
          </label>
        `).join("")
      : '<div class="craft-empty">该纸张暂无附加工艺</div>';

    const triggerText = currentPaper ? (currentPaper.shortName || currentPaper.name) : "无可用纸张";
    const triggerDesc = currentPaper ? currentPaper.name : "请导入报价表";
    // v6.2：无直接系数的纸张（directCoeff 缺失或最高/最低倍数为空）提示切换标准报价
    const noDirectCoeff = !paperHasDirectCoeff(currentPaper);

    card.innerHTML = `
      <div class="sheet-title">纸张 ${index + 1}</div>
      <div class="form-group">
        <label>纸张材质 <span class="hint">点击展开，选择后自动收起</span></label>
        <div class="paper-dropdown" data-sheet="${index}">
          <div class="paper-trigger">
            <div>
              <span class="paper-trigger-text">${escapeHtml(triggerText)}</span>
              <span class="paper-trigger-desc">${escapeHtml(triggerDesc)}</span>
            </div>
            <span class="paper-trigger-arrow"></span>
          </div>
          <div class="paper-options">
            ${paperOptions}
          </div>
        </div>
        ${noDirectCoeff ? `
        <div class="sheet-direct-warning">
          <span class="sheet-direct-warning-icon">⚠</span>
          <span>该纸张无直接系数（最高/最低倍数为空），请切换为<strong>标准报价</strong>计算</span>
        </div>` : ""}
      </div>
      <div class="form-group" style="margin-bottom: 0;">
        <label>吊牌展开尺寸 <span class="hint">自动加 3mm 出血</span></label>
        <div class="sheet-size-row">
          <div class="form-group">
            <label class="unit">宽 (mm)</label>
            <input type="number" class="sheet-width" data-sheet="${index}" min="1" step="1" placeholder="55" value="${sheet.width}" />
          </div>
          <div class="form-group">
            <label class="unit">长 (mm)</label>
            <input type="number" class="sheet-length" data-sheet="${index}" min="1" step="1" placeholder="30" value="${sheet.length}" />
          </div>
          <div class="form-group size-type-group">
            <label class="unit">类型</label>
            <select class="sheet-size-type" data-sheet="${index}">
              <option value="single"${sheet.sizeType === "single" ? " selected" : ""}>单张尺寸</option>
              <option value="spread"${sheet.sizeType === "spread" ? " selected" : ""}>展开尺寸</option>
            </select>
          </div>
        </div>
      </div>
      <div class="form-group" style="margin-bottom: 0; margin-top: 12px;">
        <label>附加工艺</label>
        <div class="craft-list">${craftHtml}</div>
      </div>
    `;

    els.sheetList.appendChild(card);
  });

  // 绑定纸张下拉交互
  els.sheetList.querySelectorAll(".paper-dropdown").forEach(dd => {
    const trigger = dd.querySelector(".paper-trigger");
    trigger.addEventListener("click", e => {
      e.stopPropagation();
      togglePaperDropdown(dd);
    });
  });
  els.sheetList.querySelectorAll(".paper-option").forEach(opt => {
    opt.addEventListener("click", e => {
      e.stopPropagation();
      onPaperChange(opt);
    });
  });
  els.sheetList.querySelectorAll("input[type=checkbox][data-craft]").forEach(cb => {
    cb.addEventListener("change", onCraftChange);
  });
  // 绑定每张纸的尺寸输入（input 元素同时绑定 input+change，select 元素仅绑定 change 避免双触发）
  els.sheetList.querySelectorAll(".sheet-width, .sheet-length").forEach(input => {
    input.addEventListener("input", onSheetSizeChange);
    input.addEventListener("change", onSheetSizeChange);
  });
  els.sheetList.querySelectorAll(".sheet-size-type").forEach(sel => {
    sel.addEventListener("change", onSheetSizeChange);
  });

  // 若第一张纸发生变化，需要更新档位选项
  updateTierOptions(false);
}

function togglePaperDropdown(dropdown) {
  const isOpen = dropdown.classList.contains("open");
  // 先关闭所有下拉并清理定位类
  closeAllPaperDropdowns();
  if (!isOpen) {
    dropdown.classList.add("open");
    // 下一帧测量弹窗位置，避免 display:none 状态下 getBoundingClientRect 返回 0
    requestAnimationFrame(() => adjustDropdownPosition(dropdown));
  }
}

/**
 * 视口边界检测：弹窗自动避让屏幕边缘
 * - 动态设置 max-height 确保下拉框始终在视口内可完整滚动
 * - 下方空间不足时向上翻转（flip-up）
 * - 右侧溢出时左移（shift-left）
 * - 超小屏切换为全屏固定定位（full-width）
 */
function adjustDropdownPosition(dropdown) {
  const options = dropdown.querySelector(".paper-options");
  if (!options) return;

  // 清除上一次的定位类和内联 max-height
  options.classList.remove("flip-up", "shift-left", "full-width");
  options.style.maxHeight = "";

  const trigger = dropdown.querySelector(".paper-trigger");
  const triggerRect = trigger.getBoundingClientRect();
  const vw = window.innerWidth;
  const vh = window.innerHeight;

  // 计算上方和下方的可用空间（留 12px 安全边距）
  const GAP = 6; // trigger 与 options 之间的间距
  const SAFE_MARGIN = 12;
  const spaceBelow = vh - triggerRect.bottom - GAP - SAFE_MARGIN;
  const spaceAbove = triggerRect.top - GAP - SAFE_MARGIN;

  // 选择空间更大的方向，并动态设置 max-height
  if (spaceBelow >= spaceAbove || spaceBelow >= 200) {
    // 下方打开
    options.style.maxHeight = Math.min(spaceBelow, 400) + "px";
  } else {
    // 上方打开
    options.classList.add("flip-up");
    options.style.maxHeight = Math.min(spaceAbove, 400) + "px";
  }

  // 2. 水平方向：判断右侧是否溢出
  const optionsRect = options.getBoundingClientRect();
  if (optionsRect.right > vw - 8) {
    options.classList.add("shift-left");
  }

  // 3. 超小屏：弹窗宽度超出视口时使用固定全宽定位
  if (triggerRect.width < vw * 0.6 && vw <= 480) {
    options.classList.add("full-width");
    options.style.maxHeight = Math.min(Math.max(spaceBelow, spaceAbove), 400) + "px";
  }
}

function closeAllPaperDropdowns() {
  document.querySelectorAll(".paper-dropdown.open").forEach(dd => {
    dd.classList.remove("open");
    const opt = dd.querySelector(".paper-options");
    if (opt) {
      opt.classList.remove("flip-up", "shift-left", "full-width");
      opt.style.maxHeight = "";
    }
  });
}

// 滚动时关闭所有打开的下拉（避免定位错乱，同时提升移动端体验）
window.addEventListener("scroll", throttle(() => {
  closeAllPaperDropdowns();
}, 100), { passive: true });

// 窗口尺寸变化时重新检测已打开的下拉位置
let resizeTimer = null;
window.addEventListener("resize", () => {
  if (resizeTimer) clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    document.querySelectorAll(".paper-dropdown.open").forEach(dd => adjustDropdownPosition(dd));
  }, 150);
}, { passive: true });

// 屏幕旋转时关闭下拉
window.addEventListener("orientationchange", () => {
  closeAllPaperDropdowns();
});

function onPaperChange(optionEl) {
  const index = Number(optionEl.dataset.sheet);
  const newPaperId = optionEl.dataset.paper;
  sheetsState[index].paperId = newPaperId;
  // 切换纸张时清空工艺选择和手动代码
  sheetsState[index].craftIds = [];
  sheetsState[index].manualCode = null;
  closeAllPaperDropdowns();
  renderSheets();
  onCalculate();
}

function onCraftChange(e) {
  const index = Number(e.target.dataset.sheet);
  const craftId = e.target.value;
  const checked = e.target.checked;
  const craftIds = sheetsState[index].craftIds;
  if (checked) {
    if (!craftIds.includes(craftId)) craftIds.push(craftId);
  } else {
    sheetsState[index].craftIds = craftIds.filter(id => id !== craftId);
  }
  onCalculate();
}

function onSheetSizeChange(e) {
  const index = Number(e.target.dataset.sheet);
  const input = e.target;
  if (input.classList.contains("sheet-width")) {
    sheetsState[index].width = input.value;
  } else if (input.classList.contains("sheet-length")) {
    sheetsState[index].length = input.value;
  } else if (input.classList.contains("sheet-size-type")) {
    sheetsState[index].sizeType = input.value;
  }
  // 尺寸变化时清除手动代码，恢复自动匹配
  sheetsState[index].manualCode = null;
  onCalculate();
}

/**
 * 手动切换纸张代码（上/下），覆盖自动匹配结果。
 */
function onCodeSwitch(sheetIdx, dir) {
  const state = sheetsState[sheetIdx];
  if (!state) return;
  const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === state.paperId);
  if (!paper) return;
  const codes = getPaperDisplayCodes(paper);
  if (codes.length <= 1) return;

  // 确定当前代码
  let currentCode = state.manualCode;
  if (!currentCode) {
    const w = parseFloat(state.width);
    const l = parseFloat(state.length);
    if (w && l) {
      const area = calcBleedArea(l, w);
      const spec = matchSpec(paper, area);
      currentCode = spec.code;
    }
  }
  const currentIdx = codes.indexOf(currentCode);
  if (currentIdx < 0) return;

  let newIdx;
  if (dir === "up") {
    newIdx = Math.max(0, currentIdx - 1);
  } else {
    newIdx = Math.min(codes.length - 1, currentIdx + 1);
  }
  if (newIdx === currentIdx) return;

  state.manualCode = codes[newIdx];
  onCalculate();
}

// -------------------- 实时计算 --------------------
function onCalculate() {
  const sheetCount = parseInt(els.sheetCount.value, 10);
  const tier = parseInt(els.tier.value, 10);
  const ropeId = els.rope.querySelector('input[name="rope"]:checked')?.value || "";
  const regionId = els.region.value;

  // 直接系数模式不需要 ropeId 和 regionId
  if (!sheetCount || !tier || sheetCount <= 0 || tier <= 0) {
    clearResult();
    return;
  }
  if (calcMode === "standard" && (!ropeId || !regionId)) {
    clearResult();
    return;
  }

  // 收集每张纸的纸材、工艺、尺寸
  const sheets = sheetsState.map(s => ({
    paperId: s.paperId,
    craftIds: s.craftIds.slice(),
    width: parseFloat(s.width),
    length: parseFloat(s.length),
    sizeType: s.sizeType,
    manualCode: s.manualCode || null
  }));

  // 只要有任何一张纸缺少有效尺寸就清空结果
  if (sheets.some(s => !s.width || !s.length || s.width <= 0 || s.length <= 0)) {
    clearResult();
    return;
  }

  const result = calculate({ sheetCount, tier, sheets, ropeId, regionId, mode: calcMode });

  if (!result) {
    clearResult();
    return;
  }

  if (result.error) {
    clearResult();
    showToast(result.error);
    return;
  }

  // 渲染各纸张尺寸/面积/代码
  if (els.resSheetTable) {
    const tbody = els.resSheetTable.querySelector("tbody");
    tbody.innerHTML = result.sheetDetails.map((s, idx) => {
      // 获取该纸张可选代码列表
      const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === s.paperId);
      const codes = paper ? getPaperDisplayCodes(paper) : [];
      const currentCodeIdx = codes.indexOf(s.code);
      const canGoUp = currentCodeIdx > 0;
      const canGoDown = currentCodeIdx >= 0 && currentCodeIdx < codes.length - 1;
      const isManual = sheetsState[idx] && sheetsState[idx].manualCode != null;
      return `
      <tr>
        <td>${escapeHtml(s.paperName)}</td>
        <td>${s.width} × ${s.length}${s.sizeType === "spread" ? "（展开）" : ""}</td>
        <td>${s.area} mm²</td>
        <td>
          <div class="code-switcher${isManual ? " manual" : ""}">
            <span class="code-text">${escapeHtml(s.code)}${isManual ? '<span class="code-manual-tag">手动</span>' : ''}</span>
            <div class="code-arrows">
              <button class="code-arrow up${!canGoUp ? ' disabled' : ''}" data-sheet-idx="${idx}" data-dir="up" ${!canGoUp ? 'disabled' : ''} title="切换到上一档代码">&#9650;</button>
              <button class="code-arrow down${!canGoDown ? ' disabled' : ''}" data-sheet-idx="${idx}" data-dir="down" ${!canGoDown ? 'disabled' : ''} title="切换到下一档代码">&#9660;</button>
            </div>
          </div>
        </td>
        <td>${s.missing ? '<span class="price-missing">无该批量定价</span>' : '¥ ' + formatMoney(s.unitPrice)}</td>
      </tr>
    `;}).join("");
    // 绑定代码切换按钮事件
    tbody.querySelectorAll(".code-arrow:not(.disabled)").forEach(btn => {
      btn.addEventListener("click", () => {
        onCodeSwitch(Number(btn.dataset.sheetIdx), btn.dataset.dir);
      });
    });
  }

  els.resTier.textContent = result.tier + " 张";
  // 纸张价合计（折扣前）/ 纸张折后价合计：多纸张时显示 "¥价格1 + ¥价格2 = ¥合计"，合计蓝字
  // 当面积系数 > 1 时，每张纸显示 "¥基础价 × 系数" 格式
  function renderPaperTotal(getFinalPrice, getBasePrice) {
    const details = result.sheetDetails;
    if (details.length === 1) {
      const p = getFinalPrice(details[0]);
      if (p == null) return '<span class="price-missing">无该批量定价</span>';
      if (details[0].areaCoefficient > 1) {
        const bp = getBasePrice(details[0]);
        const coeff = details[0].areaCoefficient;
        return '¥ ' + formatMoney(bp) + ' × ' + coeff + ' = <span style="color:var(--brand);font-weight:600;">¥ ' + formatMoney(p) + '</span>';
      }
      return "¥ " + formatMoney(p);
    }
    // 多纸张：逐张价格 + = 合计
    const parts = details.map(s => {
      const p = getFinalPrice(s);
      if (p == null) return '<span class="price-missing">缺价</span>';
      if (s.areaCoefficient > 1) {
        const bp = getBasePrice(s);
        return '¥ ' + formatMoney(bp) + ' × ' + s.areaCoefficient;
      }
      return "¥ " + formatMoney(p);
    });
    const total = getFinalPrice(details[0]) != null && details.every(s => getFinalPrice(s) != null);
    const totalStr = total
      ? '<span style="color:var(--brand);font-weight:600;">¥ ' + formatMoney(details.reduce((sum, s) => sum + getFinalPrice(s), 0)) + '</span>'
      : '<span class="price-missing">部分缺价</span>';
    return parts.join(' + ') + ' = ' + totalStr;
  }
  els.resPaperOriginalPrice.innerHTML = renderPaperTotal(s => s.originalUnitPrice, s => s.baseOriginalUnitPrice);
  els.resPaperPrice.innerHTML = renderPaperTotal(s => s.unitPrice, s => s.baseUnitPrice);
  els.resCraftPrice.innerHTML = result.craftTotal ? "¥ " + formatMoney(result.craftTotal) : '<span class="price-missing">无该批量定价</span>';

  // 吊绳/邮费行：直接系数模式隐藏
  const ropeRow = els.resRopePrice ? els.resRopePrice.closest(".result-row") : null;
  const shippingRow = els.resShippingPrice ? els.resShippingPrice.closest(".result-row") : null;
  // 纸张折后价合计行：直接系数模式不使用折扣，隐藏该行
  const paperDiscountRow = els.resPaperPrice ? els.resPaperPrice.closest(".result-row") : null;
  if (calcMode === "direct") {
    if (ropeRow) ropeRow.style.display = "none";
    if (shippingRow) shippingRow.style.display = "none";
    if (paperDiscountRow) paperDiscountRow.style.display = "none";
  } else {
    if (ropeRow) ropeRow.style.display = "";
    if (shippingRow) shippingRow.style.display = "";
    if (paperDiscountRow) paperDiscountRow.style.display = "";
    els.resRopePrice.innerHTML = result.ropePrice != null ? "¥ " + formatMoney(result.ropePrice) : '<span class="price-missing">无该批量定价</span>';
    els.resShippingPrice.innerHTML = result.shippingPrice != null ? "¥ " + formatMoney(result.shippingPrice) : '<span class="price-missing">无该批量定价</span>';
  }
  els.resCost.innerHTML = result.costIncomplete
    ? '<span class="price-missing">部分缺价</span>'
    : "¥ " + formatMoney(result.cost);

  // 渲染档位缺失警告
  if (result.warnings && result.warnings.length && els.resWarnings) {
    els.resWarnings.innerHTML = result.warnings.map(w => "• " + escapeHtml(w)).join("<br>");
    els.resWarnings.style.display = "block";
  } else if (els.resWarnings) {
    els.resWarnings.innerHTML = "";
    els.resWarnings.style.display = "none";
  }

  // 面积系数提示：超 10000 时显示，否则隐藏
  if (els.areaCoefficientRow) {
    if (result.hasAreaCoefficient) {
      // 收集所有有面积系数的纸张信息
      const coeffSheets = result.sheetDetails.filter(s => s.areaCoefficient > 1);
      const coeffList = coeffSheets.map(s => `${s.paperName}: ${s.area}`).join('，');
      if (els.resAreaCoefficient) {
        els.resAreaCoefficient.textContent = coeffSheets.map(s => s.areaCoefficient).filter((v, i, a) => a.indexOf(v) === i).join(' / ');
      }
      if (els.resAreaCoeffDesc) {
        els.resAreaCoeffDesc.textContent = `计算面积 ${coeffList} 超过 10000mm²，已按面积系数计算`;
      }
      els.areaCoefficientRow.style.display = "flex";
    } else {
      els.areaCoefficientRow.style.display = "none";
    }
  }

  // 渲染三个客户等级报价卡片（右上角显示系数徽章）
  els.priceCards.innerHTML = result.pricesByLevel.map((item, idx) => `
    <div class="price-card${idx === 0 ? " highlight" : ""}">
      <span class="coeff-badge">×${item.coefficient}</span>
      <div class="level-name">${escapeHtml(item.levelName)}</div>
      <div class="level-price">${result.costIncomplete
        ? '<span class="price-missing">部分缺价</span>'
        : formatMoney(item.price) + '<span class="unit">元</span>'}</div>
    </div>
  `).join("");

  // 渲染批量档位快速切换按钮
  if (els.tierQuickSwitch && els.tierQuickBtns) {
    const tierSet = new Set();
    for (const paper of getPapersByPriceList(CURRENT_PRICE_LIST_ID)) {
      for (const spec of paper.specs) {
        Object.keys(spec.prices).forEach(t => tierSet.add(Number(t)));
      }
    }
    const allTiers = Array.from(tierSet).sort((a, b) => a - b);
    els.tierQuickBtns.innerHTML = allTiers.map(t =>
      `<button class="tier-quick-btn${t === result.tier ? " active" : ""}" data-tier="${t}">${t}</button>`
    ).join("");
    els.tierQuickBtns.querySelectorAll(".tier-quick-btn").forEach(btn => {
      btn.addEventListener("click", function() {
        els.tier.value = this.dataset.tier;
        onCalculate();
      });
    });
    els.tierQuickSwitch.style.display = "flex";
  }

  // -------------------- 临时系数 & 邮费快速修改 --------------------
  // 缓存本次计算结果供临时系数/邮费修改使用
  _lastResult = result;

  // 临时系数输入栏：直接系数模式下标签改为"临时直接系数"
  if (els.customCoeffBar) {
    els.customCoeffBar.style.display = "flex";
    const labelEl = els.customCoeffBar.querySelector(".custom-coeff-label");
    if (labelEl) {
      labelEl.textContent = calcMode === "direct" ? "临时直接系数：" : "临时毛利系数：";
    }
    renderCustomCoeffCard();
  }

  // 邮费快速修改栏：直接系数模式隐藏
  if (els.shippingOverrideBar) {
    if (calcMode === "direct") {
      els.shippingOverrideBar.style.display = "none";
      if (els.shippingOverrideInput) els.shippingOverrideInput.value = "";
      if (els.shippingOverrideCards) { els.shippingOverrideCards.style.display = "none"; els.shippingOverrideCards.innerHTML = ""; }
      if (els.shippingOverrideLabel) els.shippingOverrideLabel.style.display = "none";
    } else {
      els.shippingOverrideBar.style.display = "flex";
      renderShippingOverrideCards();
    }
  }

  // 默认报价标签：直接系数模式改为"直接系数报价"
  if (els.defaultPriceLabel) {
    els.defaultPriceLabel.textContent = calcMode === "direct" ? "直接系数报价" : "默认报价（原邮费）";
  }
}

// 缓存上一次计算结果
let _lastResult = null;

/**
 * 渲染临时毛利系数报价卡片
 */
function renderCustomCoeffCard() {
  if (!els.customPriceCard || !_lastResult) return;
  const raw = els.customCoeffInput ? els.customCoeffInput.value.trim() : "";
  const coeff = parseFloat(raw);
  if (!raw || isNaN(coeff) || coeff < 1) {
    els.customPriceCard.style.display = "none";
    els.customPriceCard.innerHTML = "";
    return;
  }
  const price = _lastResult.cost * coeff;
  els.customPriceCard.innerHTML = `
    <div class="price-card custom-coeff">
      <span class="coeff-badge">×${coeff}</span>
      <div class="level-name">临时系数 ${coeff}</div>
      <div class="level-price">${_lastResult.costIncomplete
        ? '<span class="price-missing">部分缺价</span>'
        : formatMoney(price) + '<span class="unit">元</span>'}</div>
    </div>
  `;
  els.customPriceCard.style.display = "flex";
}

/**
 * 渲染邮费修改后的报价卡片（3 个默认等级，基于修改后的邮费重算）
 */
function renderShippingOverrideCards() {
  if (!els.shippingOverrideCards || !_lastResult) return;
  const raw = els.shippingOverrideInput ? els.shippingOverrideInput.value.trim() : "";
  const newShipping = parseFloat(raw);
  // 没输入或无效值时隐藏
  if (!raw || isNaN(newShipping) || newShipping < 0) {
    els.shippingOverrideCards.style.display = "none";
    els.shippingOverrideLabel.style.display = "none";
    els.shippingOverrideCards.innerHTML = "";
    return;
  }
  // 用修改后的邮费重算成本
  const origShipping = _lastResult.shippingPrice || 0;
  const newCost = _lastResult.cost - origShipping + newShipping;
  const costIncomplete = _lastResult.costIncomplete;
  // 渲染 3 个默认等级报价卡片
  els.shippingOverrideCards.innerHTML = _lastResult.pricesByLevel.map((item, idx) => {
    const newPrice = newCost * item.coefficient;
    return `
      <div class="price-card${idx === 0 ? " highlight" : ""}">
        <div class="level-name">${escapeHtml(item.levelName)}</div>
        <div class="level-price">${costIncomplete
          ? '<span class="price-missing">部分缺价</span>'
          : formatMoney(newPrice) + '<span class="unit">元</span>'}</div>
      </div>
    `;
  }).join("");
  els.shippingOverrideLabel.style.display = "block";
  els.shippingOverrideCards.style.display = "flex";
}

function clearResult() {
  if (els.resSheetTable) {
    els.resSheetTable.querySelector("tbody").innerHTML = "";
  }
  els.resTier.textContent = "-";
  els.resPaperPrice.textContent = "-";
  els.resPaperOriginalPrice.textContent = "-";
  els.resCraftPrice.textContent = "-";
  // 吊绳/邮费行在直接系数模式下保持隐藏
  const ropeRow = els.resRopePrice ? els.resRopePrice.closest(".result-row") : null;
  const shippingRow = els.resShippingPrice ? els.resShippingPrice.closest(".result-row") : null;
  const paperDiscountRow = els.resPaperPrice ? els.resPaperPrice.closest(".result-row") : null;
  if (calcMode === "direct") {
    if (ropeRow) ropeRow.style.display = "none";
    if (shippingRow) shippingRow.style.display = "none";
    if (paperDiscountRow) paperDiscountRow.style.display = "none";
  } else {
    if (ropeRow) ropeRow.style.display = "";
    if (shippingRow) shippingRow.style.display = "";
    if (paperDiscountRow) paperDiscountRow.style.display = "";
    els.resRopePrice.textContent = "-";
    els.resShippingPrice.textContent = "-";
  }
  els.resCost.textContent = "-";
  if (els.resWarnings) {
    els.resWarnings.innerHTML = "";
    els.resWarnings.style.display = "none";
  }
  if (els.areaCoefficientRow) {
    els.areaCoefficientRow.style.display = "none";
  }
  if (els.tierQuickSwitch) {
    els.tierQuickSwitch.style.display = "none";
  }
  els.priceCards.innerHTML = "";
  // 清理临时毛利系数 & 邮费快速修改
  _lastResult = null;
  if (els.customCoeffBar) els.customCoeffBar.style.display = "none";
  if (els.customPriceCard) { els.customPriceCard.style.display = "none"; els.customPriceCard.innerHTML = ""; }
  if (els.shippingOverrideBar) els.shippingOverrideBar.style.display = "none";
  if (els.shippingOverrideInput) els.shippingOverrideInput.value = "";
  if (els.shippingOverrideLabel) els.shippingOverrideLabel.style.display = "none";
  if (els.shippingOverrideCards) { els.shippingOverrideCards.style.display = "none"; els.shippingOverrideCards.innerHTML = ""; }
}

// -------------------- 报价表渲染 --------------------
function renderPriceTable() {
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  const paper = currentPapers[currentPaperIndex];
  if (!paper) return;

  const keyword = els.searchInput.value.trim().toLowerCase();
  const filtered = paper.specs.filter(s =>
    s.code.toLowerCase().includes(keyword) ||
    String(s.maxArea).includes(keyword)
  );

  // 动态表头：取第一个规格的价格 keys 作为档位列
  const tiers = paper.specs.length ? Object.keys(paper.specs[0].prices).map(Number).sort((a, b) => a - b) : [];

  const theadRow = els.priceTable.querySelector("thead tr");
  theadRow.innerHTML = '<th>代码</th><th>最大含出血面积 (mm²)</th>' +
    tiers.map(t => `<th>${t} 张</th>`).join("");

  const tbody = els.priceTable.querySelector("tbody");
  tbody.innerHTML = filtered.map(s => `
    <tr>
      <td>${escapeHtml(s.code)}</td>
      <td>${s.maxArea}</td>
      ${tiers.map(t => {
        const v = s.prices[t];
        if (v == null) return '<td><span class="price-missing">无该批量定价</span></td>';
        return `<td>¥ ${formatMoney(Number(v))}</td>`;
      }).join("")}
    </tr>
  `).join("");

  els.paperDiscount.textContent = `${paper.name} | ${paper.discount === 1 ? "无折扣" : (paper.discount * 10).toFixed(1) + "折"}`;
  // 直接系数显示（统一用 paperHasDirectCoeff 判断，确保三行完整且数量一致）
  if (els.paperDirectCoeff) {
    if (paperHasDirectCoeff(paper)) {
      const dc = paper.directCoeff;
      const dcInfo = dc.tiers.map((t, i) => `${t}张:×${dc.max[i]}/×${dc.min[i]}`).join("  ");
      els.paperDirectCoeff.textContent = `直接系数档位：${dcInfo}`;
      els.paperDirectCoeff.style.display = "inline";
    } else {
      els.paperDirectCoeff.textContent = "";
      els.paperDirectCoeff.style.display = "none";
    }
  }
  els.tableMeta.textContent = `行数：${filtered.length} / ${paper.specs.length}`;
  els.paperPageInfo.textContent = `第 ${currentPaperIndex + 1} / ${currentPapers.length} 张`;
  els.prevPaper.disabled = currentPaperIndex === 0;
  els.nextPaper.disabled = currentPaperIndex === currentPapers.length - 1;
  els.paperNotes.textContent = `备注：当前展示「${paper.name}」（${getCurrentPriceList().name}）。007 与 008 合并为 007/008；面积超过 10000 mm² 时按面积系数计算（面积÷10000，四舍五入保留两位小数）。`;

  // 同步下拉选择器
  if (els.paperSelector) {
    syncPaperSelector();
  }

  // 渲染工艺价格表
  const crafts = CRAFT_CONFIG[paper.id] || [];
  const craftTheadRow = els.craftTable.querySelector("thead tr");
  craftTheadRow.innerHTML = '<th>工艺名称</th>' + tiers.map(t => `<th>${t} 张</th>`).join("");
  const craftTbody = els.craftTable.querySelector("tbody");
  craftTbody.innerHTML = crafts.length
    ? crafts.map(c => `
        <tr>
          <td>${escapeHtml(c.name)}</td>
          ${tiers.map(t => {
            const v = c.prices[t];
            if (v == null) return '<td><span class="price-missing">无该批量定价</span></td>';
            return `<td>¥ ${formatMoney(Number(v))}</td>`;
          }).join("")}
        </tr>
      `).join("")
    : '<tr><td colspan="' + (tiers.length + 1) + '" style="text-align:center;color:var(--text-secondary);">该纸张暂无工艺配置</td></tr>';
}

// -------------------- Sheet 纸张选择器 --------------------
/**
 * 同步下拉选择器：填充选项并选中当前纸张
 * 当 Sheet 数量 > 10 时显示编号前缀，方便快速定位
 */
function syncPaperSelector() {
  if (!els.paperSelector) return;
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  const showIndex = currentPapers.length > 10;
  els.paperSelector.innerHTML = currentPapers.map((p, i) => {
    const label = showIndex
      ? `${i + 1}. ${escapeHtml(p.shortName || p.name)}`
      : escapeHtml(p.shortName || p.name);
    return `<option value="${i}"${i === currentPaperIndex ? " selected" : ""}>${label}</option>`;
  }).join("");
}

// -------------------- 报价表组切换 --------------------
function renderPriceListSelector() {
  if (!els.priceListSelector) return;
  els.priceListSelector.innerHTML = PRICE_LISTS.map(pl => {
    const selected = pl.id === CURRENT_PRICE_LIST_ID ? " selected" : "";
    return `<option value="${escapeHtml(pl.id)}"${selected}>${escapeHtml(pl.name)}</option>`;
  }).join("");
}

function onPriceListChange() {
  if (!els.priceListSelector) return;
  const newId = els.priceListSelector.value;
  if (newId && newId !== CURRENT_PRICE_LIST_ID) {
    setCurrentPriceList(newId);
    currentPaperIndex = 0; // 重置纸张索引
    // 纸张按报价表隔离，需重渲染；吊绳/邮费全局共享，无需重渲染
    rebuildPaperUI();
    renderPriceTable();
    onCalculate();
    showToast(`已切换到「${getCurrentPriceList().name}」`);
  }
}

function onDeletePriceList() {
  if (PRICE_LISTS.length <= 1) {
    showToast("至少保留一个报价表，无法删除");
    return;
  }
  const pl = getCurrentPriceList();
  if (!confirm(`确定删除报价表「${pl.name}」吗？\n\n该报价表下的 ${getPapersByPriceList(CURRENT_PRICE_LIST_ID).length} 张纸张及关联工艺将被删除。\n吊绳/邮费为全局共享，不受影响。\n\n此操作不可撤销。`)) return;
  deletePriceList(CURRENT_PRICE_LIST_ID);
  currentPaperIndex = 0;
  renderPriceListSelector();
  rebuildPaperUI();
  renderPriceTable();
  onCalculate();
  showToast(`已删除报价表「${pl.name}」`);
}

// -------------------- 翻页 --------------------
function prevPaper() {
  if (currentPaperIndex > 0) {
    currentPaperIndex--;
    renderPriceTable();
  }
}

function nextPaper() {
  if (currentPaperIndex < getPapersByPriceList(CURRENT_PRICE_LIST_ID).length - 1) {
    currentPaperIndex++;
    renderPriceTable();
  }
}

// -------------------- 提示 --------------------
function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    els.toast.classList.remove("show");
  }, 3000);
}

// ============================================================
// ===================== 本地存储与扩展功能 =====================
// ============================================================

// -------------------- 个人主页：客户等级渲染 --------------------
function renderLevelSettings() {
  if (!els.levelSettings) return;
  els.levelSettings.innerHTML = CUSTOMER_LEVELS.map((level, index) => `
    <div class="level-editor" data-level-id="${level.id}">
      <input type="text" class="level-name" value="${escapeHtml(level.name)}" placeholder="等级名称" />
      <input type="number" class="level-coefficient" value="${level.coefficient}" min="1" step="0.01" />
      <span class="unit">倍</span>
      <button class="btn danger sm" data-action="remove-level" data-index="${index}">删除</button>
    </div>
  `).join("");

  // 绑定删除按钮
  els.levelSettings.querySelectorAll("[data-action='remove-level']").forEach(btn => {
    btn.addEventListener("click", () => {
      const index = Number(btn.dataset.index);
      if (CUSTOMER_LEVELS.length <= 1) {
        showToast("至少保留一个客户等级");
        return;
      }
      const levelName = CUSTOMER_LEVELS[index]?.name || "该等级";
      if (!confirm(`确定要删除客户等级「${levelName}」吗？\n\n删除后不可撤销，如需恢复可点击「恢复默认」。`)) return;
      CUSTOMER_LEVELS.splice(index, 1);
      saveToStorage("customerLevels", CUSTOMER_LEVELS);
      renderLevelSettings();
      onCalculate();
      showToast("已删除等级");
    });
  });

  // 实时保存输入变化
  els.levelSettings.querySelectorAll(".level-name, .level-coefficient").forEach(input => {
    input.addEventListener("change", () => {
      const rows = els.levelSettings.querySelectorAll(".level-editor");
      CUSTOMER_LEVELS = Array.from(rows).map(row => ({
        id: row.dataset.levelId || generateId(),
        name: row.querySelector(".level-name").value.trim() || "未命名",
        coefficient: parseFloat(row.querySelector(".level-coefficient").value) || 1
      }));
      saveToStorage("customerLevels", CUSTOMER_LEVELS);
      onCalculate();
      showToast("系数已更新");
    });
  });
}

function addCustomerLevel() {
  CUSTOMER_LEVELS.push({
    id: generateId(),
    name: "新客户等级 " + (CUSTOMER_LEVELS.length + 1),
    coefficient: 1.2
  });
  saveToStorage("customerLevels", CUSTOMER_LEVELS);
  renderLevelSettings();
  onCalculate();
  showToast("已新增等级");
}

function resetCustomerLevels() {
  CUSTOMER_LEVELS = DEFAULT_CUSTOMER_LEVELS.map(l => ({ ...l }));
  saveToStorage("customerLevels", CUSTOMER_LEVELS);
  renderLevelSettings();
  onCalculate();
  showToast("已恢复默认等级");
}

// -------------------- 模式切换 --------------------
function switchCalcMode(mode) {
  calcMode = mode;
  saveToStorage("currentCalcMode", mode);
  els.modeBtns.forEach(btn => {
    btn.classList.toggle("active", btn.dataset.mode === mode);
  });
  // 直接系数模式隐藏吊绳和邮费输入区域
  const ropeGroup = els.rope ? els.rope.closest(".form-group") : null;
  const regionGroup = els.region ? els.region.closest(".form-group") : null;
  if (mode === "direct") {
    if (ropeGroup) ropeGroup.style.display = "none";
    if (regionGroup) regionGroup.style.display = "none";
  } else {
    if (ropeGroup) ropeGroup.style.display = "";
    if (regionGroup) regionGroup.style.display = "";
  }
  onCalculate();
}

/**
 * 一键恢复所有配置为出厂默认（纸张 / 工艺 / 吊绳 / 客户等级）。
 * 仅清 4 个 key 的 localStorage；保留报价历史、本地快照、APP_PROFILE、个人偏好等。
 * 用途：换设备或被旧 localStorage 污染后，一键回到 DEFAULT 状态。
 */
function resetToDefaults() {
  const confirmMsg = "将清空以下本地修改并恢复出厂默认：\n\n• 报价表组（恢复为单个1号报价表）\n• 纸张配置（10 张1号报价表）\n• 工艺配置（含 烫金/UV/鸡眼/凹凸 等）\n• 吊绳配置\n• 邮费配置\n• 客户等级\n\n报价历史与本地快照不会被删除。\n\n确定继续？";
  if (!confirm(confirmMsg)) return;

  const keysToReset = ["paperConfig", "craftConfig", "ropeConfig", "shippingConfig", "customerLevels", "priceLists", "currentPriceListId"];
  keysToReset.forEach(k => {
    try { localStorage.removeItem("tagPricing_" + k); } catch (e) { /* 忽略 */ }
  });

  // 重新从 DEFAULT 派生（深拷贝，避免后续修改污染源对象）
  PRICE_LISTS = DEFAULT_PRICE_LISTS.map(p => ({ ...p }));
  CURRENT_PRICE_LIST_ID = "priceList1";
  PAPER_CONFIG = DEFAULT_PAPER_CONFIG.map(p => ({
    ...p,
    specs: p.specs.map(s => ({ ...s, prices: { ...s.prices } }))
  }));
  CRAFT_CONFIG = {};
  Object.keys(DEFAULT_CRAFT_CONFIG).forEach(k => {
    CRAFT_CONFIG[k] = DEFAULT_CRAFT_CONFIG[k].map(c => ({ ...c, prices: { ...c.prices } }));
  });
  ROPE_CONFIG = DEFAULT_ROPE_CONFIG.map(r => ({ ...r, prices: { ...r.prices } }));
  SHIPPING_CONFIG = DEFAULT_SHIPPING_CONFIG.map(s => ({ ...s, basePrices: { ...s.basePrices } }));
  CUSTOMER_LEVELS = DEFAULT_CUSTOMER_LEVELS.map(l => ({ ...l }));
  currentPaperIndex = 0;

  rebuildPaperUI();
  renderLevelSettings();
  renderRopeRadios();
  initOptions();
  // 重新填充下拉框默认值
  updateDefaultPaperOptions && updateDefaultPaperOptions();
  updateDefaultRopeOptions && updateDefaultRopeOptions();
  renderPriceListSelector && renderPriceListSelector();
  onCalculate();
  showToast("已恢复默认配置（含报价表组）");
}

/**
 * 一键恢复全局默认设置：清除浏览器中所有本地配置并刷新页面。
 * 清除范围：纸张 / 工艺 / 吊绳 / 客户等级 / 公司信息 / 个人偏好（appProfile）。
 * 保留范围：报价历史、本地快照。
 */
function resetAllLocalSettings() {
  const confirmMsg = "⚠️ 确定要恢复全局默认设置吗？\n\n将清除以下本地配置：\n• 报价表组（恢复为单个1号报价表）\n• 纸张配置（10 张1号报价表）\n• 工艺配置（含 烫金/UV/鸡眼/凹凸 等）\n• 吊绳配置\n• 邮费配置\n• 客户等级与毛利系数\n• 公司信息与个人偏好\n\n报价历史与本地快照不受影响。\n\n此操作不可撤销，恢复后页面将自动刷新。";
  if (!confirm(confirmMsg)) return;

  // 清除所有本地配置 key（保留 history 和 snapshots）
  const keysToWipe = ["paperConfig", "craftConfig", "ropeConfig", "shippingConfig", "customerLevels", "appProfile", "priceLists", "currentPriceListId"];
  keysToWipe.forEach(k => {
    try { localStorage.removeItem("tagPricing_" + k); } catch (e) { /* 忽略 */ }
  });

  showToast("已恢复全局默认设置，正在刷新…");
  setTimeout(() => location.reload(), 600);
}

function loadProfileToUI() {
  if (els.companyName) els.companyName.value = APP_PROFILE.companyName || "";
  if (els.companyPhone) els.companyPhone.value = APP_PROFILE.companyPhone || "";
  if (els.decimalPlaces) els.decimalPlaces.value = parseDecimalPlaces(APP_PROFILE.decimalPlaces);
  updateDefaultTierOptions();
  if (els.defaultTier) els.defaultTier.value = APP_PROFILE.defaultTier || "";
  updateDefaultRopeOptions();
  if (els.defaultRope) els.defaultRope.value = APP_PROFILE.defaultRope || "rope1";
  updateDefaultPaperOptions();
  if (els.defaultPaper) els.defaultPaper.value = APP_PROFILE.defaultPaperId || "";
}

function updateDefaultTierOptions() {
  if (!els.defaultTier) return;
  const firstPaper = getPapersByPriceList(CURRENT_PRICE_LIST_ID)[0];
  const tiers = firstPaper && firstPaper.specs.length
    ? Object.keys(firstPaper.specs[0].prices).map(Number).sort((a, b) => a - b)
    : [];
  els.defaultTier.innerHTML = '<option value="">跟随纸张</option>' +
    tiers.map(t => `<option value="${t}">${t} 张</option>`).join("");
}

function updateDefaultRopeOptions() {
  if (!els.defaultRope) return;
  const opts = ['<option value="">跟随默认（普通吊绳）</option>'].concat(
    ROPE_CONFIG.map(r => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.name)}</option>`)
  );
  els.defaultRope.innerHTML = opts.join("");
}

function updateDefaultPaperOptions() {
  if (!els.defaultPaper) return;
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  const defaultLabel = currentPapers.length ? `跟随默认（${currentPapers[0].shortName || currentPapers[0].name}）` : "跟随默认";
  const opts = [`<option value="">${escapeHtml(defaultLabel)}</option>`].concat(
    currentPapers.map(p => `<option value="${escapeHtml(p.id)}">${escapeHtml(p.shortName || p.name)}</option>`)
  );
  els.defaultPaper.innerHTML = opts.join("");
}

function saveProfile() {
  APP_PROFILE = {
    companyName: els.companyName ? els.companyName.value.trim() : APP_PROFILE.companyName,
    companyPhone: els.companyPhone ? els.companyPhone.value.trim() : APP_PROFILE.companyPhone,
    defaultTier: els.defaultTier ? els.defaultTier.value : APP_PROFILE.defaultTier,
    defaultRope: els.defaultRope ? (els.defaultRope.value || "rope1") : APP_PROFILE.defaultRope,
    defaultPaperId: els.defaultPaper ? (els.defaultPaper.value || "") : APP_PROFILE.defaultPaperId,
    decimalPlaces: els.decimalPlaces ? parseDecimalPlaces(els.decimalPlaces.value) : parseDecimalPlaces(APP_PROFILE.decimalPlaces)
  };
  saveToStorage("appProfile", APP_PROFILE);
  saveToStorage("customerLevels", CUSTOMER_LEVELS);
  showToast("设置已保存");
  onCalculate();
}

function exportProfile() {
  const data = {
    customerLevels: CUSTOMER_LEVELS,
    appProfile: APP_PROFILE,
    exportAt: new Date().toISOString()
  };
  downloadJson(data, "KOKALabel配置_" + formatDateFile() + ".json");
  showToast("配置已导出");
}

function formatDateFile() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}_${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}`;
}

function downloadJson(data, filename) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// -------------------- 数据管理：标签页 --------------------
function switchTab(tabName) {
  els.tabs.forEach(t => t.classList.toggle("active", t.dataset.tab === tabName));
  els.tabPanels.forEach(p => p.classList.toggle("active", p.id === "panel-" + tabName));
  if (tabName === "snapshot") renderSnapshots();
  if (tabName === "history") renderHistory();
}

// -------------------- 快照管理 --------------------
function getSnapshots() {
  return loadFromStorage("snapshots", []);
}

function saveSnapshots(list) {
  saveToStorage("snapshots", list);
}

function createSnapshot() {
  const name = prompt("请输入快照名称", "快照 " + formatDateFile());
  if (!name) return;
  const list = getSnapshots();
  list.unshift({
    id: generateId(),
    name: name,
    createdAt: new Date().toISOString(),
    data: {
      priceLists: PRICE_LISTS,
      currentPriceListId: CURRENT_PRICE_LIST_ID,
      customerLevels: CUSTOMER_LEVELS,
      appProfile: APP_PROFILE,
      paperConfig: PAPER_CONFIG,
      ropeConfig: ROPE_CONFIG,
      craftConfig: CRAFT_CONFIG,
      shippingConfig: SHIPPING_CONFIG
    }
  });
  saveSnapshots(list);
  renderSnapshots();
  showToast("快照已创建");
}

function renderSnapshots() {
  const list = getSnapshots();
  const tbody = els.snapshotTable.querySelector("tbody");
  els.snapshotEmpty.style.display = list.length ? "none" : "block";
  els.snapshotTable.style.display = list.length ? "table" : "none";
  tbody.innerHTML = list.map(s => `
    <tr>
      <td>${escapeHtml(s.name)}</td>
      <td>${new Date(s.createdAt).toLocaleString()}</td>
      <td>
        <button class="btn sm" data-action="restore-snapshot" data-id="${s.id}">恢复</button>
        <button class="btn danger sm" data-action="delete-snapshot" data-id="${s.id}">删除</button>
      </td>
    </tr>
  `).join("");

  tbody.querySelectorAll("[data-action='restore-snapshot']").forEach(btn => {
    btn.addEventListener("click", () => {
      const item = getSnapshots().find(x => x.id === btn.dataset.id);
      if (!item) return;
      if (item.data.priceLists) {
        PRICE_LISTS = item.data.priceLists;
        saveToStorage("priceLists", PRICE_LISTS);
      }
      if (item.data.currentPriceListId) {
        CURRENT_PRICE_LIST_ID = item.data.currentPriceListId;
        saveToStorage("currentPriceListId", CURRENT_PRICE_LIST_ID);
      }
      if (item.data.customerLevels) {
        CUSTOMER_LEVELS = item.data.customerLevels;
        saveToStorage("customerLevels", CUSTOMER_LEVELS);
      }
      if (item.data.appProfile) {
        APP_PROFILE = item.data.appProfile;
        saveToStorage("appProfile", APP_PROFILE);
      }
      if (item.data.paperConfig) {
        PAPER_CONFIG = item.data.paperConfig;
        saveToStorage("paperConfig", PAPER_CONFIG);
      }
      if (item.data.craftConfig) {
        CRAFT_CONFIG = item.data.craftConfig;
        saveToStorage("craftConfig", CRAFT_CONFIG);
      }
      currentPaperIndex = 0;
      renderPriceListSelector && renderPriceListSelector();
      rebuildPaperUI();
      renderLevelSettings();
      loadProfileToUI();
      onCalculate();
      showToast("已恢复快照");
    });
  });

  tbody.querySelectorAll("[data-action='delete-snapshot']").forEach(btn => {
    btn.addEventListener("click", () => {
      const next = getSnapshots().filter(x => x.id !== btn.dataset.id);
      saveSnapshots(next);
      renderSnapshots();
      showToast("已删除快照");
    });
  });
}

function clearSnapshots() {
  if (!confirm("确定清空所有快照？")) return;
  saveSnapshots([]);
  renderSnapshots();
  showToast("快照已清空");
}

// -------------------- 报价历史 --------------------
function getHistory() {
  return loadFromStorage("history", []);
}

function saveHistory(list) {
  saveToStorage("history", list);
}

function addHistoryPlaceholder() {
  const list = getHistory();
  list.unshift({
    id: generateId(),
    createdAt: new Date().toISOString(),
    size: "55 × 30",
    paper: "1号报价表-3",
    tier: 1000,
    cost: 45.5,
    price: 59.15
  });
  saveHistory(list);
  renderHistory();
  showToast("已添加示例记录");
}

function renderHistory() {
  const list = getHistory();
  const tbody = els.historyTable.querySelector("tbody");
  els.historyEmpty.style.display = list.length ? "none" : "block";
  els.historyTable.style.display = list.length ? "table" : "none";
  tbody.innerHTML = list.map(h => `
    <tr>
      <td>${new Date(h.createdAt).toLocaleString()}</td>
      <td>${escapeHtml(h.size)}</td>
      <td>${escapeHtml(h.paper)}</td>
      <td>${h.tier} 张</td>
      <td>¥ ${formatMoney(h.cost)}</td>
      <td>¥ ${formatMoney(h.price)}</td>
      <td><button class="btn danger sm" data-action="delete-history" data-id="${h.id}">删除</button></td>
    </tr>
  `).join("");

  tbody.querySelectorAll("[data-action='delete-history']").forEach(btn => {
    btn.addEventListener("click", () => {
      const next = getHistory().filter(x => x.id !== btn.dataset.id);
      saveHistory(next);
      renderHistory();
      showToast("已删除记录");
    });
  });
}

function clearHistory() {
  if (!confirm("确定清空所有报价历史？")) return;
  saveHistory([]);
  renderHistory();
  showToast("历史已清空");
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// -------------------- 本地备份 / 恢复 --------------------
function setLocalBackupStatus(html, isError) {
  if (!els.localBackupStatus) return;
  els.localBackupStatus.innerHTML = html;
  els.localBackupStatus.style.color = isError ? "var(--danger)" : "var(--text-secondary)";
}

function exportLocalBackup() {
  const data = {
    version: "6.0",
    kind: "local-backup",
    exportAt: new Date().toISOString(),
    priceLists: PRICE_LISTS,
    currentPriceListId: CURRENT_PRICE_LIST_ID,
    customerLevels: CUSTOMER_LEVELS,
    appProfile: APP_PROFILE,
    paperConfig: PAPER_CONFIG,
    ropeConfig: ROPE_CONFIG,
    craftConfig: CRAFT_CONFIG,
    shippingConfig: SHIPPING_CONFIG,
    snapshots: getSnapshots(),
    history: getHistory()
  };
  downloadJson(data, "KOKALabel本地备份_" + formatDateFile() + ".json");
  setLocalBackupStatus(`已保存到本地文件（${new Date().toLocaleString()}）。建议同时存一份到 U 盘 / 网盘。`, false);
  showToast("本地备份已下载");
}


// -------------------- P0 导入数据校验 --------------------
/**
 * 校验导入的配置数据结构，防止恶意/损坏数据导致 XSS 或崩溃。
 */
function validateImportedData(data, expectedKind) {
  if (!data || typeof data !== "object") {
    throw new Error("数据格式无效");
  }
  if (expectedKind && data.kind && data.kind !== expectedKind && data.kind !== "full-config" && data.kind !== "local-backup") {
    throw new Error("文件类型不匹配");
  }
  // 校验字符串字段长度限制
  const MAX_STR_LEN = 200;
  const validateString = (val, fieldName) => {
    if (val != null && typeof val === "string" && val.length > MAX_STR_LEN) {
      throw new Error(fieldName + " 超过最大长度限制");
    }
  };
  if (data.customerLevels && Array.isArray(data.customerLevels)) {
    data.customerLevels.forEach(l => {
      validateString(l.name, "客户等级名称");
      if (typeof l.coefficient !== "number" || l.coefficient < 1 || l.coefficient > 100) {
        throw new Error("客户等级系数无效");
      }
    });
  }
  if (data.paperConfig && Array.isArray(data.paperConfig)) {
    data.paperConfig.forEach(p => {
      validateString(p.name, "纸张名称");
      validateString(p.shortName, "纸张简称");
      if (typeof p.discount !== "number" || p.discount <= 0 || p.discount > 10) {
        throw new Error("折扣系数无效: " + (p.shortName || p.name));
      }
    });
  }
  if (data.appProfile && typeof data.appProfile === "object") {
    validateString(data.appProfile.companyName, "公司名");
    validateString(data.appProfile.companyPhone, "电话");
  }
  return true;
}

function importLocalBackup(file) {
  if (!file) return;
  if (!confirm("恢复本地备份会覆盖当前所有配置（纸张/吊绳/工艺/客户等级/历史/快照）。是否继续？")) {
    if (els.importLocalBackupFile) els.importLocalBackupFile.value = "";
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      validateImportedData(data);
      if (data.kind && data.kind !== "local-backup" && data.kind !== "full-config") {
        throw new Error("文件类型不匹配，请使用「保存到本地文件」生成的文件");
      }
      // 还原报价表组结构
      if (data.priceLists && Array.isArray(data.priceLists)) {
        PRICE_LISTS = data.priceLists;
        saveToStorage("priceLists", PRICE_LISTS);
      }
      if (data.currentPriceListId) {
        CURRENT_PRICE_LIST_ID = data.currentPriceListId;
        saveToStorage("currentPriceListId", CURRENT_PRICE_LIST_ID);
      }
      if (data.customerLevels) {
        CUSTOMER_LEVELS = data.customerLevels;
        saveToStorage("customerLevels", CUSTOMER_LEVELS);
      }
      if (data.appProfile) {
        APP_PROFILE = data.appProfile;
        APP_PROFILE.decimalPlaces = parseDecimalPlaces(APP_PROFILE.decimalPlaces);
        saveToStorage("appProfile", APP_PROFILE);
      }
      if (data.paperConfig) {
        PAPER_CONFIG = data.paperConfig;
        saveToStorage("paperConfig", PAPER_CONFIG);
      }
      if (data.ropeConfig) {
        ROPE_CONFIG = data.ropeConfig;
        saveToStorage("ropeConfig", ROPE_CONFIG);
        rebuildRopeUI();
      }
      if (data.craftConfig) {
        CRAFT_CONFIG = data.craftConfig;
        saveToStorage("craftConfig", CRAFT_CONFIG);
      }
      if (data.shippingConfig) {
        SHIPPING_CONFIG.length = 0;
        data.shippingConfig.forEach(s => SHIPPING_CONFIG.push(s));
        saveToStorage("shippingConfig", SHIPPING_CONFIG);
        if (els.region) {
          els.region.innerHTML = '<option value="">请选择地区</option>' +
            SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
          if (SHIPPING_CONFIG.length > 0) els.region.value = SHIPPING_CONFIG[0].id;
        }
      }
      if (data.snapshots) saveSnapshots(data.snapshots);
      if (data.history) saveHistory(data.history);
      currentPaperIndex = 0;
      renderPriceListSelector && renderPriceListSelector();
      rebuildPaperUI();
      renderLevelSettings();
      loadProfileToUI();
      renderSnapshots();
      renderHistory();
      onCalculate();
      setLocalBackupStatus(`已从本地文件恢复：${escapeHtml(file.name)}（${new Date().toLocaleString()}）`, false);
      showToast("本地备份已恢复");
    } catch (err) {
      setLocalBackupStatus("恢复失败：" + err.message, true);
      showToast("本地备份恢复失败");
    }
    if (els.importLocalBackupFile) els.importLocalBackupFile.value = "";
  };
  reader.onerror = () => {
    setLocalBackupStatus("文件读取失败", true);
    if (els.importLocalBackupFile) els.importLocalBackupFile.value = "";
  };
  reader.readAsText(file);
}

// -------------------- 导入 / 导出完整配置 --------------------
function exportFullData() {
  const data = {
    version: "6.0",
    exportAt: new Date().toISOString(),
    priceLists: PRICE_LISTS,
    currentPriceListId: CURRENT_PRICE_LIST_ID,
    customerLevels: CUSTOMER_LEVELS,
    appProfile: APP_PROFILE,
    paperConfig: PAPER_CONFIG,
    ropeConfig: ROPE_CONFIG,
    craftConfig: CRAFT_CONFIG,
    shippingConfig: SHIPPING_CONFIG,
    snapshots: getSnapshots(),
    history: getHistory()
  };
  downloadJson(data, "KOKALabel完整配置_" + formatDateFile() + ".json");
  showToast("完整配置已导出");
}

function importFullData(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      // 还原报价表组结构
      if (data.priceLists && Array.isArray(data.priceLists)) {
        PRICE_LISTS = data.priceLists;
        saveToStorage("priceLists", PRICE_LISTS);
      }
      if (data.currentPriceListId) {
        CURRENT_PRICE_LIST_ID = data.currentPriceListId;
        saveToStorage("currentPriceListId", CURRENT_PRICE_LIST_ID);
      }
      if (data.customerLevels) {
        CUSTOMER_LEVELS = data.customerLevels;
        saveToStorage("customerLevels", CUSTOMER_LEVELS);
      }
      if (data.appProfile) {
        APP_PROFILE = data.appProfile;
        saveToStorage("appProfile", APP_PROFILE);
      }
      if (data.paperConfig) {
        PAPER_CONFIG = data.paperConfig;
        saveToStorage("paperConfig", PAPER_CONFIG);
      }
      if (data.ropeConfig) {
        ROPE_CONFIG = data.ropeConfig;
        saveToStorage("ropeConfig", ROPE_CONFIG);
        rebuildRopeUI();
      }
      if (data.craftConfig) {
        CRAFT_CONFIG = data.craftConfig;
        saveToStorage("craftConfig", CRAFT_CONFIG);
      }
      if (data.shippingConfig) {
        SHIPPING_CONFIG.length = 0;
        data.shippingConfig.forEach(s => SHIPPING_CONFIG.push(s));
        saveToStorage("shippingConfig", SHIPPING_CONFIG);
        if (els.region) {
          els.region.innerHTML = '<option value="">请选择地区</option>' +
            SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
          if (SHIPPING_CONFIG.length > 0) els.region.value = SHIPPING_CONFIG[0].id;
        }
      }
      if (data.snapshots) saveSnapshots(data.snapshots);
      if (data.history) saveHistory(data.history);
      currentPaperIndex = 0;
      renderPriceListSelector && renderPriceListSelector();
      rebuildPaperUI();
      renderLevelSettings();
      loadProfileToUI();
      renderSnapshots();
      renderHistory();
      onCalculate();
      showToast("配置导入成功");
    } catch (err) {
      showToast("JSON 解析失败，请检查文件格式");
    }
  };
  reader.readAsText(file);
}

// -------------------- Excel 导入 / 导出 / 模板 --------------------

function setExcelStatus(html, isError) {
  if (!els.excelImportStatus) return;
  els.excelImportStatus.innerHTML = html;
  els.excelImportStatus.style.color = isError ? "var(--danger)" : "var(--text-secondary)";
}

function paperToSheetRows(paper) {
  // 返回二维数组，按规划模板结构（含工艺区）
  // 关键：取规格档位 + 工艺档位的并集并排序，确保规格行与工艺行列宽一致
  const specTierKeys = paper.specs.length
    ? Object.keys(paper.specs[0].prices).map(Number)
    : [];
  const crafts = CRAFT_CONFIG[paper.id] || [];
  const craftTierKeys = crafts.length
    ? Object.keys(crafts[0].prices).map(Number)
    : [];
  const tierSet = new Set([...specTierKeys, ...craftTierKeys]);
  const tierKeys = Array.from(tierSet).sort((a, b) => a - b).map(String);

  const headerRow = ["代码", "最大含出血面积", ...tierKeys];
  const dataRows = paper.specs.map(spec => [
    spec.code,
    spec.maxArea,
    ...tierKeys.map(t => {
      // 缺值时留空（与源数据一致，不臆造 0）
      const v = spec.prices[t];
      return v == null ? "" : v;
    })
  ]);
  const craftRows = crafts.length
    ? [
        [],
        ["工艺名称", ...tierKeys],
        ...crafts.map(craft => [
          craft.name,
          ...tierKeys.map(t => {
            const v = craft.prices[t];
            return v == null ? "" : v;
          })
        ])
      ]
    : [];
  // 直接系数三行（与报价表表格格式一致：直接系数档位 / 最高倍数 / 最低倍数）
  // 有直接系数 → 填写实际档位/最高/最低；无直接系数 → 用纸张价格档位作为占位档位，最高/最低留空
  function directCoeffRows(paper, fallbackTiers) {
    const dc = paper.directCoeff;
    const hasDC = dc && Array.isArray(dc.tiers) && dc.tiers.length > 0 &&
      Array.isArray(dc.max) && dc.max.length > 0 && Array.isArray(dc.min) && dc.min.length > 0;
    if (hasDC) {
      return [
        ["直接系数档位", ...dc.tiers],
        ["最高倍数", ...dc.max],
        ["最低倍数", ...dc.min]
      ];
    }
    const tiers = fallbackTiers && fallbackTiers.length ? fallbackTiers : [];
    return [
      ["直接系数档位", ...tiers],
      ["最高倍数", ...tiers.map(() => "")],
      ["最低倍数", ...tiers.map(() => "")]
    ];
  }
  return [
    ["所属小组", GROUP_META.name],
    ["总报价表", getCurrentPriceList().name],
    ["报价表全称", paper.name],
    ["简称", paper.shortName],
    ["折扣系数", paper.discount],
    ...directCoeffRows(paper, tierKeys),
    ["备注", ""],
    [],
    headerRow,
    ...dataRows,
    ...craftRows
  ];
}

function exportPaperExcel() {
  // P1.4: 确保 SheetJS 已加载
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => exportPaperExcel()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  for (const paper of currentPapers) {
    const ws = XLSX.utils.aoa_to_sheet(paperToSheetRows(paper));
    XLSX.utils.book_append_sheet(wb, ws, paper.shortName || paper.name);
  }
  const plName = getCurrentPriceList().name;
  XLSX.writeFile(wb, `KOKALabel${plName}_${formatDateFile()}.xlsx`);
  showToast(`${plName} Excel 已导出（${currentPapers.length} 张）`);
}

function downloadPaperTemplate() {
  // P1.4: 确保 SheetJS 已加载
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => downloadPaperTemplate()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();

  // === 模板 = 默认报价表（DEFAULT_PAPER_CONFIG + DEFAULT_CRAFT_CONFIG）动态生成 ===
  // 每个 Sheet 均含「直接系数档位 / 最高倍数 / 最低倍数」三行：
  //   - 有直接系数的 Sheet（如 350/400/702铜版纸 等）已填实际档位/最高/最低
  //   - 无直接系数的 Sheet（700布纹纸 / 40C棉麻布）提供占位档位行，最高/最低留空，方便填写后重新导入
  DEFAULT_PAPER_CONFIG.forEach((paper, idx) => {
    const rows = [];
    // 元信息区
    rows.push(["所属小组", GROUP_META.name]);
    rows.push(["总报价表", getCurrentPriceList().name]);
    rows.push(["报价表全称", paper.name]);
    rows.push(["简称", paper.shortName]);
    rows.push(["折扣系数", paper.discount]);
    // 直接系数三行（有系数填实际值，无系数用价格档位占位）
    const dc = paper.directCoeff;
    const hasDC = dc && Array.isArray(dc.tiers) && dc.tiers.length > 0 &&
      Array.isArray(dc.max) && dc.max.length > 0 && Array.isArray(dc.min) && dc.min.length > 0;
    const specTierKeys = paper.specs.length ? Object.keys(paper.specs[0].prices).map(Number) : [];
    const dcTiers = hasDC ? dc.tiers : specTierKeys;
    rows.push(["直接系数档位", ...dcTiers]);
    rows.push(["最高倍数", ...(hasDC ? dc.max : dcTiers.map(() => ""))]);
    rows.push(["最低倍数", ...(hasDC ? dc.min : dcTiers.map(() => ""))]);
    rows.push(["备注", "直接系数档位/最高倍数/最低倍数三行：档位为批量张数，最高倍数→普通客户，最低倍数→大客户，中间等级自动等差插值。无直接系数的纸张按标准报价（乘折扣系数）计算。"]);
    rows.push([]);

    // 规格区：取规格档位 + 工艺档位并集
    const crafts = DEFAULT_CRAFT_CONFIG[paper.id] || [];
    const craftTierKeys = crafts.length ? Object.keys(crafts[0].prices).map(Number) : [];
    const tierSet = new Set([...specTierKeys, ...craftTierKeys]);
    const tierKeys = Array.from(tierSet).sort((a, b) => a - b).map(String);

    rows.push(["代码", "最大含出血面积", ...tierKeys]);
    paper.specs.forEach(spec => {
      rows.push([
        spec.code,
        spec.maxArea,
        ...tierKeys.map(t => {
          const v = spec.prices[t];
          return v == null ? "" : v;
        })
      ]);
    });

    // 工艺区
    if (crafts.length) {
      rows.push([]);
      rows.push(["工艺名称", ...tierKeys]);
      crafts.forEach(craft => {
        rows.push([
          craft.name,
          ...tierKeys.map(t => {
            const v = craft.prices[t];
            return v == null ? "" : v;
          })
        ]);
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, paper.shortName || ("Sheet" + (idx + 1)));
  });

  XLSX.writeFile(wb, "KOKALabel1号报价表模板_" + formatDateFile() + ".xlsx");
  showToast("模板已下载（" + DEFAULT_PAPER_CONFIG.length + "张表，含直接系数档位规则行）");
}

function parsePaperExcel(arrayBuffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel 库未加载，请稍后重试");
    throw new Error("Excel 库未加载，请检查网络");
  }
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  if (!wb.SheetNames.length) throw new Error("Excel 没有工作表");

  const papers = [];
  const craftsByPaper = {};
  const shortNames = new Set();
  const errors = [];
  // 收集所有已使用的 paper id（默认配置 + 运行时 PAPER_CONFIG + 本次导入已分配），防止 ID 碰撞
  const usedIds = new Set([...DEFAULT_PAPER_CONFIG.map(p => p.id), ...PAPER_CONFIG.map(p => p.id)]);
  // 解析报价表名称（取第一个 Sheet 的"总报价表"行）
  let priceListName = "";

  wb.SheetNames.forEach((sheetName, sheetIndex) => {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows.length) {
      errors.push(`「${sheetName}」为空，已跳过`);
      return;
    }

    // 读取表头信息区（按标签匹配，兼容有无"所属小组"/"总报价表"行）
    let name = "", shortName = "", discountRaw = "";
    for (let i = 0; i < Math.min(rows.length, 8); i++) {
      const label = String(rows[i] && rows[i][0] || "").trim();
      const val = rows[i] && rows[i][1];
      if (label === "所属小组") continue;
      if (label === "总报价表") {
        // 捕获报价表名称（取第一个非空值）
        if (!priceListName && val) priceListName = String(val).trim();
        continue;
      }
      if (label === "1号报价表全称" || label === "报价表全称") name = String(val || "").trim();
      else if (label === "简称") shortName = String(val || "").trim();
      else if (label === "折扣系数") discountRaw = val;
    }
    const discount = discountRaw === "" || discountRaw == null ? 1 : Number(discountRaw);

    // 解析直接系数三行（直接系数档位 / 最高倍数 / 最低倍数），与报价表表格格式一致
    // 三行均填有值 → { tiers, max, min }；仅档位占位或缺失 → null（按标准报价计算）
    let directCoeff = null;
    {
      const dcRows = { tiers: null, max: null, min: null };
      for (let i = 0; i < rows.length; i++) {
        const label = String(rows[i] && rows[i][0] || "").trim();
        if (label === "直接系数档位") dcRows.tiers = rows[i];
        else if (label === "最高倍数") dcRows.max = rows[i];
        else if (label === "最低倍数") dcRows.min = rows[i];
      }
      const toNumArr = (row) => {
        const arr = [];
        if (!row) return arr;
        for (let c = 1; c < row.length; c++) {
          const v = row[c];
          if (v === "" || v == null) continue;
          const n = Number(v);
          if (!isNaN(n)) arr.push(n);
        }
        return arr;
      };
      const tiers = toNumArr(dcRows.tiers).filter(n => n > 0 && Number.isInteger(n));
      const maxs = toNumArr(dcRows.max);
      const mins = toNumArr(dcRows.min);
      // 三行都有值且档位数量一致 → 有效直接系数
      if (tiers.length && maxs.length === tiers.length && mins.length === tiers.length) {
        directCoeff = { tiers, max: maxs, min: mins };
      }
      // 仅档位占位（无最高/最低）或无任何行 → directCoeff 保持 null
    }

    if (!name) {
      errors.push(`「${sheetName}」缺少报价表全称，已跳过`);
      return;
    }
    if (!shortName) {
      errors.push(`「${sheetName}」缺少简称，已跳过`);
      return;
    }
    if (shortNames.has(shortName)) {
      errors.push(`简称「${shortName}」重复，已跳过`);
      return;
    }
    if (isNaN(discount) || discount <= 0) {
      errors.push(`「${sheetName}」折扣系数无效，已使用默认值 1`);
    }

    shortNames.add(shortName);

    // 查找规格标题行（包含"代码"和"最大含出血面积"）
    let headerRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row && String(row[0] || "").trim() === "代码" && String(row[1] || "").trim() === "最大含出血面积") {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex === -1) {
      errors.push(`「${sheetName}」未找到规格标题行，已跳过`);
      return;
    }

    const headerRow = rows[headerRowIndex];
    const tierKeys = [];
    for (let c = 2; c < headerRow.length; c++) {
      const val = headerRow[c];
      if (val !== "" && !isNaN(Number(val)) && Number(val) > 0) {
        tierKeys.push(String(Number(val)));
      }
    }

    const specs = [];
    const seenCodes = new Set();
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.length) continue;
      // 遇到「工艺名称」标题即停止代码区（与下方工艺读取自然衔接）
      if (String(row[0] || "").trim() === "工艺名称") break;
      const code = String(row[0] || "").trim();
      if (!code) continue;
      if (seenCodes.has(code)) {
        errors.push(`「${sheetName}」代码 ${code} 重复，已跳过第 ${r + 1} 行`);
        continue;
      }
      const maxArea = Number(row[1]);
      if (isNaN(maxArea) || maxArea <= 0 || !Number.isInteger(maxArea)) {
        errors.push(`「${sheetName}」第 ${r + 1} 行最大含出血面积无效，已跳过`);
        continue;
      }
      seenCodes.add(code);

      const prices = {};
      for (let i = 0; i < tierKeys.length; i++) {
        const tier = tierKeys[i];
        const val = row[2 + i];
        if (val === "" || val == null) {
          // 数据源没有该批量价格 -> 保留 null（占位，UI 显示"无该批量定价"）
          prices[tier] = null;
        } else {
          const price = Number(val);
          if (isNaN(price)) {
            errors.push(`「${sheetName}」${code} 的 ${tier} 档价格无效，已按占位处理`);
            prices[tier] = null;
          } else {
            prices[tier] = price;
          }
        }
      }
      specs.push({ code, maxArea, prices });
    }

    if (!specs.length) {
      errors.push(`「${sheetName}」无规格数据，已跳过`);
      return;
    }

    // 尽量沿用默认配置中相同简称的 paper id，使工艺配置 CRAFT_CONFIG 保持对应
    const defaultPaper = DEFAULT_PAPER_CONFIG.find(p => p.shortName === shortName);
    let id;
    if (defaultPaper) {
      id = defaultPaper.id;
    } else {
      // 不匹配默认简称：生成唯一 ID，避免与默认 paper1~9 或其他导入纸碰撞
      let counter = usedIds.size + 1;
      do {
        id = "paper_import_" + counter;
        counter++;
      } while (usedIds.has(id));
    }
    usedIds.add(id);

    // 读取工艺区：查找 "工艺名称" 标题行
    const crafts = [];
    const seenCraftKeys = new Set();
    let craftHeaderIndex = -1;
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (row && String(row[0] || "").trim() === "工艺名称") {
        craftHeaderIndex = i;
        break;
      }
    }
    if (craftHeaderIndex !== -1) {
      for (let r = craftHeaderIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row.length) continue;
        const craftName = String(row[0] || "").trim();
        if (!craftName) continue;
        // 允许同名工艺：用 (名称 + 行号) 作为去重键，避免误丢
        const craftKey = craftName + "@row" + r;
        if (seenCraftKeys.has(craftKey)) {
          errors.push(`「${sheetName}」工艺「${craftName}」第 ${r + 1} 行重复，已跳过`);
          continue;
        }
        seenCraftKeys.add(craftKey);
        const prices = {};
        for (let i = 0; i < tierKeys.length; i++) {
          const tier = tierKeys[i];
          const val = row[1 + i];
          if (val === "" || val == null) {
            // 数据源没有该工艺批量价格 -> 占位
            prices[tier] = null;
          } else {
            const price = Number(val);
            if (isNaN(price)) {
              errors.push(`「${sheetName}」工艺「${craftName}」的 ${tier} 档价格无效，已按占位处理`);
              prices[tier] = null;
            } else {
              prices[tier] = price;
            }
          }
        }
        crafts.push({ id: `craft_${id}_${crafts.length + 1}`, name: craftName, prices });
      }
    }

    papers.push({
      id,
      name,
      shortName,
      discount: isNaN(discount) || discount <= 0 ? 1 : discount,
      // 直接系数：Sheet 专属配置，只读取报价表表格三行（直接系数档位/最高倍数/最低倍数）。
      // 表格未填有效直接系数时，匹配默认简称则继承默认配置，否则为 null（按标准报价计算）
      directCoeff: directCoeff || (defaultPaper && defaultPaper.directCoeff
        ? JSON.parse(JSON.stringify(defaultPaper.directCoeff))
        : null),
      specs
    });
    if (crafts.length) {
      craftsByPaper[id] = crafts;
    }
  });

  if (!papers.length) {
    throw new Error(errors.join("；") || "没有可导入的报价表");
  }

  return { papers, crafts: craftsByPaper, errors, priceListName: priceListName || "" };
}

function importPaperExcel(file) {
  if (!file) return;
  if (typeof XLSX === "undefined") {
    loadSheetJS().then(() => importPaperExcel(file)).catch(() => { setExcelStatus("Excel 库加载失败，请检查网络", true); showToast("Excel 库加载失败"); });
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const { papers, crafts, errors, priceListName } = parsePaperExcel(e.target.result);
      // 创建新报价表
      const plName = priceListName || `报价表${PRICE_LISTS.length + 1}`;
      const newPriceListId = addPriceList(plName);
      // 为导入的纸张设置 priceListId
      papers.forEach(p => { p.priceListId = newPriceListId; });
      // 追加到 PAPER_CONFIG（不替换已有报价表的纸张）
      PAPER_CONFIG = PAPER_CONFIG.concat(papers);
      saveToStorage("paperConfig", PAPER_CONFIG);
      // 合并工艺配置
      CRAFT_CONFIG = { ...CRAFT_CONFIG, ...crafts };
      saveToStorage("craftConfig", CRAFT_CONFIG);
      // 切换到新报价表
      currentPaperIndex = 0;
      rebuildPaperUI();
      renderPriceListSelector();
      renderPriceTable();
      onCalculate();
      const craftCount = Object.values(crafts).reduce((sum, arr) => sum + arr.length, 0);
      const successMsg = `成功导入报价表「${plName}」（${papers.length} 张纸张${craftCount ? `，含 ${craftCount} 条工艺` : ""}）：${papers.map(p => p.shortName).join("、")}`;
      const errMsg = errors.length ? `<br><span style="color:var(--danger)">警告：${errors.join("；")}</span>` : "";
      setExcelStatus(successMsg + errMsg, false);
      showToast(`已导入报价表「${plName}」`);
    } catch (err) {
      setExcelStatus("导入失败：" + err.message, true);
      showToast("Excel 导入失败");
    }
    if (els.importExcelFile) els.importExcelFile.value = "";
  };
  reader.onerror = () => {
    setExcelStatus("文件读取失败", true);
    if (els.importExcelFile) els.importExcelFile.value = "";
  };
  reader.readAsArrayBuffer(file);
}

function rebuildPaperUI() {
  // 重建纸张选择 UI、纸张设置列表、报价表查询页
  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  currentPaperIndex = Math.min(currentPaperIndex, Math.max(0, currentPapers.length - 1));
  // 重置纸张设置，保留尺寸数据，默认使用第一张纸
  const fallbackPaperId = currentPapers[0]?.id || "";
  sheetsState = sheetsState.map(s => ({
    paperId: currentPapers.find(p => p.id === s.paperId) ? s.paperId : fallbackPaperId,
    craftIds: [],
    width: s.width || "",
    length: s.length || "",
    sizeType: s.sizeType || "custom",
    manualCode: null
  }));
  updateTierOptions(false);
  renderSheets();
  renderPriceTable();
}

// -------------------- 吊绳 Excel 导入 / 导出 / 模板 --------------------

function setRopeExcelStatus(html, isError) {
  if (!els.ropeExcelImportStatus) return;
  els.ropeExcelImportStatus.innerHTML = html;
  els.ropeExcelImportStatus.style.color = isError ? "var(--danger)" : "var(--text-secondary)";
}

function ropeToSheetRows() {
  const tierKeys = ROPE_CONFIG.length ? Object.keys(ROPE_CONFIG[0].prices) : [];
  const headerRow = ["吊绳名称", ...tierKeys.map(String)];
  const dataRows = ROPE_CONFIG.map(rope => [
    rope.name,
    ...tierKeys.map(t => rope.prices[t] ?? "")
  ]);
  return [headerRow, ...dataRows];
}

function exportRopeExcel() {
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => exportRopeExcel()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(ropeToSheetRows());
  XLSX.utils.book_append_sheet(wb, ws, "吊绳报价");
  XLSX.writeFile(wb, "KOKALabel吊绳报价_" + formatDateFile() + ".xlsx");
  showToast("吊绳报价 Excel 已导出");
}

function downloadRopeTemplate() {
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => downloadRopeTemplate()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();
  const sample = DEFAULT_ROPE_CONFIG[0];
  const tierKeys = sample ? Object.keys(sample.prices) : ["500", "1000", "2000", "3000", "5000"];
  const rows = [
    ["吊绳名称", ...tierKeys.map(String)],
    ["不加吊绳", ...tierKeys.map(() => 0)],
    ["普通吊绳", 5, 10, 20, 30, 50]
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "吊绳报价模板");
  XLSX.writeFile(wb, "KOKALabel吊绳报价模板_" + formatDateFile() + ".xlsx");
  showToast("吊绳模板已下载");
}

function parseRopeExcel(arrayBuffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel 库未加载，请检查网络");
  }
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("Excel 没有工作表");

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) throw new Error("工作表为空");

  const headerRow = rows[0];
  const tierKeys = [];
  for (let c = 1; c < headerRow.length; c++) {
    const val = headerRow[c];
    if (val !== "" && !isNaN(Number(val)) && Number(val) > 0) {
      tierKeys.push(String(Number(val)));
    }
  }
  if (!tierKeys.length) throw new Error("未找到档位列，表头第 2 列起应为数字档位");

  const ropes = [];
  const seenNames = new Set();
  const errors = [];
  let hasZeroOption = false;

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const name = String(row[0] || "").trim();
    if (!name) continue;
    if (seenNames.has(name)) {
      errors.push(`吊绳「${name}」重复，已跳过第 ${r + 1} 行`);
      continue;
    }
    seenNames.add(name);

    const prices = {};
    let allZero = true;
    for (let i = 0; i < tierKeys.length; i++) {
      const tier = tierKeys[i];
      const val = row[1 + i];
      const price = val === "" || val == null ? 0 : Number(val);
      if (isNaN(price)) {
        errors.push(`「${name}」的 ${tier} 档价格无效，已按 0 处理`);
        prices[tier] = 0;
      } else {
        prices[tier] = price;
      }
      if (prices[tier] !== 0) allZero = false;
    }
    if (allZero) hasZeroOption = true;

    ropes.push({ id: "rope" + (ropes.length), name, prices });
  }

  if (!ropes.length) throw new Error("没有可导入的吊绳");

  // 确保第一个是不加吊绳（全 0 价格）
  if (!hasZeroOption) {
    errors.push("未找到全 0 价格的「不加吊绳」选项，已自动在首位添加");
    ropes.unshift({ id: "rope0", name: "不加吊绳", prices: Object.fromEntries(tierKeys.map(t => [t, 0])) });
  } else {
    // 把全 0 的移到第一位
    const zeroIndex = ropes.findIndex(r => Object.values(r.prices).every(p => p === 0));
    if (zeroIndex > 0) {
      const [zeroRope] = ropes.splice(zeroIndex, 1);
      ropes.unshift(zeroRope);
    }
  }
  // 重新分配 id
  ropes.forEach((rope, idx) => { rope.id = "rope" + idx; });

  return { ropes, errors };
}

function importRopeExcel(file) {
  if (!file) return;
  if (typeof XLSX === "undefined") {
    loadSheetJS().then(() => importRopeExcel(file)).catch(() => { setRopeExcelStatus("Excel 库加载失败，请检查网络", true); showToast("Excel 库加载失败"); });
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const { ropes, errors } = parseRopeExcel(e.target.result);
      ROPE_CONFIG = ropes;
      saveToStorage("ropeConfig", ROPE_CONFIG);
      rebuildRopeUI();
      onCalculate();
      const successMsg = `成功导入 ${ropes.length} 种吊绳：${ropes.map(r => r.name).join("、")}`;
      const errMsg = errors.length ? `<br><span style="color:var(--danger)">警告：${errors.join("；")}</span>` : "";
      setRopeExcelStatus(successMsg + errMsg, false);
      showToast("吊绳 Excel 导入成功");
    } catch (err) {
      setRopeExcelStatus("导入失败：" + err.message, true);
      showToast("吊绳 Excel 导入失败");
    }
    if (els.importRopeExcelFile) els.importRopeExcelFile.value = "";
  };
  reader.onerror = () => {
    setRopeExcelStatus("文件读取失败", true);
    if (els.importRopeExcelFile) els.importRopeExcelFile.value = "";
  };
  reader.readAsArrayBuffer(file);
}

function rebuildRopeUI() {
  // 重新渲染吊绳选择区（沿用与初始化一致的默认选择逻辑）
  els.rope.innerHTML = renderRopeRadios();
}

// -------------------- 邮费 Excel 导入 / 导出 / 模板 --------------------

function setShippingExcelStatus(html, isError) {
  if (!els.shippingExcelImportStatus) return;
  els.shippingExcelImportStatus.innerHTML = html;
  els.shippingExcelImportStatus.style.color = isError ? "var(--danger)" : "var(--text-secondary)";
}

function shippingToSheetRows() {
  const tierKeys = SHIPPING_CONFIG.length ? Object.keys(SHIPPING_CONFIG[0].basePrices) : [];
  const headerRow = ["地区名称", ...tierKeys.map(String), "小面积折扣阈值", "折扣系数"];
  const dataRows = SHIPPING_CONFIG.map(region => [
    region.name,
    ...tierKeys.map(t => region.basePrices[t] ?? ""),
    region.smallAreaThreshold,
    region.discount
  ]);
  return [headerRow, ...dataRows];
}

function exportShippingExcel() {
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => exportShippingExcel()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(shippingToSheetRows());
  XLSX.utils.book_append_sheet(wb, ws, "邮费报价");
  XLSX.writeFile(wb, "KOKALabel邮费报价_" + formatDateFile() + ".xlsx");
  showToast("邮费报价 Excel 已导出");
}

function downloadShippingTemplate() {
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => downloadShippingTemplate()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const sample = DEFAULT_SHIPPING_CONFIG[0];
  const tierKeys = sample ? Object.keys(sample.basePrices) : ["500", "1000", "2000", "3000", "5000"];
  const rows = [
    ["地区名称", ...tierKeys.map(String), "小面积折扣阈值", "折扣系数"],
    ["广东省内", ...tierKeys.map(t => sample ? sample.basePrices[t] : 0), 2500, 0.8],
    ["江浙沪", ...tierKeys.map(() => 0), 2500, 0.8],
    ["其他省份", ...tierKeys.map(() => 0), 2500, 0.8]
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "邮费报价模板");
  XLSX.writeFile(wb, "KOKALabel邮费报价模板_" + formatDateFile() + ".xlsx");
  showToast("邮费模板已下载");
}

function parseShippingExcel(arrayBuffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel 库未加载，请检查网络");
  }
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("Excel 没有工作表");

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) throw new Error("工作表为空");

  const headerRow = rows[0];
  // 解析档位列（第2列起到倒数第2列），最后两列是阈值和折扣
  const tierKeys = [];
  for (let c = 1; c < headerRow.length - 2; c++) {
    const val = headerRow[c];
    if (val !== "" && !isNaN(Number(val)) && Number(val) > 0) {
      tierKeys.push(String(Number(val)));
    }
  }
  if (!tierKeys.length) throw new Error("未找到档位列，表头第 2 列起应为数字档位");

  const regions = [];
  const errors = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;
    const name = String(row[0] || "").trim();
    if (!name) continue;

    const basePrices = {};
    for (let i = 0; i < tierKeys.length; i++) {
      const tier = tierKeys[i];
      const val = row[1 + i];
      const price = val === "" || val == null ? 0 : Number(val);
      if (isNaN(price)) {
        errors.push(`「${name}」的 ${tier} 档价格无效，已按 0 处理`);
        basePrices[tier] = 0;
      } else {
        basePrices[tier] = price;
      }
    }

    // 解析最后两列：小面积折扣阈值、折扣系数
    const thresholdIdx = 1 + tierKeys.length;
    const discountIdx = thresholdIdx + 1;
    const threshold = row[thresholdIdx] !== "" && row[thresholdIdx] != null
      ? Number(row[thresholdIdx]) : 2500;
    const discount = row[discountIdx] !== "" && row[discountIdx] != null
      ? Number(row[discountIdx]) : 0.8;
    if (isNaN(threshold)) errors.push(`「${name}」的折扣阈值无效，已按 2500 处理`);
    if (isNaN(discount)) errors.push(`「${name}」的折扣系数无效，已按 0.8 处理`);

    regions.push({
      id: "region" + (regions.length + 1),
      name,
      basePrices,
      smallAreaThreshold: isNaN(threshold) ? 2500 : threshold,
      discount: isNaN(discount) ? 0.8 : discount
    });
  }

  if (!regions.length) throw new Error("没有可导入的邮费配置");
  return { regions, errors };
}

function importShippingExcel(file) {
  if (!file) return;
  if (typeof XLSX === "undefined") {
    loadSheetJS().then(() => importShippingExcel(file)).catch(() => { setShippingExcelStatus("Excel 库加载失败，请检查网络", true); showToast("Excel 库加载失败"); });
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const { regions, errors } = parseShippingExcel(e.target.result);
      // 直接替换 SHIPPING_CONFIG 的内容（保持引用不变）
      SHIPPING_CONFIG.length = 0;
      regions.forEach(r => SHIPPING_CONFIG.push(r));
      // 重新渲染地区下拉
      if (els.region) {
        els.region.innerHTML = '<option value="">请选择地区</option>' +
          SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
        if (SHIPPING_CONFIG.length > 0) els.region.value = SHIPPING_CONFIG[0].id;
      }
      onCalculate();
      const successMsg = `成功导入 ${regions.length} 个地区的邮费：${regions.map(r => r.name).join("、")}`;
      const errMsg = errors.length ? `<br><span style="color:var(--danger)">警告：${errors.join("；")}</span>` : "";
      setShippingExcelStatus(successMsg + errMsg, false);
      showToast("邮费 Excel 导入成功");
    } catch (err) {
      setShippingExcelStatus("导入失败：" + err.message, true);
      showToast("邮费 Excel 导入失败");
    }
    if (els.importShippingExcelFile) els.importShippingExcelFile.value = "";
  };
  reader.onerror = () => {
    setShippingExcelStatus("文件读取失败", true);
    if (els.importShippingExcelFile) els.importShippingExcelFile.value = "";
  };
  reader.readAsArrayBuffer(file);
}

// -------------------- 页面切换增强 --------------------
function switchPage(pageName) {
  els.navBtns.forEach(btn => {
    btn.classList.toggle("active", btn.dataset.page === pageName);
  });
  els.pages.forEach(page => {
    page.classList.toggle("active", page.id === "page-" + pageName);
  });
  if (pageName === "table") {
    renderPriceListSelector();
    renderPriceTable();
  } else if (pageName === "profile") {
    renderLevelSettings();
    loadProfileToUI();
    renderSnapshots();
    renderHistory();
  }
}

// -------------------- 事件绑定 --------------------
function bindEvents() {
  els.navBtns.forEach(btn => {
    btn.addEventListener("click", () => switchPage(btn.dataset.page));
  });

  [els.tier, els.region].forEach(el => {
    if (!el) return;
    el.addEventListener("change", onCalculate);
  });
  if (els.rope) {
    els.rope.querySelectorAll('input[name="rope"]').forEach(radio => {
      radio.addEventListener("change", onCalculate);
    });
  }

  // 纸张数量变化时重渲染纸张卡片再计算
  if (els.sheetCount) {
    els.sheetCount.addEventListener("input", () => {
      renderSheets();
      onCalculate();
    });
  }

  if (els.searchInput) els.searchInput.addEventListener("input", debounce(renderPriceTable, 200));
  if (els.prevPaper) els.prevPaper.addEventListener("click", prevPaper);
  if (els.nextPaper) els.nextPaper.addEventListener("click", nextPaper);
  if (els.paperSelector) els.paperSelector.addEventListener("change", () => {
    const idx = parseInt(els.paperSelector.value, 10);
    if (!isNaN(idx) && idx >= 0 && idx < getPapersByPriceList(CURRENT_PRICE_LIST_ID).length) {
      currentPaperIndex = idx;
      renderPriceTable();
    }
  });
  if (els.priceListSelector) els.priceListSelector.addEventListener("change", onPriceListChange);
  if (els.deletePriceListBtn) els.deletePriceListBtn.addEventListener("click", onDeletePriceList);

  // 个人主页事件
  if (els.addLevelBtn) els.addLevelBtn.addEventListener("click", addCustomerLevel);
  if (els.resetLevelBtn) els.resetLevelBtn.addEventListener("click", resetCustomerLevels);
  if (els.saveProfileBtn) els.saveProfileBtn.addEventListener("click", saveProfile);
  if (els.exportProfileBtn) els.exportProfileBtn.addEventListener("click", exportProfile);

  // 模式切换
  els.modeBtns.forEach(btn => {
    btn.addEventListener("click", () => switchCalcMode(btn.dataset.mode));
  });

  // 报价小数位数：实时保存并立即重算
  if (els.decimalPlaces) {
    els.decimalPlaces.addEventListener("input", () => {
      APP_PROFILE.decimalPlaces = parseDecimalPlaces(els.decimalPlaces.value);
      saveToStorage("appProfile", APP_PROFILE);
      onCalculate();
    });
    els.decimalPlaces.addEventListener("change", () => {
      APP_PROFILE.decimalPlaces = parseDecimalPlaces(els.decimalPlaces.value);
      els.decimalPlaces.value = APP_PROFILE.decimalPlaces;
      saveToStorage("appProfile", APP_PROFILE);
      onCalculate();
      showToast("小数位数已更新为 " + APP_PROFILE.decimalPlaces + " 位");
    });
  }

  // 数据管理标签页
  if (els.tabs) els.tabs.forEach(tab => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  // 数据管理按钮
  if (els.exportDataBtn) els.exportDataBtn.addEventListener("click", exportFullData);
  if (els.importDataBtn) els.importDataBtn.addEventListener("click", () => els.importFile.click());
  if (els.importFile) els.importFile.addEventListener("change", e => importFullData(e.target.files[0]));
  if (els.resetToDefaultsBtn) els.resetToDefaultsBtn.addEventListener("click", resetToDefaults);
  const resetAllBtn = document.getElementById("resetAllSettingsBtn");
  if (resetAllBtn) resetAllBtn.addEventListener("click", resetAllLocalSettings);
  if (els.exportLocalBackupBtn) els.exportLocalBackupBtn.addEventListener("click", exportLocalBackup);
  if (els.importLocalBackupBtn) els.importLocalBackupBtn.addEventListener("click", () => els.importLocalBackupFile.click());
  if (els.importLocalBackupFile) els.importLocalBackupFile.addEventListener("change", e => importLocalBackup(e.target.files[0]));
  if (els.exportExcelBtn) els.exportExcelBtn.addEventListener("click", exportPaperExcel);
  if (els.importExcelBtn) els.importExcelBtn.addEventListener("click", () => els.importExcelFile.click());
  if (els.importExcelFile) els.importExcelFile.addEventListener("change", e => importPaperExcel(e.target.files[0]));
  if (els.downloadTemplateBtn) els.downloadTemplateBtn.addEventListener("click", downloadPaperTemplate);
  if (els.exportRopeExcelBtn) els.exportRopeExcelBtn.addEventListener("click", exportRopeExcel);
  if (els.importRopeExcelBtn) els.importRopeExcelBtn.addEventListener("click", () => els.importRopeExcelFile.click());
  if (els.importRopeExcelFile) els.importRopeExcelFile.addEventListener("change", e => importRopeExcel(e.target.files[0]));
  if (els.downloadRopeTemplateBtn) els.downloadRopeTemplateBtn.addEventListener("click", downloadRopeTemplate);
  if (els.exportShippingExcelBtn) els.exportShippingExcelBtn.addEventListener("click", exportShippingExcel);
  if (els.importShippingExcelBtn) els.importShippingExcelBtn.addEventListener("click", () => els.importShippingExcelFile.click());
  if (els.importShippingExcelFile) els.importShippingExcelFile.addEventListener("change", e => importShippingExcel(e.target.files[0]));
  if (els.downloadShippingTemplateBtn) els.downloadShippingTemplateBtn.addEventListener("click", downloadShippingTemplate);
  if (els.createSnapshotBtn) els.createSnapshotBtn.addEventListener("click", createSnapshot);
  if (els.clearSnapshotBtn) els.clearSnapshotBtn.addEventListener("click", clearSnapshots);
  if (els.addHistoryPlaceholderBtn) els.addHistoryPlaceholderBtn.addEventListener("click", addHistoryPlaceholder);
  if (els.clearHistoryBtn) els.clearHistoryBtn.addEventListener("click", clearHistory);

  // 点击页面其他区域关闭纸张材质下拉
  document.addEventListener("click", () => closeAllPaperDropdowns());

  // 临时毛利系数：输入时实时渲染报价卡片（不触发 onCalculate）
  if (els.customCoeffInput) {
    els.customCoeffInput.addEventListener("input", renderCustomCoeffCard);
  }

  // 邮费快速修改：输入时实时渲染修改后报价卡片
  if (els.shippingOverrideInput) {
    els.shippingOverrideInput.addEventListener("input", renderShippingOverrideCards);
  }
  // 清除邮费修改
  if (els.shippingOverrideClear) {
    els.shippingOverrideClear.addEventListener("click", () => {
      if (els.shippingOverrideInput) els.shippingOverrideInput.value = "";
      renderShippingOverrideCards();
    });
  }

  // 禁用所有 number 类型输入框的鼠标滚轮调整，防止误操作
  document.addEventListener("wheel", e => {
    if (e.target && e.target.type === "number") {
      e.preventDefault();
      e.target.blur();
    }
    // 阻止纸张下拉框滚动穿透到页面
    const opts = e.target.closest('.paper-options');
    if (opts) {
      const { scrollTop, scrollHeight, clientHeight } = opts;
      const isAtTop = scrollTop <= 0;
      const isAtBottom = scrollTop + clientHeight >= scrollHeight - 2;
      if ((isAtTop && e.deltaY < 0) || (isAtBottom && e.deltaY > 0)) {
        e.preventDefault();
      }
    }
  }, { passive: false });
}

// -------------------- 启动 --------------------
// 兜底：任何 init 步骤失败不影响其他步骤，错误会显示在控制台
(function safeInit() {
  const steps = [
    ["initOptions", initOptions],
    ["updateDefaultTierOptions", updateDefaultTierOptions],
    ["renderSheets", renderSheets],
    ["bindEvents", bindEvents],
    ["clearResult", clearResult],
    ["loadProfileToUI", loadProfileToUI],
    ["renderLevelSettings", renderLevelSettings],
    ["renderSnapshots", renderSnapshots],
    ["renderHistory", renderHistory],
    ["renderPriceListSelector", renderPriceListSelector]
  ];
  steps.forEach(([name, fn]) => {
    try { if (typeof fn === "function") fn(); }
    catch (e) { console.error(`[init] ${name} 失败:`, e); }
  });

  // 应用个人主页的默认档位设置
  try {
    if (APP_PROFILE.defaultTier && els.tier) {
      const tierExists = Array.from(els.tier.options).some(o => o.value === String(APP_PROFILE.defaultTier));
      if (tierExists) els.tier.value = String(APP_PROFILE.defaultTier);
    }
  } catch (e) { console.error("[init] 默认档位失败:", e); }

  // 同步直接系数模式按钮状态（默认直接系数模式）
  try {
    els.modeBtns.forEach(btn => {
      btn.classList.toggle("active", btn.dataset.mode === calcMode);
    });
    // 同步直接系数模式下吊绳/邮费输入区域的显隐
    const ropeGroup = els.rope ? els.rope.closest(".form-group") : null;
    const regionGroup = els.region ? els.region.closest(".form-group") : null;
    if (calcMode === "direct") {
      if (ropeGroup) ropeGroup.style.display = "none";
      if (regionGroup) regionGroup.style.display = "none";
    } else {
      if (ropeGroup) ropeGroup.style.display = "";
      if (regionGroup) regionGroup.style.display = "";
    }
  } catch (e) { console.error("[init] 模式同步失败:", e); }

  // 应用个人主页的默认纸张材质设置
  try {
    if (APP_PROFILE.defaultPaperId) {
      const idx = getPapersByPriceList(CURRENT_PRICE_LIST_ID).findIndex(p => p.id === APP_PROFILE.defaultPaperId);
      if (idx >= 0) {
        currentPaperIndex = idx;
        renderPriceTable();
      }
    }
  } catch (e) { console.error("[init] 默认纸张失败:", e); }

  // 初始化完成后触发一次计算，确保页面有结果显示
  try { onCalculate(); }
  catch (e) { console.error("[init] onCalculate 失败:", e); }
})();