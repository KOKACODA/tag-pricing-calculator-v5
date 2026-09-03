// ============================================================
// KOKALabel报价系统 v7.9.0 - 主程序（计算 + 渲染 + 交互 + 初始化）
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
 * v7.9: SheetJS 已本地化（js/vendor/xlsx.full.min.js，由 index.html 同步引入），
 * XLSX 全局变量在业务脚本执行前即可用。本函数作为兜底：
 *   - 若 XLSX 已存在（本地脚本加载成功）→ 立即 resolve；
 *   - 若不存在（异常场景）→ 动态加载本地 vendor 文件，确保离线双击打开也能用。
 */
let _sheetjsLoaded = false;
let _sheetjsLoading = null;

function loadSheetJS() {
  if (typeof XLSX !== "undefined") {
    _sheetjsLoaded = true;
    return Promise.resolve();
  }
  if (_sheetjsLoaded) return Promise.resolve();
  if (_sheetjsLoading) return _sheetjsLoading;
  _sheetjsLoading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = "js/vendor/xlsx.full.min.js";
    script.onload = () => { _sheetjsLoaded = true; _sheetjsLoading = null; resolve(); };
    script.onerror = () => { _sheetjsLoading = null; reject(new Error("SheetJS 本地库加载失败")); };
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

// v6.14：整数价格去尾零显示（如 90 → "90"，90.5 → "90.50"），用于折后价合计的计算过程展示
function formatPriceRaw(value) {
  const n = Number(value);
  if (typeof n !== "number" || isNaN(n)) return "-";
  return Number.isInteger(n) ? String(n) : formatMoney(n);
}

// 工艺与吊绳选项右上角统一展示 1000 档的初始价格，不随当前报价档位变化。
const INITIAL_OPTION_PRICE_TIER = 1000;

function getInitialOptionPrice(prices) {
  if (!prices || prices[INITIAL_OPTION_PRICE_TIER] == null || prices[INITIAL_OPTION_PRICE_TIER] === "") {
    return "";
  }
  const value = Number(prices[INITIAL_OPTION_PRICE_TIER]);
  return Number.isFinite(value) ? String(value) : "";
}

function normalizePaperSearchText(value) {
  return String(value == null ? "" : value).toLocaleLowerCase().replace(/\s+/g, "");
}

/**
 * 纸张模糊匹配：每个输入词都需要出现在简称或全称中。
 */
function paperMatchesSearch(paper, query) {
  if (!paper) return false;
  const tokens = String(query == null ? "" : query)
    .toLocaleLowerCase()
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map(normalizePaperSearchText);
  if (!tokens.length) return true;

  const haystack = normalizePaperSearchText(`${paper.shortName || ""} ${paper.name || ""}`);
  return tokens.every(token => haystack.includes(token));
}

/**
 * 默认客户报价显隐只作用于标准报价；直接系数结果始终显示。
 */
function shouldShowDefaultQuoteCards(mode, visiblePreference) {
  return mode !== "standard" || visiblePreference !== false;
}

function padQuoteDatePart(value) {
  return String(value).padStart(2, "0");
}

function formatQuoteRecordNumber(dateValue) {
  const date = new Date(dateValue);
  const safeDate = isNaN(date.getTime()) ? new Date() : date;
  return "BJ-" +
    safeDate.getFullYear() +
    padQuoteDatePart(safeDate.getMonth() + 1) +
    padQuoteDatePart(safeDate.getDate()) + "-" +
    padQuoteDatePart(safeDate.getHours()) +
    padQuoteDatePart(safeDate.getMinutes()) +
    padQuoteDatePart(safeDate.getSeconds());
}

function cloneQuoteData(value) {
  if (value == null) return value;
  return JSON.parse(JSON.stringify(value));
}

/**
 * 构造不可受后续页面状态变化影响的报价历史记录。
 */
function buildQuoteHistoryRecord({
  id,
  createdAt,
  customerName,
  note,
  mode,
  priceList,
  inputs,
  result,
  ropeName,
  regionName
}) {
  const date = new Date(createdAt);
  const safeDate = isNaN(date.getTime()) ? new Date() : date;
  const recordNo = formatQuoteRecordNumber(safeDate);
  const safeCustomerName = String(customerName || "").trim();
  const snapshot = cloneQuoteData(result || {});
  snapshot.ropeName = String(ropeName || "");
  snapshot.regionName = String(regionName || "");

  return {
    id: String(id || recordNo),
    schemaVersion: 2,
    recordNo,
    title: safeCustomerName || recordNo,
    customerName: safeCustomerName,
    note: String(note || "").trim(),
    createdAt: safeDate.toISOString(),
    mode: mode === "direct" ? "direct" : "standard",
    priceList: cloneQuoteData(priceList || {}),
    inputs: cloneQuoteData(inputs || {}),
    snapshot
  };
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
 * v7.3：强制面积规格映射（不读取表格数据）。
 * 面积（加出血后）范围 → 强制规格代码：
 *   4000-4999 → 005
 *   5000-5499 → 055
 *   5500-6000 → 006（v8.2：6000 边界也强制 006，避免回落到 055 导致价格下降）
 * 命中返回对应规格副本；未命中返回 null（调用方按原逻辑处理）。
 */
function forceSpecByArea(paper, area) {
  if (!paper || !Array.isArray(paper.specs)) return null;
  let targetCode = null;
  if (area >= 4000 && area <= 4999) targetCode = "005";
  else if (area >= 5000 && area <= 5499) targetCode = "055";
  else if (area >= 5500 && area <= 6000) targetCode = "006";
  if (!targetCode) return null;
  const spec = paper.specs.find(s => s.code === targetCode);
  if (!spec) return null;
  return { ...spec };
}

/**
 * 根据有效面积向上匹配尺寸规格。
 * 面积超过 10000 时使用最大规格（code "100"）并附带面积系数；007 与 008 合并显示为 007/008。
 * v7.3：面积 4000-6000 时强制映射 005/055/006，不读取表格数据。
 */
function matchSpec(paper, area) {
  // v7.3：强制面积规格映射（不读取表格数据）
  const forced = forceSpecByArea(paper, area);
  if (forced) return forced;
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
  // v8.5：空字符串视为无价格（Number("") === 0，会导致空档位被误判为价格 0）
  return v != null && v !== "" && !isNaN(Number(v));
}

/**
 * 标准计价流程：支持多纸张、每张纸独立吊牌展开尺寸、统一批量档位。
 */
function calculate(inputs) {
  const { sheetCount, tier, sheets, ropeId, regionId, mode, shippingWeight, manualShippingPrice } = inputs;
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
  // v7.0：批量直接报价累计（固定价，不打折、不乘系数）
  let batchDirectTotal = 0;       // 批量直接报价纸张价合计
  let batchDirectCraftTotal = 0;  // 批量直接报价纸张的工艺费用合计
  const warnings = [];
  const sheetDetails = [];
  // 是否有任一纸张/工艺在所选档位没有定价（用于显示"无该批量定价"占位）
  let hasMissingTier = false;
  // 是否有任一纸张面积超过 10000 触发面积系数
  let hasAreaCoefficient = false;
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

    // v7.0：批量直接报价判断
    // 纸张配置了 batchDirect 且面积（出血后）在最大面积范围内 → 直接按批量报价
    // 价格固定：不打折、不乘直接系数、不乘面积系数；工艺费用直接叠加
    const bd = paper.batchDirect;
    const isBatchDirect = !!(bd && bd.maxArea > 0 && singleArea <= bd.maxArea);
    let batchDirectPrice = null;
    if (isBatchDirect && hasExactTier(bd.prices, tier)) {
      batchDirectPrice = Number(bd.prices[tier]);
    }

    const areaCoeff = isBatchDirect ? 1 : (spec.areaCoefficient || 1);
    if (areaCoeff > 1) hasAreaCoefficient = true;

    // 纸张基础价（未乘面积系数）
    // v7.0：批量直接报价 → 直接用批量报价价格（不打折、不乘系数）
    const baseOriginalPrice = isBatchDirect
      ? batchDirectPrice
      : (hasExactTier(spec.prices, tier) ? Number(spec.prices[tier]) : null);
    // v6.14：直接系数模式：有直接系数的纸张用原价（不打折），无直接系数的纸张按折扣价（原价 × discount）
    // 标准报价模式：纸张乘 discount（打折），再 × 客户等级系数 = 最终报价
    const baseUnitPrice = isBatchDirect
      ? batchDirectPrice
      : (hasExactTier(spec.prices, tier)
        ? Number(spec.prices[tier]) * (isDirect ? (paperHasDirectCoeff(paper) ? 1 : paper.discount) : paper.discount)
        : null);
    // 纸张最终价（乘面积系数后）
    const paperOriginalPrice = baseOriginalPrice != null ? baseOriginalPrice * areaCoeff : null;
    const paperUnitPrice = baseUnitPrice != null ? baseUnitPrice * areaCoeff : null;

    if (paperUnitPrice == null) {
      hasMissingTier = true;
      warnings.push(`「${paper.shortName || paper.name}」${isBatchDirect ? "无 " + tier + " 张批量直接报价" : "无 " + tier + " 张批量定价"}`);
    } else {
      if (isBatchDirect) {
        batchDirectTotal += paperUnitPrice;
      } else {
        paperTotal += paperUnitPrice;
        paperOriginalTotal += paperOriginalPrice;
      }
    }

    // 工艺费用：每个工艺独立检查档位，无值则跳过并提示
    const crafts = CRAFT_CONFIG[sheet.paperId] || [];
    const sheetCraftDetails = [];
    // v7.2：每张纸自己的工艺费用合计（直接系数模式下：纸张价 + 工艺价 后再乘直接系数）
    let sheetCraftTotal = 0;
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
          // v7.0：批量直接报价纸张的工艺费用单独累计（直接叠加，不乘任何系数）
          if (isBatchDirect) {
            batchDirectCraftTotal += cPrice;
          } else {
            craftTotal += cPrice;
            sheetCraftTotal += cPrice;
          }
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
      discount: paper.discount != null ? paper.discount : 1,
      missing: paperUnitPrice == null,
      // v7.0：批量直接报价标记
      isBatchDirect,
      batchDirectPrice,
      // v7.2：该纸张自己的工艺费用合计（用于直接系数模式：纸张 + 工艺 再乘系数）
      sheetCraftTotal,
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
  // v8.0：三种情况 - A 固定档直接查价；B >10000 重量×系数；C 中间档手动输入
  let shippingPrice = null;
  let shippingMode = null;        // "fixed" | "weight" | "manual"
  let shippingWeightUsed = null;  // 用于展示的重量值
  if (!isDirect) {
    if (hasExactTier(region.basePrices, tier)) {
      // A：固定档直接查价
      shippingPrice = Number(region.basePrices[tier]);
      shippingMode = "fixed";
    } else if (tier > 10000) {
      // B：数量超过 10000，邮费 = 吊牌重量(kg) × 地区系数
      const w = Number(shippingWeight);
      if (shippingWeight != null && !isNaN(w) && w > 0) {
        shippingPrice = w * region.overTierCoeff;
        shippingWeightUsed = w;
        shippingMode = "weight";
      } else {
        shippingMode = "weight";
        hasMissingTier = true;
        warnings.push(`地区「${region.name}」数量 ${tier} 张超过 10000，请填写吊牌重量(kg)计算邮费`);
      }
    } else {
      // C：中间档（2500/3000/4000 等），手动输入邮费
      const m = Number(manualShippingPrice);
      if (manualShippingPrice != null && !isNaN(m) && m > 0) {
        shippingPrice = m;
        shippingMode = "manual";
      } else {
        shippingMode = "manual";
        hasMissingTier = true;
        warnings.push(`地区「${region.name}」无 ${tier} 张批量定价，请手动填写邮费`);
      }
    }
  }

  // 成本合计：直接系数模式 = 纸张 + 工艺；标准模式 = 纸张 + 工艺 + 吊绳 + 邮费
  // v7.0：增加批量直接报价合计（批量纸张价 + 批量纸张工艺费，直接叠加）
  let cost, costKnown, costIncomplete;
  if (isDirect) {
    cost = paperTotal + craftTotal + batchDirectTotal + batchDirectCraftTotal;
    costKnown = true;
    costIncomplete = false;
  } else {
    // v8.2：costKnown 纳入全部缺价状态（hasMissingTier 已覆盖纸张/工艺/吊绳/邮费缺价）
    costKnown = !hasMissingTier;
    cost = costKnown
      ? (paperTotal + craftTotal + ropePrice + shippingPrice + batchDirectTotal + batchDirectCraftTotal)
      : (paperTotal + craftTotal + (ropePrice || 0) + (shippingPrice || 0) + batchDirectTotal + batchDirectCraftTotal);
    costIncomplete = !costKnown;
  }

  // 报价系数：直接系数模式每张纸独立系数（各等级分别计算），标准模式用 CUSTOMER_LEVELS 统一系数
  let pricesByLevel;
  if (isDirect) {
    // v6.14：每张纸按各自 Sheet 直接系数计算；无直接系数的纸张不乘系数，按折扣价（原价 × discount）计入
    // 工艺费用不乘系数，直接累加（防止多纸张混合时工艺计算错误）
    // v7.0：批量直接报价纸张不参与系数计算，价格直接累加
    pricesByLevel = DIRECT_COEFF_LEVELS.map((level, li) => {
      const paperDetails = sheetDetails.map(sd => {
        const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === sd.paperId);
        if (sd.isBatchDirect) {
          // v7.0：批量直接报价纸张 → 不乘系数，不参与系数计算，直接显示固定价格
          // contributedPrice 置 0：价格由 batchDirectTotal 统一累加，避免重复计数
          return {
            paperId: sd.paperId,
            paperName: sd.paperName,
            unitPrice: sd.unitPrice,
            originalUnitPrice: sd.originalUnitPrice,
            coefficient: 1,
            hasDirectCoeff: false,
            isBatchDirect: true,
            batchDirectPrice: sd.batchDirectPrice,
            discount: paper && paper.discount != null ? paper.discount : 1,
            contributedPrice: sd.unitPrice != null ? 0 : null
          };
        }
        let coeff = null;
        let hasDirectCoeff = false;
        if (paper && paperHasDirectCoeff(paper)) {
          const coeffs = getDirectCoeffsForTier(tier, paper);
          if (coeffs && coeffs[li]) {
            coeff = coeffs[li].coefficient;
            hasDirectCoeff = true;
          }
        }
        let contributedPrice = null;
        // v7.2：工艺价先加到该纸张价上，再乘直接系数：(纸张价 + 工艺价) × 系数
        // 无直接系数的纸张系数为 1，等价于 纸张价 + 工艺价
        const craftOfSheet = sd.sheetCraftTotal || 0;
        if (coeff == null) {
          // v6.14：无直接系数 → 不乘系数，unitPrice 已是折扣价（原价 × discount）
          coeff = 1;
          contributedPrice = sd.unitPrice != null ? (sd.unitPrice + craftOfSheet) : null;
        } else {
          contributedPrice = sd.unitPrice != null ? (sd.unitPrice + craftOfSheet) * coeff : null;
        }
        return {
          paperId: sd.paperId,
          paperName: sd.paperName,
          unitPrice: sd.unitPrice,
          originalUnitPrice: sd.originalUnitPrice,
          coefficient: coeff,
          hasDirectCoeff,
          isBatchDirect: false,
          batchDirectPrice: null,
          discount: paper && paper.discount != null ? paper.discount : 1,
          // v7.2：该纸张自己的工艺费用（用于明细卡片显示）
          craftTotal: craftOfSheet,
          contributedPrice
        };
      });
      // v7.2：工艺费已并入各纸张（纸张+工艺）× 系数，不再单独累加 craftTotal
      const total = paperDetails.reduce((sum, p) => sum + (p.contributedPrice != null ? p.contributedPrice : 0), 0) + batchDirectTotal + batchDirectCraftTotal;
      const incomplete = paperDetails.some(p => p.contributedPrice == null);
      return {
        levelId: level.id,
        levelName: level.name,
        coefficient: null, // 无统一系数（每张纸独立）
        price: total,
        costIncomplete: incomplete,
        paperDetails,
        craftTotal,
        // v7.0：批量直接报价合计（用于明细卡片显示）
        batchDirectTotal,
        batchDirectCraftTotal
      };
    });
  } else {
    pricesByLevel = CUSTOMER_LEVELS.map(level => ({
      levelId: level.id,
      levelName: level.name,
      coefficient: level.coefficient,
      price: cost * level.coefficient
    }));
  }

  return {
    sheetDetails,
    tier,
    paperTotal,
    paperOriginalTotal,
    craftTotal,
    ropePrice,
    shippingPrice,
    shippingMode,
    shippingWeightUsed,
    cost,
    costIncomplete,
    hasMissingTier,
    hasAreaCoefficient,
    pricesByLevel,
    warnings,
    // v7.0：批量直接报价结果
    batchDirectTotal,
    batchDirectCraftTotal
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
  // v6.8：每纸独立临时直接系数
  tempCoeffBar: document.getElementById("tempCoeffBar"),
  tempCoeffInputs: document.getElementById("tempCoeffInputs"),
  tempCoeffResults: document.getElementById("tempCoeffResults"),
  shippingOverrideBar: document.getElementById("shippingOverrideBar"),
  shippingOverrideInput: document.getElementById("shippingOverrideInput"),
  shippingOverrideClear: document.getElementById("shippingOverrideClear"),
  shippingOverrideLabel: document.getElementById("shippingOverrideLabel"),
  shippingOverrideCost: document.getElementById("shippingOverrideCost"),
  shippingOverrideCostValue: document.getElementById("shippingOverrideCostValue"),
  shippingOverrideCards: document.getElementById("shippingOverrideCards"),
  searchInput: document.getElementById("searchInput"),
  priceTable: document.getElementById("priceTable"),
  craftTable: document.getElementById("craftTable"),
  paperDiscount: document.getElementById("paperDiscount"),
  directCoeffTable: document.getElementById("directCoeffTable"),
  directCoeffWrap: document.getElementById("directCoeffWrap"),
  // v7.0：批量直接报价表
  batchDirectTable: document.getElementById("batchDirectTable"),
  batchDirectWrap: document.getElementById("batchDirectWrap"),
  // v7.9：吊绳/邮费价格表（就地编辑）
  ropePriceTable: document.getElementById("ropePriceTable"),
  shippingPriceTable: document.getElementById("shippingPriceTable"),
  tableMeta: document.getElementById("tableMeta"),
  prevPaper: document.getElementById("prevPaper"),
  nextPaper: document.getElementById("nextPaper"),
  paperPageInfo: document.getElementById("paperPageInfo"),
  paperSelector: document.getElementById("paperSelector"),
  priceListSelector: document.getElementById("priceListSelector"),
  quickPriceListSelector: document.getElementById("quickPriceListSelector"),
  defaultQuoteToggle: document.getElementById("defaultQuoteToggle"),
  defaultQuoteToggleState: document.getElementById("defaultQuoteToggleState"),
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
  clearHistoryBtn: document.getElementById("clearHistoryBtn"),
  quoteCustomerName: document.getElementById("quoteCustomerName"),
  quoteOrderNote: document.getElementById("quoteOrderNote"),
  saveQuoteBtn: document.getElementById("saveQuoteBtn"),
  openQuoteHistoryBtn: document.getElementById("openQuoteHistoryBtn"),
  historyDetailDialog: document.getElementById("historyDetailDialog"),
  historyDetailTitle: document.getElementById("historyDetailTitle"),
  historyDetailContent: document.getElementById("historyDetailContent"),
  historyDetailLoadBtn: document.getElementById("historyDetailLoadBtn"),
  historyDetailCloseBtn: document.getElementById("historyDetailCloseBtn"),
  historyDetailCancelBtn: document.getElementById("historyDetailCancelBtn"),
  // v8.0：邮费输入对话框
  shippingWeightDialog: document.getElementById("shippingWeightDialog"),
  shippingWeightInput: document.getElementById("shippingWeightInput"),
  shippingWeightDesc: document.getElementById("shippingWeightDesc"),
  shippingWeightConfirm: document.getElementById("shippingWeightConfirm"),
  manualShippingDialog: document.getElementById("manualShippingDialog"),
  manualShippingInput: document.getElementById("manualShippingInput"),
  manualShippingDesc: document.getElementById("manualShippingDesc"),
  manualShippingConfirm: document.getElementById("manualShippingConfirm"),
  // 模式切换
  modeBtns: document.querySelectorAll(".mode-btn"),
  // 报价结果中需要模式控制的行
  resRopePriceRow: null, // 稍后在 onCalculate 中动态获取
  resShippingPriceRow: null,
  // 默认报价标签
  defaultPriceLabel: document.getElementById("defaultPriceLabel")
};

let currentPaperIndex = 2;
let toastTimer = null;
let sheetsState = [];
let activeHistoryRecordId = null;
// 计算模式：默认直接系数计算（v6.1 起），持久化到 localStorage
let calcMode = loadFromStorage("currentCalcMode", "direct"); // "standard" | "direct"
let defaultQuoteVisible = loadFromStorage("defaultQuoteVisible", true) !== false;
// v8.0：邮费输入缓存（会话级，不持久化，刷新后重新询问）
let shippingWeightCache = {};   // regionId -> 吊牌重量(kg)
let manualShippingCache = {};   // regionId_tier -> 手动邮费金额(元)

function applyDefaultQuoteVisibility() {
  const showCards = shouldShowDefaultQuoteCards(calcMode, defaultQuoteVisible);
  if (els.defaultPriceLabel) els.defaultPriceLabel.classList.toggle("default-quote-hidden", !showCards);
  if (els.priceCards) els.priceCards.classList.toggle("default-quote-hidden", !showCards);
  // v8.3：邮费快速修改区（标签/成本/卡片）也随默认报价显隐
  if (els.shippingOverrideLabel) els.shippingOverrideLabel.classList.toggle("default-quote-hidden", !showCards);
  if (els.shippingOverrideCost) els.shippingOverrideCost.classList.toggle("default-quote-hidden", !showCards);
  if (els.shippingOverrideCards) els.shippingOverrideCards.classList.toggle("default-quote-hidden", !showCards);
  if (els.defaultQuoteToggle) {
    els.defaultQuoteToggle.setAttribute("aria-checked", String(defaultQuoteVisible));
    els.defaultQuoteToggle.classList.toggle("is-off", !defaultQuoteVisible);
    els.defaultQuoteToggle.title = defaultQuoteVisible
      ? "标准报价中的默认客户报价当前显示"
      : "标准报价中的默认客户报价当前隐藏";
  }
  if (els.defaultQuoteToggleState) {
    els.defaultQuoteToggleState.textContent = defaultQuoteVisible ? "显示" : "隐藏";
  }
}

function toggleDefaultQuoteVisibility() {
  defaultQuoteVisible = !defaultQuoteVisible;
  saveToStorage("defaultQuoteVisible", defaultQuoteVisible);
  applyDefaultQuoteVisibility();
  showToast(defaultQuoteVisible ? "标准默认报价已显示" : "标准默认报价已隐藏");
}

// -------------------- 初始化下拉选项 --------------------
// 根据 APP_PROFILE.defaultRope 生成吊绳单选项 HTML（顶层作用域，供 initOptions 和 rebuildRopeUI 共用）
function renderRopeRadios() {
  const fallback = "rope1"; // 默认「普通吊绳」
  const desired = APP_PROFILE.defaultRope || fallback;
  const ropes = ROPE_CONFIG;
  const hasDesired = ropes.some(r => r.id === desired);
  const checkedId = hasDesired ? desired : (ropes.find(r => r.id === fallback) ? fallback : (ropes[0] && ropes[0].id));
  return ropes.map(r => {
    const initialPrice = getInitialOptionPrice(r.prices);
    return `
      <label class="rope-item">
        <input type="radio" name="rope" value="${escapeHtml(r.id)}"${r.id === checkedId ? " checked" : ""} />
        <span class="option-label">${escapeHtml(r.name)}</span>
        ${initialPrice ? `<sup class="option-initial-price" title="1000档初始价" aria-label="1000档初始价 ${initialPrice}">${initialPrice}</sup>` : ""}
      </label>
    `;
  }).join("");
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
  // v8.2：默认纸张优先使用个人主页设置的默认纸张（若存在于当前报价表），否则用第一张纸
  const defaultPaperId = (APP_PROFILE.defaultPaperId && currentPapers.some(p => p.id === APP_PROFILE.defaultPaperId))
    ? APP_PROFILE.defaultPaperId
    : (currentPapers[0]?.id || "");
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
      ? crafts.map(c => {
          const initialPrice = getInitialOptionPrice(c.prices);
          return `
            <label class="craft-item">
              <input type="checkbox" value="${escapeHtml(c.id)}" data-sheet="${index}" data-craft="${escapeHtml(c.id)}"${sheet.craftIds.includes(c.id) ? " checked" : ""} />
              <span class="option-label">${escapeHtml(c.name)}</span>
              ${initialPrice ? `<sup class="option-initial-price" title="1000档初始价" aria-label="1000档初始价 ${initialPrice}">${initialPrice}</sup>` : ""}
            </label>
          `;
        }).join("")
      : '<div class="craft-empty">该纸张暂无附加工艺</div>';

    const triggerText = currentPaper ? (currentPaper.shortName || currentPaper.name) : "无可用纸张";
    const triggerDesc = currentPaper ? currentPaper.name : "请导入报价表";
    // v6.14：无直接系数的纸张提示只在直接系数模式下显示（标准报价模式隐藏）
    // v7.0：有批量直接报价的纸张（如 40C棉麻布）不显示"请切换为标准报价"提示，改为显示批量直接报价提示
    const hasBatchDirect = !!(currentPaper && currentPaper.batchDirect && currentPaper.batchDirect.maxArea > 0);
    const noDirectCoeff = calcMode === "direct" && !paperHasDirectCoeff(currentPaper) && !hasBatchDirect;

    card.innerHTML = `
      <div class="sheet-title">纸张 ${index + 1}</div>
      <div class="form-group">
        <label>纸张材质 <span class="hint">点击展开，选择后自动收起</span></label>
        <div class="paper-dropdown" data-sheet="${index}">
          <div class="paper-trigger" role="button" tabindex="0" aria-label="选择纸张材质">
            <div>
              <span class="paper-trigger-text">${escapeHtml(triggerText)}</span>
              <span class="paper-trigger-desc">${escapeHtml(triggerDesc)}</span>
            </div>
            <span class="paper-trigger-arrow"></span>
          </div>
          <div class="paper-options">
            <div class="paper-search">
              <input type="search" class="paper-search-input" data-sheet="${index}" placeholder="输入简称或全称搜索" autocomplete="off" aria-label="搜索纸张材质" />
            </div>
            <div class="paper-search-empty" hidden>未找到匹配的纸张材质</div>
            ${paperOptions}
          </div>
        </div>
        ${noDirectCoeff ? `
        <div class="sheet-direct-warning">
          <span class="sheet-direct-warning-icon">⚠</span>
          <span>该纸张无直接系数（最高/最低倍数为空），请切换为<strong>标准报价</strong>计算</span>
        </div>` : ""}
        ${hasBatchDirect ? `
        <div class="sheet-batchdirect-hint">
          <span>批量直接报价：面积 ≤ ${currentPaper.batchDirect.maxArea} mm²（出血后）时按固定批量价报价，工艺费用直接叠加</span>
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
    trigger.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        e.stopPropagation();
        togglePaperDropdown(dd);
      }
    });
  });
  els.sheetList.querySelectorAll(".paper-search-input").forEach(input => {
    input.addEventListener("click", e => e.stopPropagation());
    input.addEventListener("input", e => {
      e.stopPropagation();
      filterPaperOptions(input);
    });
    input.addEventListener("keydown", e => {
      e.stopPropagation();
      if (e.key === "Enter") {
        e.preventDefault();
        const firstMatch = input.closest(".paper-dropdown")?.querySelector(".paper-option:not(.is-filtered)");
        if (firstMatch) onPaperChange(firstMatch);
      } else if (e.key === "Escape") {
        closeAllPaperDropdowns();
      }
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
    requestAnimationFrame(() => {
      const searchInput = dropdown.querySelector(".paper-search-input");
      if (searchInput) {
        searchInput.value = "";
        filterPaperOptions(searchInput);
      }
      adjustDropdownPosition(dropdown);
    });
  }
}

function filterPaperOptions(input) {
  const dropdown = input.closest(".paper-dropdown");
  if (!dropdown) return;

  const currentPapers = getPapersByPriceList(CURRENT_PRICE_LIST_ID);
  let visibleCount = 0;
  dropdown.querySelectorAll(".paper-option").forEach(option => {
    const paper = currentPapers.find(item => item.id === option.dataset.paper);
    const isMatch = paperMatchesSearch(paper, input.value);
    option.classList.toggle("is-filtered", !isMatch);
    if (isMatch) visibleCount += 1;
  });

  const emptyState = dropdown.querySelector(".paper-search-empty");
  if (emptyState) emptyState.hidden = visibleCount > 0;
  if (dropdown.classList.contains("open")) adjustDropdownPosition(dropdown);
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

  // v8.0：标准模式邮费输入检测（固定档直接查价；>10000 需重量；中间档需手动输入）
  let shippingWeight = null;
  let manualShippingPrice = null;
  let needShippingWeight = false;
  let needManualShipping = false;
  if (calcMode === "standard" && regionId) {
    const region = SHIPPING_CONFIG.find(s => s.id === regionId);
    if (region && !hasExactTier(region.basePrices, tier)) {
      if (tier > 10000) {
        if (shippingWeightCache[regionId] != null) {
          shippingWeight = shippingWeightCache[regionId];
        } else {
          needShippingWeight = true;
        }
      } else {
        const manualKey = regionId + "_" + tier;
        if (manualShippingCache[manualKey] != null) {
          manualShippingPrice = manualShippingCache[manualKey];
        } else {
          needManualShipping = true;
        }
      }
    }
  }

  const result = calculate({ sheetCount, tier, sheets, ropeId, regionId, mode: calcMode, shippingWeight, manualShippingPrice });

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
        <td class="col-area">${s.area} mm²</td>
        <td class="col-code">
          <div class="code-switcher${isManual ? " manual" : ""}">
            <span class="code-text">${escapeHtml(s.code)}${isManual ? '<span class="code-manual-tag">手动</span>' : ''}</span>
            <div class="code-arrows">
              <button class="code-arrow up${!canGoUp ? ' disabled' : ''}" data-sheet-idx="${idx}" data-dir="up" ${!canGoUp ? 'disabled' : ''} title="切换到上一档代码">&#9650;</button>
              <button class="code-arrow down${!canGoDown ? ' disabled' : ''}" data-sheet-idx="${idx}" data-dir="down" ${!canGoDown ? 'disabled' : ''} title="切换到下一档代码">&#9660;</button>
            </div>
          </div>
        </td>
        <td>${s.missing ? '<span class="price-missing">无该批量定价</span>' : '¥ ' + formatMoney(s.baseOriginalUnitPrice)}</td>
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
  // v6.14：折后价合计行（showDiscountCalc=true）对有折扣的纸张显示 "¥原价 × 折扣" 计算过程
  function renderPaperTotal(getFinalPrice, getBasePrice, showDiscountCalc) {
    const details = result.sheetDetails;
    if (details.length === 1) {
      const p = getFinalPrice(details[0]);
      if (p == null) return '<span class="price-missing">无该批量定价</span>';
      if (details[0].areaCoefficient > 1) {
        const bp = getBasePrice(details[0]);
        const coeff = details[0].areaCoefficient;
        return '¥ ' + formatMoney(bp) + ' × ' + coeff + ' = <span style="color:var(--brand);font-weight:600;">¥ ' + formatMoney(p) + '</span>';
      }
      // v7.0：批量直接报价纸张不显示折扣计算过程（价格固定，不打折）
      if (showDiscountCalc && details[0].discount && details[0].discount !== 1 && !details[0].isBatchDirect) {
        return '¥ ' + formatPriceRaw(details[0].baseOriginalUnitPrice) + ' × ' + details[0].discount + ' = <span style="color:var(--brand);font-weight:600;">¥ ' + formatMoney(p) + '</span>';
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
      // v6.14：有折扣的纸张显示计算过程（原价 × 折扣）；v7.0：批量直接报价纸张不显示折扣过程
      if (showDiscountCalc && s.discount && s.discount !== 1 && !s.isBatchDirect) {
        return '¥ ' + formatPriceRaw(s.baseOriginalUnitPrice) + ' × ' + s.discount;
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
  els.resPaperPrice.innerHTML = renderPaperTotal(s => s.unitPrice, s => s.baseUnitPrice, true);
  // v7.0：工艺费用合计 = 常规工艺费 + 批量直接报价纸张的工艺费（直接叠加）
  const craftDisplayTotal = (result.craftTotal || 0) + (result.batchDirectCraftTotal || 0);
  els.resCraftPrice.innerHTML = craftDisplayTotal ? "¥ " + formatMoney(craftDisplayTotal) : '<span class="price-missing">无该批量定价</span>';

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
    const selectedRope = ROPE_CONFIG.find(item => item.id === ropeId);
    const ropeName = selectedRope ? selectedRope.name : "未选择吊绳";
    const ropePrice = result.ropePrice != null
      ? `<span class="rope-result-price">¥ ${formatMoney(result.ropePrice)}</span>`
      : '<span class="price-missing">无该批量定价</span>';
    els.resRopePrice.innerHTML = `
      <span class="rope-result-detail">
        <span class="rope-result-name">${escapeHtml(ropeName)}</span>
        <span class="rope-result-formula"> × ${result.tier}个 = </span>
        ${ropePrice}
      </span>
    `;
    // v8.0：邮费按三种情况展示明细
    if (result.shippingPrice != null) {
      if (result.shippingMode === "weight") {
        const shipRegion = SHIPPING_CONFIG.find(s => s.id === regionId);
        const coeff = shipRegion ? shipRegion.overTierCoeff : "";
        els.resShippingPrice.innerHTML = `重量 ${result.shippingWeightUsed}kg × ${coeff} = <span style="color:var(--brand);font-weight:600;">¥ ${formatMoney(result.shippingPrice)}</span>`;
      } else if (result.shippingMode === "manual") {
        els.resShippingPrice.innerHTML = `<span style="color:var(--brand);font-weight:600;">¥ ${formatMoney(result.shippingPrice)}</span>（手动输入）`;
      } else {
        els.resShippingPrice.innerHTML = "¥ " + formatMoney(result.shippingPrice);
      }
    } else {
      els.resShippingPrice.innerHTML = '<span class="price-missing">待输入</span>';
    }
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

  // 渲染报价卡片：直接系数模式显示每张纸独立系数的明细卡片，标准模式显示统一系数卡片
  if (calcMode === "direct") {
    // v6.8：明细卡片，每张纸单独乘系数后累加，工艺直接累加
    els.priceCards.innerHTML = result.pricesByLevel.map((item, idx) => {
      // v6.14：徽章只统计有直接系数的纸张，无直接系数纸张不参与系数范围
      const coeffs = item.paperDetails.filter(p => p.hasDirectCoeff).map(p => p.coefficient).filter(v => v != null);
      const badge = coeffs.length === 1
        ? `×${coeffs[0]}`
        : (coeffs.length > 1 ? `×${Math.min(...coeffs)}-${Math.max(...coeffs)}` : "");
      return `
      <div class="price-card direct-detail-card${idx === 0 ? " highlight" : ""}">
        <span class="coeff-badge">${badge}</span>
        <div class="level-name">${escapeHtml(item.levelName)}</div>
        <div class="direct-detail-list">
          ${item.paperDetails.map(p => `
            <div class="direct-detail-row">
              <span class="dd-name">${escapeHtml(p.paperName)}</span>
              <span class="dd-calc">${p.contributedPrice != null
                ? p.isBatchDirect
                  ? `¥${formatMoney(p.batchDirectPrice)}（批量直接价）`
                  : (p.hasDirectCoeff)
                  ? (p.craftTotal > 0
                      ? `(¥${formatMoney(p.originalUnitPrice)} + ¥${formatMoney(p.craftTotal)}) × ${p.coefficient} = ¥${formatMoney(p.contributedPrice)}`
                      : `¥${formatMoney(p.originalUnitPrice)} × ${p.coefficient} = ¥${formatMoney(p.contributedPrice)}`)
                  : (p.discount !== 1)
                  ? (p.craftTotal > 0
                      ? `¥${formatMoney(p.originalUnitPrice)} × ${p.discount}（折扣）+ ¥${formatMoney(p.craftTotal)} = ¥${formatMoney(p.contributedPrice)}`
                      : `¥${formatMoney(p.originalUnitPrice)} × ${p.discount}（折扣） = ¥${formatMoney(p.contributedPrice)}`)
                  : (p.craftTotal > 0)
                  ? `¥${formatMoney(p.originalUnitPrice)} + ¥${formatMoney(p.craftTotal)} = ¥${formatMoney(p.contributedPrice)}`
                  : `¥${formatMoney(p.contributedPrice)}`
                : '<span class="price-missing">缺价</span>'}</span>
              ${!p.hasDirectCoeff && !p.isBatchDirect ? '<span class="dd-tag">无直接系数</span>' : ''}
              ${p.isBatchDirect ? '<span class="dd-tag">批量直接价</span>' : ''}
            </div>
          `).join("")}
          ${item.batchDirectCraftTotal > 0 ? `
          <div class="direct-detail-row">
            <span class="dd-name">工艺（批量直接）</span>
            <span class="dd-calc">¥${formatMoney(item.batchDirectCraftTotal)}</span>
          </div>` : ''}
        </div>
        <div class="level-price">${item.costIncomplete
          ? '<span class="price-missing">部分缺价</span>'
          : formatMoney(item.price) + '<span class="unit">元</span>'}</div>
      </div>
    `;}).join("");
  } else {
    els.priceCards.innerHTML = result.pricesByLevel.map((item, idx) => `
      <div class="price-card${idx === 0 ? " highlight" : ""}">
        <span class="coeff-badge">×${item.coefficient}</span>
        <div class="level-name">${escapeHtml(item.levelName)}</div>
        <div class="level-price">${result.costIncomplete
          ? '<span class="price-missing">部分缺价</span>'
          : formatMoney(item.price) + '<span class="unit">元</span>'}</div>
      </div>
    `).join("");
  }

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
  if (els.saveQuoteBtn) els.saveQuoteBtn.disabled = false;

  // 临时系数输入栏：直接系数模式显示每纸独立临时直接系数，标准模式显示统一临时毛利系数
  if (calcMode === "direct") {
    if (els.customCoeffBar) { els.customCoeffBar.style.display = "none"; }
    if (els.customPriceCard) { els.customPriceCard.style.display = "none"; els.customPriceCard.innerHTML = ""; }
    if (els.tempCoeffBar) {
      els.tempCoeffBar.style.display = "flex";
      renderTempCoeffInputs();
      renderTempCoeffResults();
    }
  } else {
    if (els.tempCoeffBar) { els.tempCoeffBar.style.display = "none"; }
    if (els.tempCoeffResults) { els.tempCoeffResults.style.display = "none"; els.tempCoeffResults.innerHTML = ""; }
    if (els.customCoeffBar) {
      els.customCoeffBar.style.display = "flex";
      const labelEl = els.customCoeffBar.querySelector(".custom-coeff-label");
      if (labelEl) labelEl.textContent = "临时毛利系数：";
      renderCustomCoeffCard();
    }
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
  applyDefaultQuoteVisibility();

  // v8.0：渲染完成后，若邮费需输入，弹出对应对话框（固定档不弹）
  if (needShippingWeight) {
    openShippingWeightDialog(regionId, tier);
  } else if (needManualShipping) {
    openManualShippingDialog(regionId, tier);
  }
}

// 缓存上一次计算结果
let _lastResult = null;

// -------------------- v8.0 邮费输入对话框 --------------------
// 记录当前待确认的对话框上下文（供确认时写入缓存）
let _pendingShipping = null; // { type: "weight" | "manual", regionId, tier }

function openShippingWeightDialog(regionId, tier) {
  if (!els.shippingWeightDialog) return;
  const region = SHIPPING_CONFIG.find(s => s.id === regionId);
  const regionName = region ? region.name : "";
  _pendingShipping = { type: "weight", regionId, tier };
  if (els.shippingWeightDesc) {
    els.shippingWeightDesc.textContent = `地区「${regionName}」数量 ${tier} 张超过 10000，请填写吊牌总重量，邮费 = 重量 × 系数`;
  }
  if (els.shippingWeightInput) {
    els.shippingWeightInput.value = shippingWeightCache[regionId] != null ? shippingWeightCache[regionId] : "";
  }
  openShippingDialog(els.shippingWeightDialog);
  setTimeout(() => { if (els.shippingWeightInput) els.shippingWeightInput.focus(); }, 0);
}

function openManualShippingDialog(regionId, tier) {
  if (!els.manualShippingDialog) return;
  const region = SHIPPING_CONFIG.find(s => s.id === regionId);
  const regionName = region ? region.name : "";
  const manualKey = regionId + "_" + tier;
  _pendingShipping = { type: "manual", regionId, tier };
  if (els.manualShippingDesc) {
    els.manualShippingDesc.textContent = `地区「${regionName}」无 ${tier} 张批量定价，请手动填写邮费金额`;
  }
  if (els.manualShippingInput) {
    els.manualShippingInput.value = manualShippingCache[manualKey] != null ? manualShippingCache[manualKey] : "";
  }
  openShippingDialog(els.manualShippingDialog);
  setTimeout(() => { if (els.manualShippingInput) els.manualShippingInput.focus(); }, 0);
}

function openShippingDialog(dialog) {
  if (!dialog) return;
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
  } else {
    dialog.setAttribute("open", "");
  }
}

function closeShippingDialog(dialog) {
  if (!dialog) return;
  if (typeof dialog.close === "function") {
    dialog.close();
  } else {
    dialog.removeAttribute("open");
  }
  _pendingShipping = null;
}

function confirmShippingWeight() {
  if (!_pendingShipping || !els.shippingWeightInput) return;
  const w = parseFloat(els.shippingWeightInput.value);
  if (!isNaN(w) && w > 0) {
    shippingWeightCache[_pendingShipping.regionId] = w;
    closeShippingDialog(els.shippingWeightDialog);
    onCalculate();
  } else {
    closeShippingDialog(els.shippingWeightDialog);
  }
}

function confirmManualShipping() {
  if (!_pendingShipping || !els.manualShippingInput) return;
  const m = parseFloat(els.manualShippingInput.value);
  const manualKey = _pendingShipping.regionId + "_" + _pendingShipping.tier;
  if (!isNaN(m) && m > 0) {
    manualShippingCache[manualKey] = m;
    closeShippingDialog(els.manualShippingDialog);
    onCalculate();
  } else {
    delete manualShippingCache[manualKey];
    closeShippingDialog(els.manualShippingDialog);
  }
}

/**
 * 渲染临时毛利系数报价卡片
 */
/**
 * v8.1：读取临时毛利系数与邮费快速修改的当前输入值。
 */
function getOverrideValues() {
  const coeffRaw = els.customCoeffInput ? els.customCoeffInput.value.trim() : "";
  const coeff = parseFloat(coeffRaw);
  const hasCoeff = !!coeffRaw && !isNaN(coeff) && coeff >= 1;
  const shipRaw = els.shippingOverrideInput ? els.shippingOverrideInput.value.trim() : "";
  const newShipping = parseFloat(shipRaw);
  const hasShipOverride = !!shipRaw && !isNaN(newShipping) && newShipping >= 0;
  return { coeff, newShipping, hasCoeff, hasShipOverride };
}

function renderCustomCoeffCard() {
  if (!els.customPriceCard || !_lastResult) return;
  const { coeff, newShipping, hasCoeff, hasShipOverride } = getOverrideValues();
  if (!hasCoeff) {
    els.customPriceCard.style.display = "none";
    els.customPriceCard.innerHTML = "";
    return;
  }
  let price;
  let label;
  if (hasShipOverride) {
    // v8.1：同时填写了邮费快速修改 → 先用新邮费重算总成本，再乘临时系数
    const origShipping = _lastResult.shippingPrice || 0;
    const newCost = _lastResult.cost - origShipping + newShipping;
    price = newCost * coeff;
    label = `临时系数 ${coeff} · 新邮费 ¥${formatMoney(newShipping)}`;
  } else {
    price = _lastResult.cost * coeff;
    label = `临时系数 ${coeff}`;
  }
  els.customPriceCard.innerHTML = `
    <div class="price-card custom-coeff">
      <span class="coeff-badge">×${coeff}</span>
      <div class="level-name">${label}</div>
      <div class="level-price">${_lastResult.costIncomplete
        ? '<span class="price-missing">部分缺价</span>'
        : formatMoney(price) + '<span class="unit">元</span>'}</div>
    </div>
  `;
  els.customPriceCard.style.display = "flex";
}

/**
 * v6.8：渲染每纸独立临时直接系数输入框。
 * 有直接系数的纸张默认值 = 该纸当前档位普通客户系数；无直接系数的纸张默认值 = 1（按折扣价/原价计入）。
 */
function renderTempCoeffInputs() {
  if (!els.tempCoeffInputs || !_lastResult) return;
  const details = _lastResult.sheetDetails;
  const tier = _lastResult.tier;
  els.tempCoeffInputs.innerHTML = details.map((sd, i) => {
    const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === sd.paperId);
    const hasDirect = paper && paperHasDirectCoeff(paper);
    // v7.0：批量直接报价纸张不参与临时系数，固定显示批量直接价
    if (sd.isBatchDirect) {
      return `
        <div class="temp-coeff-item">
          <span class="temp-coeff-name">纸张${i + 1}（${escapeHtml(sd.paperName)}）</span>
          <input type="number" class="temp-coeff-input" data-sheet-idx="${i}" value="1" disabled />
          <span class="temp-coeff-tag">批量直接价</span>
        </div>
      `;
    }
    let def = "1";
    if (hasDirect) {
      const coeffs = getDirectCoeffsForTier(tier, paper);
      if (coeffs && coeffs[0]) def = String(coeffs[0].coefficient);
    }
    return `
      <div class="temp-coeff-item">
        <span class="temp-coeff-name">纸张${i + 1}（${escapeHtml(sd.paperName)}）</span>
        <input type="number" class="temp-coeff-input" data-sheet-idx="${i}" min="0.01" max="10" step="0.01" value="${def}" />
        ${!hasDirect ? '<span class="temp-coeff-tag">无直接系数</span>' : ''}
      </div>
    `;
  }).join("");
  els.tempCoeffInputs.querySelectorAll(".temp-coeff-input:not([disabled])").forEach(input => {
    input.addEventListener("input", renderTempCoeffResults);
  });
}

/**
 * v6.8：渲染每纸独立临时直接系数报价结果。
 * 有直接系数的纸张：原价（不打折）× 临时系数；
 * 无直接系数的纸张：有折扣打折扣（原价 × discount），无折扣则原价，再 × 临时系数（默认 1）。
 * 工艺费用直接累加。
 */
function renderTempCoeffResults() {
  if (!els.tempCoeffResults || !_lastResult) return;
  const details = _lastResult.sheetDetails;
  // v7.2：常规工艺费已并入各纸张（纸张+工艺）× 系数，仅批量直接报价工艺费单独累加
  const batchDirectCraftTotal = _lastResult.batchDirectCraftTotal || 0;
  const batchDirectTotal = _lastResult.batchDirectTotal || 0;
  const inputs = els.tempCoeffInputs.querySelectorAll(".temp-coeff-input");
  let total = batchDirectTotal + batchDirectCraftTotal;
  let incomplete = false;
  const rows = [];
  details.forEach((sd, i) => {
    const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID).find(p => p.id === sd.paperId);
    const hasDirect = paper && paperHasDirectCoeff(paper);
    // v7.0：批量直接报价纸张 → 固定价格，不乘临时系数
    if (sd.isBatchDirect) {
      if (sd.unitPrice == null) {
        incomplete = true;
        rows.push({
          name: sd.paperName,
          display: '<span class="price-missing">无该批量定价</span>',
          tag: "批量直接价"
        });
      } else {
        rows.push({
          name: sd.paperName,
          display: `¥${formatMoney(sd.unitPrice)}（批量直接价）`,
          tag: "批量直接价"
        });
      }
      return;
    }
    const raw = inputs[i] ? inputs[i].value.trim() : "";
    const coeff = parseFloat(raw);
    if (!raw || isNaN(coeff) || coeff < 0.01) {
      incomplete = true;
      rows.push({
        name: sd.paperName,
        display: '<span class="price-missing">无效系数</span>',
        tag: !hasDirect ? "无直接系数" : ""
      });
      return;
    }
    // 基础价：乘面积系数后的原价
    const base = sd.originalUnitPrice != null ? sd.originalUnitPrice : sd.unitPrice;
    // 无直接系数的纸张：有折扣打折扣，无折扣则原价
    const discount = hasDirect ? 1 : (paper ? (paper.discount || 1) : 1);
    // v7.2：工艺价先加到纸张价上再乘系数：(纸张价 + 工艺价) × 系数
    const craftOfSheet = sd.sheetCraftTotal || 0;
    const price = (base * discount + craftOfSheet) * coeff;
    total += price;
    // v6.14：无直接系数且系数为 1 时不显示冗余的 ×1
    let calcStr;
    if (!hasDirect && coeff === 1) {
      calcStr = discount !== 1
        ? (craftOfSheet > 0
            ? `¥${formatMoney(base)} × ${discount}（折扣）+ ¥${formatMoney(craftOfSheet)} = ¥${formatMoney(price)}`
            : `¥${formatMoney(base)} × ${discount}（折扣） = ¥${formatMoney(price)}`)
        : (craftOfSheet > 0
            ? `¥${formatMoney(base)} + ¥${formatMoney(craftOfSheet)} = ¥${formatMoney(price)}`
            : `¥${formatMoney(price)}`);
    } else if (!hasDirect) {
      calcStr = discount !== 1
        ? (craftOfSheet > 0
            ? `(¥${formatMoney(base)} × ${discount}（折扣）+ ¥${formatMoney(craftOfSheet)}) × ${coeff} = ¥${formatMoney(price)}`
            : `¥${formatMoney(base)} × ${discount}（折扣）× ${coeff} = ¥${formatMoney(price)}`)
        : (craftOfSheet > 0
            ? `(¥${formatMoney(base)} + ¥${formatMoney(craftOfSheet)}) × ${coeff} = ¥${formatMoney(price)}`
            : `¥${formatMoney(base)} × ${coeff} = ¥${formatMoney(price)}`);
    } else {
      calcStr = craftOfSheet > 0
        ? `(¥${formatMoney(base)} + ¥${formatMoney(craftOfSheet)}) × ${coeff} = ¥${formatMoney(price)}`
        : `¥${formatMoney(base)} × ${coeff} = ¥${formatMoney(price)}`;
    }
    rows.push({
      name: sd.paperName,
      display: calcStr,
      tag: !hasDirect ? "无直接系数" : ""
    });
  });
  els.tempCoeffResults.innerHTML = `
    <div class="price-card custom-coeff temp-result-card">
      <span class="coeff-badge">临时</span>
      <div class="level-name">临时报价结果</div>
      <div class="direct-detail-list">
        ${rows.map(r => `
          <div class="direct-detail-row">
            <span class="dd-name">${escapeHtml(r.name)}</span>
            <span class="dd-calc">${r.display}</span>
            ${r.tag ? `<span class="dd-tag">${r.tag}</span>` : ''}
          </div>
        `).join("")}
        ${batchDirectCraftTotal > 0 ? `
        <div class="direct-detail-row">
          <span class="dd-name">工艺（批量直接）</span>
          <span class="dd-calc">¥${formatMoney(batchDirectCraftTotal)}</span>
        </div>` : ''}
      </div>
      <div class="level-price">${incomplete
        ? '<span class="price-missing">部分缺价</span>'
        : formatMoney(total) + '<span class="unit">元</span>'}</div>
    </div>
  `;
  els.tempCoeffResults.style.display = "flex";
}

/**
 * 渲染邮费修改后的报价卡片（3 个默认等级，基于修改后的邮费重算）
 */
function renderShippingOverrideCards() {
  if (!els.shippingOverrideCards || !_lastResult) return;
  const { newShipping, hasCoeff, hasShipOverride } = getOverrideValues();
  // 没输入/无效值，或同时填写了临时毛利系数（v8.1 合并由临时系数卡片展示）时隐藏
  if (!hasShipOverride || hasCoeff) {
    els.shippingOverrideCards.style.display = "none";
    els.shippingOverrideLabel.style.display = "none";
    if (els.shippingOverrideCost) els.shippingOverrideCost.style.display = "none";
    els.shippingOverrideCards.innerHTML = "";
    return;
  }
  // 用修改后的邮费重算成本
  const origShipping = _lastResult.shippingPrice || 0;
  const newCost = _lastResult.cost - origShipping + newShipping;
  const costIncomplete = _lastResult.costIncomplete;
  // v8.3：显示修改后成本（高亮红字）
  if (els.shippingOverrideCost) {
    if (els.shippingOverrideCostValue) {
      els.shippingOverrideCostValue.textContent = costIncomplete
        ? "部分缺价"
        : "¥ " + formatMoney(newCost);
    }
    els.shippingOverrideCost.style.display = "flex";
  }
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
  if (els.saveQuoteBtn) els.saveQuoteBtn.disabled = true;
  if (els.customCoeffBar) els.customCoeffBar.style.display = "none";
  if (els.customPriceCard) { els.customPriceCard.style.display = "none"; els.customPriceCard.innerHTML = ""; }
  if (els.tempCoeffBar) els.tempCoeffBar.style.display = "none";
  if (els.tempCoeffResults) { els.tempCoeffResults.style.display = "none"; els.tempCoeffResults.innerHTML = ""; }
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
        if (v == null) return `<td class="price-cell price-editable" data-type="spec" data-code="${escapeHtml(s.code)}" data-tier="${t}" data-empty="1" title="点击填写价格"><span class="price-missing">无该批量定价</span></td>`;
        return `<td class="price-cell price-editable" data-type="spec" data-code="${escapeHtml(s.code)}" data-tier="${t}" title="点击编辑价格">¥ ${formatMoney(Number(v))}</td>`;
      }).join("")}
    </tr>
  `).join("");

  els.paperDiscount.textContent = `${paper.name} | ${paper.discount === 1 ? "无折扣" : (paper.discount * 10).toFixed(1) + "折"}`;
  // v6.7：直接系数档位表（单独框体表格显示，位于吊牌特殊工艺价格表下方）
  // 统一用 paperHasDirectCoeff 判断，确保三行完整且数量一致
  if (els.directCoeffTable && els.directCoeffWrap) {
    if (paperHasDirectCoeff(paper)) {
      const dc = paper.directCoeff;
      const dcTiers = dc.tiers.map(t => `${t}张`);
      const dcTheadRow = els.directCoeffTable.querySelector("thead tr");
      dcTheadRow.innerHTML = '<th>项目</th>' + dcTiers.map(t => `<th>${t}</th>`).join("");
      const dcTbody = els.directCoeffTable.querySelector("tbody");
      dcTbody.innerHTML = `
        <tr><td><strong>最高倍数</strong></td>${dc.max.map((v, i) => `<td class="price-cell price-editable" data-type="dc-max" data-index="${i}" title="点击编辑最高倍数">×${formatMoney(Number(v))}</td>`).join("")}</tr>
        <tr><td><strong>最低倍数</strong></td>${dc.min.map((v, i) => `<td class="price-cell price-editable" data-type="dc-min" data-index="${i}" title="点击编辑最低倍数">×${formatMoney(Number(v))}</td>`).join("")}</tr>
      `;
      els.directCoeffWrap.style.display = "";
    } else {
      els.directCoeffTable.querySelector("thead tr").innerHTML = '<th>项目</th>';
      els.directCoeffTable.querySelector("tbody").innerHTML =
        '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);">该纸张无直接系数（最高/最低倍数为空），按标准报价计算</td></tr>';
      els.directCoeffWrap.style.display = "";
    }
  }
  // v7.0：批量直接报价表（仅部分 Sheet 有配置，无配置时显示占位提示）
  // 板块：最大面积（出血后）+ 批量档位 + 批量价格；无价格档位留空占位（读取但不计算）
  if (els.batchDirectTable && els.batchDirectWrap) {
    const bd = paper.batchDirect;
    const hasBD = !!(bd && bd.maxArea > 0 && bd.prices && Object.keys(bd.prices).length > 0);
    const bdTheadRow = els.batchDirectTable.querySelector("thead tr");
    const bdTbody = els.batchDirectTable.querySelector("tbody");
    if (hasBD) {
      const bdTiers = Object.keys(bd.prices).map(Number).sort((a, b) => a - b);
      bdTheadRow.innerHTML = '<th>项目</th>' + bdTiers.map(t => `<th>${t} 张</th>`).join("");
      bdTbody.innerHTML = `
        <tr class="bd-maxarea"><td><strong>最大面积（出血后）</strong></td><td class="price-cell price-editable" data-type="bd-maxarea" title="点击编辑最大面积" colspan="${bdTiers.length}">${bd.maxArea} mm²</td></tr>
        <tr><td><strong>批量价格</strong></td>${bdTiers.map(t => {
          const v = bd.prices[t];
          if (v == null) return `<td class="price-cell price-editable" data-type="bd" data-tier="${t}" data-empty="1" title="点击填写价格"><span class="price-missing">无该批量报价</span></td>`;
          return `<td class="price-cell price-editable" data-type="bd" data-tier="${t}" title="点击编辑价格">¥ ${formatMoney(Number(v))}</td>`;
        }).join("")}</tr>
      `;
      els.batchDirectWrap.style.display = "";
    } else {
      bdTheadRow.innerHTML = '<th>项目</th>';
      bdTbody.innerHTML =
        '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);">该纸张无批量直接报价，按常规报价计算</td></tr>';
      els.batchDirectWrap.style.display = "";
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
            if (v == null) return `<td class="price-cell price-editable" data-type="craft" data-craft-id="${escapeHtml(c.id)}" data-tier="${t}" data-empty="1" title="点击填写价格"><span class="price-missing">无该批量定价</span></td>`;
            return `<td class="price-cell price-editable" data-type="craft" data-craft-id="${escapeHtml(c.id)}" data-tier="${t}" title="点击编辑价格">¥ ${formatMoney(Number(v))}</td>`;
          }).join("")}
        </tr>
      `).join("")
    : '<tr><td colspan="' + (tiers.length + 1) + '" style="text-align:center;color:var(--text-secondary);">该纸张暂无工艺配置</td></tr>';

  // v7.9：渲染吊绳与邮费价格表（全局共享，不随报价表/纸张切换）
  renderRopePriceTable();
  renderShippingPriceTable();
}

// -------------------- v7.9 吊绳/邮费价格表渲染（就地编辑） --------------------
function renderRopePriceTable() {
  if (!els.ropePriceTable) return;
  const tierKeys = ROPE_CONFIG.length ? Object.keys(ROPE_CONFIG[0].prices).map(Number).sort((a, b) => a - b) : [];
  const theadRow = els.ropePriceTable.querySelector("thead tr");
  theadRow.innerHTML = '<th>吊绳名称</th>' + tierKeys.map(t => `<th>${t} 张</th>`).join("");
  const tbody = els.ropePriceTable.querySelector("tbody");
  tbody.innerHTML = ROPE_CONFIG.map(r => `
    <tr>
      <td>${escapeHtml(r.name)}</td>
      ${tierKeys.map(t => {
        const v = r.prices[t];
        if (v == null) return `<td class="price-cell price-editable" data-type="rope" data-rope-id="${escapeHtml(r.id)}" data-tier="${t}" data-empty="1" title="点击填写价格"><span class="price-missing">无</span></td>`;
        return `<td class="price-cell price-editable" data-type="rope" data-rope-id="${escapeHtml(r.id)}" data-tier="${t}" title="点击编辑价格">¥ ${formatMoney(Number(v))}</td>`;
      }).join("")}
    </tr>
  `).join("");
}

function renderShippingPriceTable() {
  if (!els.shippingPriceTable) return;
  const tierKeys = SHIPPING_CONFIG.length ? Object.keys(SHIPPING_CONFIG[0].basePrices).map(Number).sort((a, b) => a - b) : [];
  const theadRow = els.shippingPriceTable.querySelector("thead tr");
  theadRow.innerHTML = '<th>地区名称</th>' + tierKeys.map(t => `<th>${t} 张</th>`).join("") + '<th>超量系数</th>';
  const tbody = els.shippingPriceTable.querySelector("tbody");
  tbody.innerHTML = SHIPPING_CONFIG.map(s => `
    <tr>
      <td>${escapeHtml(s.name)}</td>
      ${tierKeys.map(t => {
        const v = s.basePrices[t];
        if (v == null) return `<td class="price-cell price-editable" data-type="shipping" data-region-id="${escapeHtml(s.id)}" data-tier="${t}" data-empty="1" title="点击填写价格"><span class="price-missing">无</span></td>`;
        return `<td class="price-cell price-editable" data-type="shipping" data-region-id="${escapeHtml(s.id)}" data-tier="${t}" title="点击编辑价格">¥ ${formatMoney(Number(v))}</td>`;
      }).join("")}
      <td class="price-cell price-editable" data-type="shipping-coeff" data-region-id="${escapeHtml(s.id)}" title="点击编辑超量系数">${s.overTierCoeff}</td>
    </tr>
  `).join("");
}

// -------------------- v7.9 价格表就地编辑 --------------------
let _editingCell = null;

// 从单元格文本中提取首个数字（去掉 ¥、×、mm² 等符号）
function extractNumericFromCell(text) {
  const m = String(text || "").match(/-?\d+(?:\.\d+)?/);
  return m ? m[0] : "";
}

function startCellEdit(cell) {
  if (_editingCell) return; // 已有编辑进行中，忽略（blur 会负责提交）
  const type = cell.dataset.type;
  const isEmpty = cell.hasAttribute("data-empty");
  const raw = extractNumericFromCell(cell.textContent);
  const input = document.createElement("input");
  input.type = "number";
  input.step = "any";
  input.min = type === "bd-maxarea" ? "1" : "0";
  input.className = "price-cell-input";
  input.value = isEmpty ? "" : raw;
  input.dataset.type = type;
  // v7.9.1：保存原数值与显示文本，用于 confirm 对话框的差异提示与取消时的恢复
  input.dataset.originalValue = isEmpty ? "" : raw;
  input.dataset.originalDisplay = cell.textContent.trim();
  ["code", "tier", "craftId", "index", "ropeId", "regionId"].forEach(k => {
    if (cell.dataset[k] !== undefined) input.dataset[k] = cell.dataset[k];
  });
  cell.textContent = "";
  cell.appendChild(input);
  _editingCell = { cell, input, type };
  input.focus();
  if (!isEmpty) input.select();
  input.addEventListener("blur", commitCellEdit);
  input.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); input.blur(); }
    else if (e.key === "Escape") { cancelCellEdit(); }
  });
}

function applyPriceEdit(type, data, value) {
  // 全局类型：吊绳 / 邮费（不依赖当前纸张）
  if (type === "rope") {
    const rope = ROPE_CONFIG.find(r => r.id === data.ropeId);
    if (!rope) return false;
    rope.prices[data.tier] = value;
    saveToStorage("ropeConfig", ROPE_CONFIG);
    return true;
  }
  if (type === "shipping" || type === "shipping-coeff") {
    const region = SHIPPING_CONFIG.find(s => s.id === data.regionId);
    if (!region) return false;
    if (type === "shipping") region.basePrices[data.tier] = value;
    else region.overTierCoeff = value;
    saveToStorage("shippingConfig", SHIPPING_CONFIG);
    return true;
  }

  // 纸张相关类型
  const paper = getPapersByPriceList(CURRENT_PRICE_LIST_ID)[currentPaperIndex];
  if (!paper) return false;
  let changed = false;
  switch (type) {
    case "spec": {
      const spec = paper.specs.find(s => s.code === data.code);
      if (spec) { spec.prices[data.tier] = value; changed = true; }
      break;
    }
    case "craft": {
      const crafts = CRAFT_CONFIG[paper.id];
      const craft = crafts && crafts.find(c => c.id === data.craftId);
      if (craft) { craft.prices[data.tier] = value; changed = true; }
      break;
    }
    case "dc-max": {
      const idx = Number(data.index);
      if (paper.directCoeff && Array.isArray(paper.directCoeff.max)) {
        paper.directCoeff.max[idx] = value; changed = true;
      }
      break;
    }
    case "dc-min": {
      const idx = Number(data.index);
      if (paper.directCoeff && Array.isArray(paper.directCoeff.min)) {
        paper.directCoeff.min[idx] = value; changed = true;
      }
      break;
    }
    case "bd": {
      if (paper.batchDirect && paper.batchDirect.prices) {
        paper.batchDirect.prices[data.tier] = value; changed = true;
      }
      break;
    }
    case "bd-maxarea": {
      if (paper.batchDirect) { paper.batchDirect.maxArea = value; changed = true; }
      break;
    }
  }
  if (changed) {
    saveToStorage("paperConfig", PAPER_CONFIG);
    if (type === "craft") saveToStorage("craftConfig", CRAFT_CONFIG);
  }
  return changed;
}

function commitCellEdit() {
  if (!_editingCell) return;
  const { input, type } = _editingCell;
  _editingCell = null;
  const valueStr = input.value.trim();
  const originalStr = input.dataset.originalValue || "";
  const originalDisplay = input.dataset.originalDisplay || "";

  // v7.9.1：数值相等视为未改动（避免 "35.00" vs "35" 这种字符串差异误弹 confirm）
  const origNum = originalStr === "" ? null : Number(originalStr);
  const newNum = valueStr === "" ? null : Number(valueStr);
  if (origNum === newNum) {
    renderPriceTable();
    return;
  }

  // 清空（原值非空）→ 弹确认后设为 null（无该批量定价）
  if (valueStr === "") {
    const ok = window.confirm(`确认清空该价格？\n\n原值：${originalDisplay}\n新值：（无该批量定价）`);
    if (ok) {
      if (applyPriceEdit(type, input.dataset, null)) {
        renderPriceTable();
        if (type === "rope") rebuildRopeUI();
        onCalculate();
        showToast("价格已清空");
      } else renderPriceTable();
    } else {
      renderPriceTable();
      showToast("已取消修改");
    }
    return;
  }

  const value = Number(valueStr);
  if (isNaN(value) || value < 0) {
    showToast("请输入非负数字");
    renderPriceTable();
    return;
  }

  // 有改动 → 弹 confirm 显示原值与新值，确认才落盘
  const newDisplay = (type === "dc-max" || type === "dc-min")
    ? "×" + formatMoney(value)
    : (type === "bd-maxarea" ? value + " mm²" : "¥ " + formatMoney(value));
  const ok = window.confirm("确认修改该价格？\n\n原值：" + originalDisplay + "\n新值：" + newDisplay);
  if (ok) {
    if (applyPriceEdit(type, input.dataset, value)) {
      renderPriceTable();
      if (type === "rope") rebuildRopeUI();
      onCalculate();
      showToast("价格已更新");
    } else renderPriceTable();
  } else {
    renderPriceTable();
    showToast("已取消修改");
  }
}

function cancelCellEdit(silent) {
  if (!_editingCell) return;
  _editingCell = null;
  renderPriceTable();
  if (!silent) showToast("已取消编辑");
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
  // 报价表查询页下拉框 + 计算器页快捷下拉框联动渲染
  const renderOne = (sel) => {
    if (!sel) return;
    sel.innerHTML = PRICE_LISTS.map(pl => {
      const selected = pl.id === CURRENT_PRICE_LIST_ID ? " selected" : "";
      return `<option value="${escapeHtml(pl.id)}"${selected}>${escapeHtml(pl.name)}</option>`;
    }).join("");
    sel.value = CURRENT_PRICE_LIST_ID;
  };
  renderOne(els.priceListSelector);
  renderOne(els.quickPriceListSelector);
}

function onPriceListChange() {
  // 兼容两个下拉框触发：以发起事件的元素取值
  const active = document.activeElement;
  let newId = "";
  if (active && (active === els.priceListSelector || active === els.quickPriceListSelector)) {
    newId = active.value;
  } else {
    newId = (els.quickPriceListSelector && els.quickPriceListSelector.value) ||
            (els.priceListSelector && els.priceListSelector.value);
  }
  if (newId && newId !== CURRENT_PRICE_LIST_ID) {
    setCurrentPriceList(newId);
    currentPaperIndex = 0; // 重置纸张索引
    // 纸张按报价表隔离，需重渲染；吊绳/邮费全局共享，无需重渲染
    rebuildPaperUI();
    renderPriceTable();
    onCalculate();
    renderPriceListSelector(); // 同步两个下拉框选中态
    showToast(`已切换到「${getCurrentPriceList().name}」`);
  } else if (newId) {
    // 值未变化时也同步选中态（例如通过另一处已切换）
    renderPriceListSelector();
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
function syncCalcModeUI() {
  els.modeBtns.forEach(btn => {
    btn.classList.toggle("active", btn.dataset.mode === calcMode);
  });
  // 直接系数模式隐藏吊绳和邮费输入区域
  const ropeGroup = els.rope ? els.rope.closest(".form-group") : null;
  const regionGroup = els.region ? els.region.closest(".form-group") : null;
  if (calcMode === "direct") {
    if (ropeGroup) ropeGroup.style.display = "none";
    if (regionGroup) regionGroup.style.display = "none";
  } else {
    if (ropeGroup) ropeGroup.style.display = "";
    if (regionGroup) regionGroup.style.display = "";
  }
  applyDefaultQuoteVisibility();
}

function switchCalcMode(mode) {
  calcMode = mode === "standard" ? "standard" : "direct";
  saveToStorage("currentCalcMode", calcMode);
  syncCalcModeUI();
  // v6.14：重新渲染纸张卡片，刷新无直接系数提示（标准报价模式隐藏提示）
  renderSheets();
  onCalculate();
}

/**
 * 一键恢复所有配置为出厂默认（纸张 / 工艺 / 吊绳 / 客户等级）。
 * 仅清 4 个 key 的 localStorage；保留报价历史、本地快照、APP_PROFILE、个人偏好等。
 * 用途：换设备或被旧 localStorage 污染后，一键回到 DEFAULT 状态。
 */
function resetToDefaults() {
  const confirmMsg = "将清空以下本地修改并恢复出厂默认：\n\n• 报价表组（恢复为默认 1楼/3楼 报价表）\n• 纸张配置（44 张）\n• 工艺配置（含 烫金/UV/鸡眼/凹凸 等）\n• 吊绳配置\n• 邮费配置\n• 客户等级\n\n报价历史与本地快照不会被删除。\n\n确定继续？";
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
  const confirmMsg = "⚠️ 确定要恢复全局默认设置吗？\n\n将清除以下本地配置：\n• 报价表组（恢复为默认 1楼/3楼 报价表）\n• 纸张配置（44 张）\n• 工艺配置（含 烫金/UV/鸡眼/凹凸 等）\n• 吊绳配置\n• 邮费配置\n• 客户等级与毛利系数\n• 公司信息与个人偏好\n\n报价历史与本地快照不受影响。\n\n此操作不可撤销，恢复后页面将自动刷新。";
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
      // v8.2：恢复吊绳与邮费配置（此前遗漏，导致快照恢复不完整）
      if (item.data.ropeConfig) {
        ROPE_CONFIG = item.data.ropeConfig;
        saveToStorage("ropeConfig", ROPE_CONFIG);
      }
      if (item.data.shippingConfig) {
        SHIPPING_CONFIG = item.data.shippingConfig;
        saveToStorage("shippingConfig", SHIPPING_CONFIG);
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
  const history = loadFromStorage("history", []);
  return Array.isArray(history) ? history : [];
}

function saveHistory(list) {
  saveToStorage("history", list);
}

function collectCurrentQuoteInputs() {
  return {
    priceListId: CURRENT_PRICE_LIST_ID,
    mode: calcMode,
    sheetCount: sheetsState.length,
    tier: Number(els.tier?.value || 0),
    sheets: cloneQuoteData(sheetsState),
    ropeId: els.rope?.querySelector('input[name="rope"]:checked')?.value || "",
    regionId: els.region?.value || "",
    customCoefficient: els.customCoeffInput?.value.trim() || "",
    shippingOverride: els.shippingOverrideInput?.value.trim() || "",
    tempCoefficients: els.tempCoeffInputs
      ? Array.from(els.tempCoeffInputs.querySelectorAll(".temp-coeff-input")).map(input => input.value)
      : []
  };
}

function saveCurrentQuote() {
  if (!_lastResult) {
    showToast("请先完成有效报价，再保存记录");
    return;
  }

  const createdAt = new Date();
  const inputs = collectCurrentQuoteInputs();
  const selectedRope = ROPE_CONFIG.find(item => item.id === inputs.ropeId);
  const selectedRegion = SHIPPING_CONFIG.find(item => item.id === inputs.regionId);
  const priceList = getCurrentPriceList() || {};
  const record = buildQuoteHistoryRecord({
    id: generateId(),
    createdAt,
    customerName: els.quoteCustomerName?.value || "",
    note: els.quoteOrderNote?.value || "",
    mode: calcMode,
    priceList: { id: priceList.id || "", name: priceList.name || "" },
    inputs,
    result: _lastResult,
    ropeName: selectedRope?.name || "",
    regionName: selectedRegion?.name || ""
  });

  const list = getHistory();
  list.unshift(record);
  saveHistory(list);

  if (!getHistory().some(item => item.id === record.id)) {
    showToast("报价保存失败：浏览器本地空间可能已满");
    return;
  }

  renderHistory();
  showToast(`报价已保存：${record.title}`);
}

function getHistoryRecordCost(record) {
  const value = record?.snapshot?.cost ?? record?.cost;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function getHistoryRecordPrimaryPrice(record) {
  const levels = record?.snapshot?.pricesByLevel;
  const value = Array.isArray(levels) && levels.length ? levels[0]?.price : record?.price;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function getHistoryRecordTier(record) {
  const value = record?.inputs?.tier ?? record?.snapshot?.tier ?? record?.tier;
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}

function formatHistoryMoney(value, emptyText = "—") {
  return value == null || !Number.isFinite(Number(value)) ? emptyText : "¥ " + formatMoney(Number(value));
}

function formatHistoryDate(value) {
  const date = new Date(value);
  return isNaN(date.getTime()) ? "未知时间" : date.toLocaleString();
}

function getHistoryModeName(record) {
  return record?.mode === "direct" ? "直接系数" : "标准报价";
}

function canReloadHistoryRecord(record) {
  return !!(record?.inputs && Array.isArray(record.inputs.sheets) && record.inputs.sheets.length);
}

function renderHistory() {
  const list = getHistory();
  if (!els.historyTable || !els.historyEmpty) return;
  const tbody = els.historyTable.querySelector("tbody");
  els.historyEmpty.style.display = list.length ? "none" : "block";
  els.historyTable.style.display = list.length ? "table" : "none";
  tbody.innerHTML = list.map(h => {
    const title = h.title || h.customerName || h.recordNo || "历史报价";
    const recordNo = h.recordNo || h.id || "旧版记录";
    const tier = getHistoryRecordTier(h);
    const priceListName = h.priceList?.name || h.paper || "旧版报价表";
    const note = String(h.note || "").trim();
    const reloadable = canReloadHistoryRecord(h);
    return `
      <tr>
        <td>
          <span class="history-record-title">${escapeHtml(title)}</span>
          <span class="history-record-sub">${escapeHtml(recordNo)}</span>
          ${note ? `<span class="history-record-note" title="${escapeHtml(note)}">${escapeHtml(note)}</span>` : ""}
        </td>
        <td>${escapeHtml(formatHistoryDate(h.createdAt))}</td>
        <td>
          <span class="history-mode-badge">${escapeHtml(getHistoryModeName(h))}</span>
          <span class="history-record-sub">${escapeHtml(priceListName)}</span>
        </td>
        <td>${tier == null ? "—" : tier + " 张"}</td>
        <td>${formatHistoryMoney(getHistoryRecordCost(h))}</td>
        <td>${formatHistoryMoney(getHistoryRecordPrimaryPrice(h))}</td>
        <td>
          <div class="history-action-group">
            <button class="btn sm" data-action="view-history" data-id="${escapeHtml(h.id)}">查看</button>
            <button class="btn secondary sm" data-action="load-history" data-id="${escapeHtml(h.id)}"${reloadable ? "" : " disabled title=\"旧版记录没有完整参数\""}>载入</button>
            <button class="btn danger sm" data-action="delete-history" data-id="${escapeHtml(h.id)}">删除</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll("[data-action='view-history']").forEach(btn => {
    btn.addEventListener("click", () => showHistoryDetail(btn.dataset.id));
  });

  tbody.querySelectorAll("[data-action='load-history']:not([disabled])").forEach(btn => {
    btn.addEventListener("click", () => loadHistoryParameters(btn.dataset.id));
  });

  tbody.querySelectorAll("[data-action='delete-history']").forEach(btn => {
    btn.addEventListener("click", () => {
      const item = getHistory().find(x => x.id === btn.dataset.id);
      if (!item) return;
      const title = item.title || item.customerName || item.recordNo || "该报价";
      if (!confirm(`确定删除报价记录「${title}」？`)) return;
      const next = getHistory().filter(x => x.id !== btn.dataset.id);
      saveHistory(next);
      renderHistory();
      showToast("已删除记录");
    });
  });
}

function showHistoryDetail(recordId) {
  const record = getHistory().find(item => item.id === recordId);
  if (!record || !els.historyDetailDialog || !els.historyDetailContent) return;
  activeHistoryRecordId = record.id;

  const title = record.title || record.customerName || record.recordNo || "历史报价";
  const recordNo = record.recordNo || record.id || "旧版记录";
  const snapshot = record.snapshot || {};
  const sheets = Array.isArray(snapshot.sheetDetails) ? snapshot.sheetDetails : [];
  const levels = Array.isArray(snapshot.pricesByLevel) ? snapshot.pricesByLevel : [];
  const tier = getHistoryRecordTier(record);
  const cost = getHistoryRecordCost(record);
  const legacyPaper = record.paper || "";
  const legacySize = record.size || "";

  if (els.historyDetailTitle) els.historyDetailTitle.textContent = title;

  const sheetHtml = sheets.length
    ? sheets.map((sheet, index) => {
        const crafts = Array.isArray(sheet.crafts) && sheet.crafts.length
          ? "工艺：" + sheet.crafts.map(craft => `${craft.name}${craft.price == null ? "（缺价）" : " ¥" + formatMoney(Number(craft.price))}`).join("、")
          : "无附加工艺";
        return `
          <div class="history-sheet-item">
            <div class="history-sheet-main">
              <strong>纸张 ${index + 1} · ${escapeHtml(sheet.paperName || "未命名纸张")}</strong>
              <span>${escapeHtml(String(sheet.width || "—"))} × ${escapeHtml(String(sheet.length || "—"))} mm · ${escapeHtml(sheet.code || "无代码")} · ${escapeHtml(crafts)}</span>
            </div>
            <span class="history-sheet-price">${formatHistoryMoney(sheet.unitPrice, "缺价")}</span>
          </div>
        `;
      }).join("")
    : `<div class="history-sheet-item"><div class="history-sheet-main"><strong>${escapeHtml(legacyPaper || "旧版记录")}</strong><span>${escapeHtml(legacySize || "没有保存详细纸张参数")}</span></div></div>`;

  const levelHtml = levels.length
    ? levels.map(level => `
        <div class="history-level-item">
          <strong>${escapeHtml(level.levelName || "客户报价")}</strong>
          <span class="history-level-price">${formatHistoryMoney(level.price, "部分缺价")}</span>
        </div>
      `).join("")
    : `<div class="history-level-item"><strong>建议价</strong><span class="history-level-price">${formatHistoryMoney(record.price)}</span></div>`;

  const costRows = [
    ["纸张原价合计", snapshot.paperOriginalTotal],
    ["纸张折后价合计", snapshot.paperTotal],
    ["工艺费用", (Number(snapshot.craftTotal) || 0) + (Number(snapshot.batchDirectCraftTotal) || 0)],
    [snapshot.ropeName ? `吊绳 · ${snapshot.ropeName}` : "吊绳费用", snapshot.ropePrice],
    [snapshot.regionName ? `邮费 · ${snapshot.regionName}` : "邮费", snapshot.shippingPrice]
  ].filter(([, value]) => value != null);

  els.historyDetailContent.innerHTML = `
    <div class="history-detail-meta">
      <div class="history-detail-meta-item"><span>记录编号</span><strong>${escapeHtml(recordNo)}</strong></div>
      <div class="history-detail-meta-item"><span>保存时间</span><strong>${escapeHtml(formatHistoryDate(record.createdAt))}</strong></div>
      <div class="history-detail-meta-item"><span>计算模式</span><strong>${escapeHtml(getHistoryModeName(record))}</strong></div>
      <div class="history-detail-meta-item"><span>报价表</span><strong>${escapeHtml(record.priceList?.name || legacyPaper || "旧版报价表")}</strong></div>
      <div class="history-detail-meta-item"><span>批量档位</span><strong>${tier == null ? "—" : tier + " 张"}</strong></div>
      <div class="history-detail-meta-item"><span>客户名称</span><strong>${escapeHtml(record.customerName || "未填写")}</strong></div>
    </div>
    <section class="history-detail-section">
      <h3>订单备注</h3>
      <div class="history-note-content">${escapeHtml(record.note || "未填写备注")}</div>
    </section>
    <section class="history-detail-section">
      <h3>纸张与工艺</h3>
      <div class="history-sheet-list">${sheetHtml}</div>
    </section>
    <section class="history-detail-section">
      <h3>成本快照</h3>
      ${costRows.map(([label, value]) => `<div class="history-cost-row"><span>${escapeHtml(label)}</span><strong>${formatHistoryMoney(Number(value))}</strong></div>`).join("")}
      <div class="history-cost-row total"><span>成本合计</span><strong>${snapshot.costIncomplete ? "部分缺价" : formatHistoryMoney(cost)}</strong></div>
    </section>
    <section class="history-detail-section">
      <h3>保存时客户报价</h3>
      <div class="history-level-list">${levelHtml}</div>
    </section>
  `;

  if (els.historyDetailLoadBtn) {
    els.historyDetailLoadBtn.disabled = !canReloadHistoryRecord(record);
    els.historyDetailLoadBtn.title = canReloadHistoryRecord(record) ? "按当前价格重新计算" : "旧版记录没有完整参数";
  }
  if (typeof els.historyDetailDialog.showModal === "function") {
    els.historyDetailDialog.showModal();
  } else {
    els.historyDetailDialog.setAttribute("open", "");
  }
}

function closeHistoryDetail() {
  activeHistoryRecordId = null;
  if (!els.historyDetailDialog) return;
  if (typeof els.historyDetailDialog.close === "function" && els.historyDetailDialog.open) {
    els.historyDetailDialog.close();
  } else {
    els.historyDetailDialog.removeAttribute("open");
  }
}

function loadHistoryParameters(recordId) {
  const record = getHistory().find(item => item.id === recordId);
  if (!record || !canReloadHistoryRecord(record)) {
    showToast("该记录没有可重新载入的完整参数");
    return;
  }

  const inputs = record.inputs;
  const priceListId = inputs.priceListId || record.priceList?.id;
  if (!PRICE_LISTS.some(item => item.id === priceListId)) {
    showToast("原报价表已不存在，只能查看保存时快照");
    return;
  }

  const papers = getPapersByPriceList(priceListId);
  const paperIds = new Set(papers.map(paper => paper.id));
  const missingPaper = inputs.sheets.find(sheet => !paperIds.has(sheet.paperId));
  if (missingPaper) {
    showToast("原报价中的纸张已不存在，只能查看保存时快照");
    return;
  }

  const missingCraft = inputs.sheets.some(sheet => {
    const availableIds = new Set((CRAFT_CONFIG[sheet.paperId] || []).map(craft => craft.id));
    return (sheet.craftIds || []).some(craftId => !availableIds.has(craftId));
  });
  if (missingCraft) {
    showToast("原报价中的工艺已不存在，只能查看保存时快照");
    return;
  }

  const availableTiers = new Set();
  papers.forEach(paper => paper.specs.forEach(spec => Object.keys(spec.prices).forEach(tier => availableTiers.add(Number(tier)))));
  if (!availableTiers.has(Number(inputs.tier))) {
    showToast("原报价档位已不存在，只能查看保存时快照");
    return;
  }

  if (inputs.mode === "standard") {
    if (!ROPE_CONFIG.some(item => item.id === inputs.ropeId) || !SHIPPING_CONFIG.some(item => item.id === inputs.regionId)) {
      showToast("原吊绳或收货地区已不存在，只能查看保存时快照");
      return;
    }
  }

  setCurrentPriceList(priceListId);
  renderPriceListSelector();
  currentPaperIndex = Math.max(0, papers.findIndex(paper => paper.id === inputs.sheets[0].paperId));
  calcMode = inputs.mode === "standard" ? "standard" : "direct";
  saveToStorage("currentCalcMode", calcMode);
  syncCalcModeUI();

  const restoredSheets = cloneQuoteData(inputs.sheets).map(sheet => ({
    paperId: sheet.paperId,
    craftIds: Array.isArray(sheet.craftIds) ? sheet.craftIds.slice() : [],
    width: String(sheet.width || ""),
    length: String(sheet.length || ""),
    sizeType: sheet.sizeType || "single",
    manualCode: sheet.manualCode || null
  }));
  if (els.sheetCount) els.sheetCount.value = String(restoredSheets.length);
  sheetsState = restoredSheets;
  renderSheets();
  if (els.tier) els.tier.value = String(inputs.tier);

  const ropeInput = els.rope
    ? Array.from(els.rope.querySelectorAll('input[name="rope"]')).find(input => input.value === inputs.ropeId)
    : null;
  if (ropeInput) ropeInput.checked = true;
  if (els.region && inputs.regionId) els.region.value = inputs.regionId;
  if (els.customCoeffInput) els.customCoeffInput.value = inputs.customCoefficient || "";
  if (els.shippingOverrideInput) els.shippingOverrideInput.value = inputs.shippingOverride || "";

  onCalculate();

  if (calcMode === "direct" && Array.isArray(inputs.tempCoefficients) && els.tempCoeffInputs) {
    const coeffInputs = els.tempCoeffInputs.querySelectorAll(".temp-coeff-input");
    inputs.tempCoefficients.forEach((value, index) => {
      if (coeffInputs[index] && !coeffInputs[index].disabled) coeffInputs[index].value = value;
    });
    renderTempCoeffResults();
  }

  if (els.quoteCustomerName) els.quoteCustomerName.value = record.customerName || "";
  if (els.quoteOrderNote) els.quoteOrderNote.value = record.note || "";
  renderPriceTable();
  closeHistoryDetail();
  switchPage("calculator");
  document.getElementById("page-calculator")?.scrollIntoView({ behavior: "smooth", block: "start" });
  showToast("已载入历史参数，并按当前价格重新计算");
}

function openQuoteHistoryPanel() {
  switchPage("profile");
  switchTab("history");
  document.getElementById("panel-history")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function clearHistory() {
  if (!confirm("确定清空所有报价历史？")) return;
  saveHistory([]);
  renderHistory();
  closeHistoryDetail();
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
    version: "8.1.0",
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
    version: "8.1.0",
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
      // v6.6：全量导入同样执行数据校验（与本地备份恢复一致），防止损坏/恶意 JSON 导致崩溃
      validateImportedData(data);
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
  // v7.0：批量直接报价三行（与报价表表格格式一致：批量直接报价最大面积 / 批量直接报价档位 / 批量直接报价价格）
  // 有批量直接报价 → 填写实际最大面积/档位/价格；无批量直接报价 → 用纸张价格档位占位，最大面积与价格留空
  function batchDirectRows(paper, fallbackTiers) {
    const bd = paper.batchDirect;
    const hasBD = bd && bd.maxArea > 0 && bd.prices && Object.keys(bd.prices).length > 0;
    if (hasBD) {
      const bdTiers = Object.keys(bd.prices).map(Number).sort((a, b) => a - b);
      return [
        ["批量直接报价最大面积", bd.maxArea],
        ["批量直接报价档位", ...bdTiers],
        ["批量直接报价价格", ...bdTiers.map(t => (bd.prices[t] == null ? "" : bd.prices[t]))]
      ];
    }
    const tiers = fallbackTiers && fallbackTiers.length ? fallbackTiers : [];
    return [
      ["批量直接报价最大面积", ""],
      ["批量直接报价档位", ...tiers],
      ["批量直接报价价格", ...tiers.map(() => "")]
    ];
  }
  return [
    ["所属小组", GROUP_NAME_MAP[getCurrentPriceList().groupId] || GROUP_META.name],
    ["总报价表", getCurrentPriceList().name],
    ["报价表全称", paper.name],
    ["简称", paper.shortName],
    ["折扣系数", paper.discount],
    ...directCoeffRows(paper, tierKeys),
    ...batchDirectRows(paper, tierKeys),
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
    rows.push(["所属小组", GROUP_NAME_MAP[getCurrentPriceList().groupId] || GROUP_META.name]);
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
    // v7.0：批量直接报价三行（有批量直接报价填实际值，无则用价格档位占位，最大面积与价格留空）
    const bd = paper.batchDirect;
    const hasBD = bd && bd.maxArea > 0 && bd.prices && Object.keys(bd.prices).length > 0;
    const bdTiers = hasBD ? Object.keys(bd.prices).map(Number).sort((a, b) => a - b) : specTierKeys;
    rows.push(["批量直接报价最大面积", hasBD ? bd.maxArea : ""]);
    rows.push(["批量直接报价档位", ...bdTiers]);
    rows.push(["批量直接报价价格", ...(hasBD ? bdTiers.map(t => (bd.prices[t] == null ? "" : bd.prices[t])) : bdTiers.map(() => ""))]);
    rows.push(["备注", "直接系数档位/最高倍数/最低倍数三行：档位为批量张数，最高倍数→普通客户，最低倍数→大客户，中间等级自动等差插值。无直接系数的纸张按标准报价（乘折扣系数）计算。批量直接报价三行：最大面积（出血后）+ 档位 + 价格，面积在最大面积内时直接按批量价格报价（不打折、不乘系数），工艺费用直接叠加；无批量直接报价的纸张留空即可。"]);
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

    // v7.0：解析批量直接报价三行（批量直接报价最大面积 / 批量直接报价档位 / 批量直接报价价格）
    // 三行均填有值且档位与价格数量一致 → { maxArea, prices }；仅占位或缺失 → null（按常规报价计算）
    let batchDirect = null;
    {
      const bdRows = { maxArea: null, tiers: null, prices: null };
      for (let i = 0; i < rows.length; i++) {
        const label = String(rows[i] && rows[i][0] || "").trim();
        if (label === "批量直接报价最大面积") bdRows.maxArea = rows[i] && rows[i][1];
        else if (label === "批量直接报价档位") bdRows.tiers = rows[i];
        else if (label === "批量直接报价价格") bdRows.prices = rows[i];
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
      const maxAreaRaw = bdRows.maxArea;
      const bdMaxArea = maxAreaRaw === "" || maxAreaRaw == null ? 0 : Number(maxAreaRaw);
      const bdTiers = toNumArr(bdRows.tiers).filter(n => n > 0 && Number.isInteger(n));
      const bdPrices = toNumArr(bdRows.prices);
      // 最大面积有效 + 档位与价格数量一致 → 有效批量直接报价
      if (bdMaxArea > 0 && bdTiers.length && bdPrices.length === bdTiers.length) {
        const prices = {};
        bdTiers.forEach((t, i) => { prices[t] = bdPrices[i]; });
        batchDirect = { maxArea: bdMaxArea, prices };
      }
      // 仅占位（无最大面积/价格）或无任何行 → batchDirect 保持 null
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
    if (isNaN(discount) || discount <= 0 || discount > 10) {
      errors.push(`「${sheetName}」折扣系数无效（应为 0~10 之间），已使用默认值 1`);
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
    const rawHeader = [];
    for (let c = 2; c < headerRow.length; c++) {
      const val = headerRow[c];
      const valid = val !== "" && !isNaN(Number(val)) && Number(val) > 0;
      if (valid) {
        tierKeys.push(String(Number(val)));
        rawHeader.push(String(Number(val)));
      } else {
        rawHeader.push(null);
      }
    }
    // v6.6：检测档位表头中间空列（误删列/合并单元格残留），避免价格静默错位
    // 规则：第一个有效档位之后出现空列，且其后仍有有效档位 → 中间空列，报错提示
    const firstValidIdx = rawHeader.findIndex(v => v !== null);
    let midGap = false;
    if (firstValidIdx !== -1) {
      for (let i = firstValidIdx; i < rawHeader.length; i++) {
        if (rawHeader[i] === null && rawHeader.slice(i + 1).some(v => v !== null)) {
          midGap = true;
          break;
        }
      }
    }
    if (midGap) {
      errors.push(`「${sheetName}」档位表头中间存在空列（第 ${firstValidIdx + 3} 列附近），价格可能错位，请检查后重新导入`);
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

    // v6.6：始终生成唯一 ID，不再复用默认简称的 paper id。
    // 原因：CRAFT_CONFIG 以 paper.id 为键且跨报价表共享，若复用默认 ID（如 paper2），
    // 新导入报价表的工艺会覆盖原报价表的工艺（静默丢失）。唯一 ID 实现各报价表工艺完全隔离。
    const defaultPaper = DEFAULT_PAPER_CONFIG.find(p => p.shortName === shortName);
    let id;
    let counter = usedIds.size + 1;
    do {
      id = "paper_import_" + counter;
      counter++;
    } while (usedIds.has(id));
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
      discount: isNaN(discount) || discount <= 0 || discount > 10 ? 1 : discount,
      // 直接系数：Sheet 专属配置，只读取报价表表格三行（直接系数档位/最高倍数/最低倍数）。
      // 表格未填有效直接系数时，匹配默认简称则继承默认配置，否则为 null（按标准报价计算）
      directCoeff: directCoeff || (defaultPaper && defaultPaper.directCoeff
        ? JSON.parse(JSON.stringify(defaultPaper.directCoeff))
        : null),
      // v7.0：批量直接报价：Sheet 专属配置，只读取报价表表格三行（最大面积/档位/价格）。
      // 表格未填有效批量直接报价时，匹配默认简称则继承默认配置，否则为 null（按常规报价计算）
      batchDirect: batchDirect || (defaultPaper && defaultPaper.batchDirect
        ? JSON.parse(JSON.stringify(defaultPaper.batchDirect))
        : null),
      specs
    });
    if (crafts.length) {
      craftsByPaper[id] = crafts;
    } else if (defaultPaper && DEFAULT_CRAFT_CONFIG[defaultPaper.id] && DEFAULT_CRAFT_CONFIG[defaultPaper.id].length) {
      // v6.6：Excel 无工艺区时，从默认配置复制同名简称纸张的工艺到新唯一 ID 下。
      // 保持"导入后工艺可用"的既有行为，同时因 ID 唯一而不会覆盖其他报价表的工艺。
      craftsByPaper[id] = DEFAULT_CRAFT_CONFIG[defaultPaper.id].map((c, ci) => ({
        id: `craft_${id}_${ci + 1}`,
        name: c.name,
        prices: { ...c.prices }
      }));
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

function bindRopeEvents() {
  if (!els.rope) return;
  els.rope.querySelectorAll('input[name="rope"]').forEach(radio => {
    radio.addEventListener("change", onCalculate);
  });
}

function rebuildRopeUI() {
  // 重新渲染吊绳选择区（沿用与初始化一致的默认选择逻辑），并重新绑定实时计算事件
  els.rope.innerHTML = renderRopeRadios();
  bindRopeEvents();
}

// -------------------- 邮费 Excel 导入 / 导出 / 模板 --------------------

function setShippingExcelStatus(html, isError) {
  if (!els.shippingExcelImportStatus) return;
  els.shippingExcelImportStatus.innerHTML = html;
  els.shippingExcelImportStatus.style.color = isError ? "var(--danger)" : "var(--text-secondary)";
}

function shippingToSheetRows() {
  const tierKeys = SHIPPING_CONFIG.length ? Object.keys(SHIPPING_CONFIG[0].basePrices) : [];
  const headerRow = ["地区名称", ...tierKeys.map(String), "大于10000"];
  const dataRows = SHIPPING_CONFIG.map(region => [
    region.name,
    ...tierKeys.map(t => region.basePrices[t] ?? ""),
    region.overTierCoeff
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
  const tierKeys = sample ? Object.keys(sample.basePrices) : ["500", "1000", "2000", "5000", "10000"];
  const rows = [
    ["地区名称", ...tierKeys.map(String), "大于10000"],
    ["广东省内", ...tierKeys.map(t => sample ? sample.basePrices[t] : 0), sample ? sample.overTierCoeff : 1],
    ["江浙沪", ...tierKeys.map(() => 0), 1.2],
    ["其他省份", ...tierKeys.map(() => 0), 1.3]
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
  // 解析档位列（第2列起到倒数第1列），最后1列是"大于10000"系数
  const tierKeys = [];
  for (let c = 1; c < headerRow.length - 1; c++) {
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

    // 解析最后一列：大于10000系数
    const coeffIdx = 1 + tierKeys.length;
    const coeff = row[coeffIdx] !== "" && row[coeffIdx] != null
      ? Number(row[coeffIdx]) : 1;
    if (isNaN(coeff)) errors.push(`「${name}」的"大于10000"系数无效，已按 1 处理`);

    regions.push({
      id: "region" + (regions.length + 1),
      name,
      basePrices,
      overTierCoeff: isNaN(coeff) ? 1 : coeff
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
      // v8.2：导入后持久化到 localStorage，刷新不丢失
      saveToStorage("shippingConfig", SHIPPING_CONFIG);
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
  bindRopeEvents();

  // 纸张数量变化时重渲染纸张卡片再计算
  if (els.sheetCount) {
    els.sheetCount.addEventListener("input", () => {
      renderSheets();
      onCalculate();
    });
  }

  if (els.searchInput) els.searchInput.addEventListener("input", debounce(renderPriceTable, 200));

  // v7.9：价格表就地编辑（事件委托，点击价格单元格进入编辑）
  [els.priceTable, els.craftTable, els.directCoeffTable, els.batchDirectTable, els.ropePriceTable, els.shippingPriceTable].forEach(table => {
    if (!table) return;
    table.addEventListener("click", e => {
      const cell = e.target.closest(".price-editable");
      if (!cell) return;
      startCellEdit(cell);
    });
  });

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
  if (els.quickPriceListSelector) els.quickPriceListSelector.addEventListener("change", onPriceListChange);
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
  if (els.defaultQuoteToggle) {
    els.defaultQuoteToggle.addEventListener("click", toggleDefaultQuoteVisibility);
  }

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
  if (els.clearHistoryBtn) els.clearHistoryBtn.addEventListener("click", clearHistory);
  if (els.saveQuoteBtn) els.saveQuoteBtn.addEventListener("click", saveCurrentQuote);
  if (els.openQuoteHistoryBtn) els.openQuoteHistoryBtn.addEventListener("click", openQuoteHistoryPanel);
  if (els.historyDetailLoadBtn) {
    els.historyDetailLoadBtn.addEventListener("click", () => {
      if (activeHistoryRecordId) loadHistoryParameters(activeHistoryRecordId);
    });
  }
  [els.historyDetailCloseBtn, els.historyDetailCancelBtn].forEach(btn => {
    if (btn) btn.addEventListener("click", closeHistoryDetail);
  });
  if (els.historyDetailDialog) {
    els.historyDetailDialog.addEventListener("click", event => {
      if (event.target === els.historyDetailDialog) closeHistoryDetail();
    });
    els.historyDetailDialog.addEventListener("close", () => {
      activeHistoryRecordId = null;
    });
  }

  // v8.0：邮费输入对话框事件
  if (els.shippingWeightConfirm) {
    els.shippingWeightConfirm.addEventListener("click", confirmShippingWeight);
  }
  if (els.manualShippingConfirm) {
    els.manualShippingConfirm.addEventListener("click", confirmManualShipping);
  }
  if (els.shippingWeightInput) {
    els.shippingWeightInput.addEventListener("keydown", e => { if (e.key === "Enter") confirmShippingWeight(); });
  }
  if (els.manualShippingInput) {
    els.manualShippingInput.addEventListener("keydown", e => { if (e.key === "Enter") confirmManualShipping(); });
  }
  document.querySelectorAll("[data-close-dialog]").forEach(btn => {
    btn.addEventListener("click", () => {
      const dialog = document.getElementById(btn.getAttribute("data-close-dialog"));
      if (dialog) closeShippingDialog(dialog);
    });
  });
  [els.shippingWeightDialog, els.manualShippingDialog].forEach(dialog => {
    if (dialog) dialog.addEventListener("close", () => { _pendingShipping = null; });
  });

  // 点击页面其他区域关闭纸张材质下拉
  document.addEventListener("click", () => closeAllPaperDropdowns());

  // v7.4：功能占位卡片事件绑定（移除内联 onclick，支持 CSP 收紧）
  document.querySelectorAll(".feature-card[data-toast]").forEach(card => {
    card.addEventListener("click", () => {
      const msg = card.getAttribute("data-toast") || "功能开发中";
      showToast(msg);
    });
  });

  // 临时毛利系数：输入时实时渲染报价卡片（不触发 onCalculate）
  // v8.1：两个输入联动，任一变化都同时刷新临时系数卡片与邮费修改卡片
  if (els.customCoeffInput) {
    els.customCoeffInput.addEventListener("input", () => {
      renderCustomCoeffCard();
      renderShippingOverrideCards();
    });
  }

  // 邮费快速修改：输入时实时渲染修改后报价卡片
  if (els.shippingOverrideInput) {
    els.shippingOverrideInput.addEventListener("input", () => {
      renderCustomCoeffCard();
      renderShippingOverrideCards();
    });
  }
  // 清除邮费修改
  if (els.shippingOverrideClear) {
    els.shippingOverrideClear.addEventListener("click", () => {
      if (els.shippingOverrideInput) els.shippingOverrideInput.value = "";
      renderCustomCoeffCard();
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
    syncCalcModeUI();
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
