// ============================================================
// KOKALabel报价系统 v5.0 - 主程序（计算 + 渲染 + 交互 + 初始化）
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
 * 根据有效面积向上匹配尺寸规格。
 * 面积超过 10000 时返回错误提示；007 与 008 合并显示为 007/008。
 */
function matchSpec(paper, area) {
  if (area > 10000) {
    return { error: true, message: "请与上级联系" };
  }
  const candidates = paper.specs
    .filter(s => s.maxArea >= area)
    .sort((a, b) => a.maxArea - b.maxArea);
  if (!candidates.length) {
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
  const { sheetCount, tier, sheets, ropeId, regionId } = inputs;

  // 校验基础参数
  if (!sheetCount || !tier || !sheets || !ropeId || !regionId) {
    return null;
  }
  if (sheetCount <= 0 || tier <= 0) {
    return null;
  }
  if (!Array.isArray(sheets) || sheets.length !== sheetCount) {
    return null;
  }

  const rope = ROPE_CONFIG.find(r => r.id === ropeId);
  const region = SHIPPING_CONFIG.find(s => s.id === regionId);
  if (!rope || !region) return null;

  // 逐张纸计算：每张纸有独立的宽、长、尺寸类型
  let paperTotal = 0;
  let paperOriginalTotal = 0; // 折扣前纸张价合计
  let craftTotal = 0;
  const warnings = [];
  const sheetDetails = [];
  // 是否有任一纸张/工艺在所选档位没有定价（用于显示"无该批量定价"占位）
  let hasMissingTier = false;

  for (const sheet of sheets) {
    const { width, length, sizeType, paperId, craftIds } = sheet;
    if (!width || !length || width <= 0 || length <= 0) return null;

    const paper = PAPER_CONFIG.find(p => p.id === paperId);
    if (!paper) return null;

    // 单张含出血面积（单张尺寸 / 展开尺寸目前均按输入长宽直接计算）
    const singleArea = calcBleedArea(length, width);

    const spec = matchSpec(paper, singleArea);
    if (spec && spec.error) {
      return { error: spec.message };
    }

    // 纸张折后价：精确匹配档位，无值则占位
    const paperUnitPrice = hasExactTier(spec.prices, tier)
      ? Number(spec.prices[tier]) * paper.discount
      : null;
    const paperOriginalPrice = hasExactTier(spec.prices, tier)
      ? Number(spec.prices[tier])
      : null;
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
      paperName: paper.shortName || paper.name,
      width,
      length,
      sizeType,
      area: singleArea,
      code: spec.code,
      unitPrice: paperUnitPrice,
      originalUnitPrice: paperOriginalPrice,
      missing: paperUnitPrice == null,
      crafts: sheetCraftDetails
    });
  }

  // 吊绳费用
  const ropePrice = hasExactTier(rope.prices, tier)
    ? Number(rope.prices[tier])
    : null;
  if (ropePrice == null) {
    hasMissingTier = true;
    warnings.push(`吊绳「${rope.name}」无 ${tier} 张批量定价`);
  }

  // 邮费（以第一张纸的面积判断小面积折扣）
  let shippingPrice = hasExactTier(region.basePrices, tier)
    ? Number(region.basePrices[tier])
    : null;
  if (shippingPrice != null && sheetDetails.length && sheetDetails[0].area < region.smallAreaThreshold) {
    shippingPrice *= region.discount;
  } else if (shippingPrice == null) {
    hasMissingTier = true;
    warnings.push(`地区「${region.name}」无 ${tier} 张批量定价`);
  }

  // 成本合计（仅累加有定价的部分；缺价时显示为"部分缺价"）
  const costKnown = [paperTotal, craftTotal, ropePrice, shippingPrice].every(v => v != null);
  const cost = costKnown
    ? (paperTotal + craftTotal + ropePrice + shippingPrice)
    : (paperTotal + craftTotal + (ropePrice || 0) + (shippingPrice || 0)); // 仍把已知的相加，但标记不完整
  const costIncomplete = !costKnown;

  // 三个客户等级的建议报价
  const pricesByLevel = CUSTOMER_LEVELS.map(level => ({
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
  tableMeta: document.getElementById("tableMeta"),
  prevPaper: document.getElementById("prevPaper"),
  nextPaper: document.getElementById("nextPaper"),
  paperPageInfo: document.getElementById("paperPageInfo"),
  paperSelector: document.getElementById("paperSelector"),
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
  clearHistoryBtn: document.getElementById("clearHistoryBtn")
};

let currentPaperIndex = 2;
let toastTimer = null;
let sheetsState = [];

// -------------------- 初始化下拉选项 --------------------
// 根据 APP_PROFILE.defaultRope 生成吊绳单选项 HTML（顶层作用域，供 initOptions 和 rebuildRopeUI 共用）
function renderRopeRadios() {
  const fallback = "rope1"; // 默认「普通吊绳」
  const desired = APP_PROFILE.defaultRope || fallback;
  const hasDesired = ROPE_CONFIG.some(r => r.id === desired);
  const checkedId = hasDesired ? desired : (ROPE_CONFIG.find(r => r.id === fallback) ? fallback : (ROPE_CONFIG[0] && ROPE_CONFIG[0].id));
  return ROPE_CONFIG.map(r => `
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
    els.region.innerHTML = '<option value="">请选择地区</option>' +
      SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
    if (SHIPPING_CONFIG.length > 0) {
      els.region.value = SHIPPING_CONFIG[0].id;
    }
  }
}

/**
 * 更新批量档位下拉选项。
 * @param {boolean} isInit 是否为初始化，若是则默认选中第一项。
 */
function updateTierOptions(isInit) {
  // 批量档位取所有1号报价表所有规格的价格档位并集，避免切换纸张时档位缺失
  const tierSet = new Set();
  for (const paper of PAPER_CONFIG) {
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
  const newSheets = [];

  for (let i = 0; i < count; i++) {
    const old = sheetsState[i];
    newSheets.push({
      paperId: old && old.paperId ? old.paperId : "paper3",
      craftIds: old && old.craftIds ? old.craftIds.slice() : [],
      width: old && old.width ? old.width : "55",
      length: old && old.length ? old.length : "30",
      sizeType: old && old.sizeType ? old.sizeType : "single"
    });
  }
  sheetsState = newSheets;

  els.sheetList.innerHTML = "";
  sheetsState.forEach((sheet, index) => {
    const card = document.createElement("div");
    card.className = "sheet-card";

    const currentPaper = PAPER_CONFIG.find(p => p.id === sheet.paperId) || PAPER_CONFIG[0];
    // Sheet 数量 > 10 时显示编号前缀，方便快速定位
    const showPaperIndex = PAPER_CONFIG.length > 10;
    const paperOptions = PAPER_CONFIG.map((p, pIdx) => `
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

    card.innerHTML = `
      <div class="sheet-title">纸张 ${index + 1}</div>
      <div class="form-group">
        <label>纸张材质 <span class="hint">点击展开，选择后自动收起</span></label>
        <div class="paper-dropdown" data-sheet="${index}">
          <div class="paper-trigger">
            <div>
              <span class="paper-trigger-text">${escapeHtml(currentPaper.shortName || currentPaper.name)}</span>
              <span class="paper-trigger-desc">${escapeHtml(currentPaper.name)}</span>
            </div>
            <span class="paper-trigger-arrow"></span>
          </div>
          <div class="paper-options">
            ${paperOptions}
          </div>
        </div>
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
 * - 下方空间不足时向上翻转（flip-up）
 * - 右侧溢出时左移（shift-left）
 * - 弹窗宽度超出容器时切换为全屏固定定位（full-width）
 */
function adjustDropdownPosition(dropdown) {
  const options = dropdown.querySelector(".paper-options");
  if (!options) return;

  // 清除上一次的定位类
  options.classList.remove("flip-up", "shift-left", "full-width");

  const trigger = dropdown.querySelector(".paper-trigger");
  const triggerRect = trigger.getBoundingClientRect();
  const optionsRect = options.getBoundingClientRect();
  const vw = window.innerWidth;
  const vh = window.innerHeight;

  // 1. 垂直方向：判断下方空间是否充足
  const spaceBelow = vh - triggerRect.bottom;
  const spaceAbove = triggerRect.top;
  const optionsHeight = optionsRect.height;

  if (spaceBelow < optionsHeight + 10 && spaceAbove > spaceBelow) {
    options.classList.add("flip-up");
  }

  // 2. 水平方向：判断右侧是否溢出
  if (optionsRect.right > vw - 8) {
    options.classList.add("shift-left");
  }

  // 3. 超小屏：弹窗宽度超出视口时使用固定全宽定位
  if (triggerRect.width < vw * 0.6 && vw <= 480) {
    options.classList.add("full-width");
    // 重新计算垂直位置
    const newRect = options.getBoundingClientRect();
    if (newRect.bottom > vh - 8) {
      options.classList.add("flip-up");
    }
  }
}

function closeAllPaperDropdowns() {
  document.querySelectorAll(".paper-dropdown.open").forEach(dd => {
    dd.classList.remove("open");
    const opt = dd.querySelector(".paper-options");
    if (opt) opt.classList.remove("flip-up", "shift-left", "full-width");
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
  // 切换纸张时清空工艺选择
  sheetsState[index].craftIds = [];
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
  onCalculate();
}

// -------------------- 实时计算 --------------------
function onCalculate() {
  const sheetCount = parseInt(els.sheetCount.value, 10);
  const tier = parseInt(els.tier.value, 10);
  const ropeId = els.rope.querySelector('input[name="rope"]:checked')?.value || "";
  const regionId = els.region.value;

  if (!sheetCount || !tier || !ropeId || !regionId || sheetCount <= 0 || tier <= 0) {
    clearResult();
    return;
  }

  // 收集每张纸的纸材、工艺、尺寸
  const sheets = sheetsState.map(s => ({
    paperId: s.paperId,
    craftIds: s.craftIds.slice(),
    width: parseFloat(s.width),
    length: parseFloat(s.length),
    sizeType: s.sizeType
  }));

  // 只要有任何一张纸缺少有效尺寸就清空结果
  if (sheets.some(s => !s.width || !s.length || s.width <= 0 || s.length <= 0)) {
    clearResult();
    return;
  }

  const result = calculate({ sheetCount, tier, sheets, ropeId, regionId });

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
    tbody.innerHTML = result.sheetDetails.map(s => `
      <tr>
        <td>${escapeHtml(s.paperName)}</td>
        <td>${s.width} × ${s.length}${s.sizeType === "spread" ? "（展开）" : ""}</td>
        <td>${s.area} mm²</td>
        <td>${escapeHtml(s.code)}</td>
        <td>${s.missing ? '<span class="price-missing">无该批量定价</span>' : '¥ ' + formatMoney(s.unitPrice)}</td>
      </tr>
    `).join("");
  }

  els.resTier.textContent = result.tier + " 张";
  // 纸张价合计（折扣前）/ 纸张折后价合计：多纸张时显示 "¥价格1 + ¥价格2 = ¥合计"，合计蓝字
  function renderPaperTotal(getPrice) {
    const details = result.sheetDetails;
    if (details.length === 1) {
      const p = getPrice(details[0]);
      return p != null ? "¥ " + formatMoney(p) : '<span class="price-missing">无该批量定价</span>';
    }
    // 多纸张：逐张价格 + = 合计
    const parts = details.map(s => {
      const p = getPrice(s);
      return p != null ? "¥ " + formatMoney(p) : '<span class="price-missing">缺价</span>';
    });
    const total = getPrice(details[0]) != null && details.every(s => getPrice(s) != null);
    const totalStr = total
      ? '<span style="color:var(--brand);font-weight:600;">¥ ' + formatMoney(details.reduce((sum, s) => sum + getPrice(s), 0)) + '</span>'
      : '<span class="price-missing">部分缺价</span>';
    return parts.join(' + ') + ' = ' + totalStr;
  }
  els.resPaperOriginalPrice.innerHTML = renderPaperTotal(s => s.originalUnitPrice);
  els.resPaperPrice.innerHTML = renderPaperTotal(s => s.unitPrice);
  els.resCraftPrice.innerHTML = result.craftTotal ? "¥ " + formatMoney(result.craftTotal) : '<span class="price-missing">无该批量定价</span>';
  els.resRopePrice.innerHTML = result.ropePrice != null ? "¥ " + formatMoney(result.ropePrice) : '<span class="price-missing">无该批量定价</span>';
  els.resShippingPrice.innerHTML = result.shippingPrice != null ? "¥ " + formatMoney(result.shippingPrice) : '<span class="price-missing">无该批量定价</span>';
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

  // 渲染三个客户等级报价卡片
  els.priceCards.innerHTML = result.pricesByLevel.map((item, idx) => `
    <div class="price-card${idx === 0 ? " highlight" : ""}">
      <div class="level-name">${escapeHtml(item.levelName)}</div>
      <div class="level-price">${result.costIncomplete
        ? '<span class="price-missing">部分缺价</span>'
        : formatMoney(item.price) + '<span class="unit">元</span>'}</div>
    </div>
  `).join("");

  // 渲染批量档位快速切换按钮
  if (els.tierQuickSwitch && els.tierQuickBtns) {
    const tierSet = new Set();
    for (const paper of PAPER_CONFIG) {
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

  // -------------------- 临时毛利系数 & 邮费快速修改 --------------------
  // 缓存本次计算结果供临时系数/邮费修改使用
  _lastResult = result;

  // 显示临时毛利系数输入栏
  if (els.customCoeffBar) {
    els.customCoeffBar.style.display = "flex";
    renderCustomCoeffCard();
  }

  // 显示邮费快速修改栏
  if (els.shippingOverrideBar) {
    const hasShipping = result.shippingPrice != null;
    els.shippingOverrideBar.style.display = "flex";
    // 如果邮费输入框为空且当前有邮费，预填原始邮费
    if (hasShipping && els.shippingOverrideInput && !els.shippingOverrideInput.value) {
      // 不自动填充，让用户主动输入
    }
    renderShippingOverrideCards();
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
  els.resRopePrice.textContent = "-";
  els.resShippingPrice.textContent = "-";
  els.resCost.textContent = "-";
  if (els.resWarnings) {
    els.resWarnings.innerHTML = "";
    els.resWarnings.style.display = "none";
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

// -------------------- 1号报价表渲染 --------------------
function renderPriceTable() {
  const paper = PAPER_CONFIG[currentPaperIndex];
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
  els.tableMeta.textContent = `行数：${filtered.length} / ${paper.specs.length}`;
  els.paperPageInfo.textContent = `第 ${currentPaperIndex + 1} / ${PAPER_CONFIG.length} 张`;
  els.prevPaper.disabled = currentPaperIndex === 0;
  els.nextPaper.disabled = currentPaperIndex === PAPER_CONFIG.length - 1;
  els.paperNotes.textContent = `备注：当前展示「${paper.name}」1号报价表。007 与 008 合并为 007/008；匹配面积超过 10000 mm² 时请与上级联系。`;

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
  const showIndex = PAPER_CONFIG.length > 10;
  els.paperSelector.innerHTML = PAPER_CONFIG.map((p, i) => {
    const label = showIndex
      ? `${i + 1}. ${escapeHtml(p.shortName || p.name)}`
      : escapeHtml(p.shortName || p.name);
    return `<option value="${i}"${i === currentPaperIndex ? " selected" : ""}>${label}</option>`;
  }).join("");
}

// -------------------- 翻页 --------------------
function prevPaper() {
  if (currentPaperIndex > 0) {
    currentPaperIndex--;
    renderPriceTable();
  }
}

function nextPaper() {
  if (currentPaperIndex < PAPER_CONFIG.length - 1) {
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

/**
 * 一键恢复所有配置为出厂默认（纸张 / 工艺 / 吊绳 / 客户等级）。
 * 仅清 4 个 key 的 localStorage；保留报价历史、本地快照、APP_PROFILE、个人偏好等。
 * 用途：换设备或被旧 localStorage 污染后，一键回到 DEFAULT 状态。
 */
function resetToDefaults() {
  const confirmMsg = "将清空以下本地修改并恢复出厂默认：\n\n• 纸张配置（9 张1号报价表）\n• 工艺配置（含 烫金/UV/鸡眼/凹凸 等）\n• 吊绳配置\n• 客户等级\n\n报价历史与本地快照不会被删除。\n\n确定继续？";
  if (!confirm(confirmMsg)) return;

  const keysToReset = ["paperConfig", "craftConfig", "ropeConfig", "customerLevels"];
  keysToReset.forEach(k => {
    try { localStorage.removeItem("tagPricing_" + k); } catch (e) { /* 忽略 */ }
  });

  // 重新从 DEFAULT 派生（深拷贝，避免后续修改污染源对象）
  PAPER_CONFIG = DEFAULT_PAPER_CONFIG.map(p => ({
    ...p,
    specs: p.specs.map(s => ({ ...s, prices: { ...s.prices } }))
  }));
  CRAFT_CONFIG = {};
  Object.keys(DEFAULT_CRAFT_CONFIG).forEach(k => {
    CRAFT_CONFIG[k] = DEFAULT_CRAFT_CONFIG[k].map(c => ({ ...c, prices: { ...c.prices } }));
  });
  ROPE_CONFIG = DEFAULT_ROPE_CONFIG.map(r => ({ ...r, prices: { ...r.prices } }));
  CUSTOMER_LEVELS = DEFAULT_CUSTOMER_LEVELS.map(l => ({ ...l }));

  rebuildPaperUI();
  renderLevelSettings();
  // 重新填充下拉框默认值
  updateDefaultPaperOptions && updateDefaultPaperOptions();
  updateDefaultRopeOptions && updateDefaultRopeOptions();
  onCalculate();
  showToast("已恢复默认配置（含工艺 单面/双面 后缀）");
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
  const firstPaper = PAPER_CONFIG[0];
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
  const opts = ['<option value="">跟随默认（1号报价表-3）</option>'].concat(
    PAPER_CONFIG.map(p => `<option value="${escapeHtml(p.id)}">${escapeHtml(p.shortName || p.name)}</option>`)
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
      if (item.data.customerLevels) {
        CUSTOMER_LEVELS = item.data.customerLevels;
        saveToStorage("customerLevels", CUSTOMER_LEVELS);
      }
      if (item.data.appProfile) {
        APP_PROFILE = item.data.appProfile;
        saveToStorage("appProfile", APP_PROFILE);
      }
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
    version: "4.3",
    kind: "local-backup",
    exportAt: new Date().toISOString(),
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
        rebuildPaperUI();
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
        if (els.region) {
          els.region.innerHTML = '<option value="">请选择地区</option>' +
            SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
          if (SHIPPING_CONFIG.length > 0) els.region.value = SHIPPING_CONFIG[0].id;
        }
      }
      if (data.snapshots) saveSnapshots(data.snapshots);
      if (data.history) saveHistory(data.history);
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
    version: "4.3",
    exportAt: new Date().toISOString(),
    customerLevels: CUSTOMER_LEVELS,
    appProfile: APP_PROFILE,
    paperConfig: PAPER_CONFIG,
    ropeConfig: ROPE_CONFIG,
    craftConfig: CRAFT_CONFIG,
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
        rebuildPaperUI();
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
        if (els.region) {
          els.region.innerHTML = '<option value="">请选择地区</option>' +
            SHIPPING_CONFIG.map(s => `<option value="${escapeHtml(s.id)}">${escapeHtml(s.name)}</option>`).join("");
          if (SHIPPING_CONFIG.length > 0) els.region.value = SHIPPING_CONFIG[0].id;
        }
      }
      if (data.snapshots) saveSnapshots(data.snapshots);
      if (data.history) saveHistory(data.history);
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
  return [
    ["所属小组", GROUP_META.name],
    ["总报价表", PRICE_LIST_META.name],
    ["1号报价表全称", paper.name],
    ["简称", paper.shortName],
    ["折扣系数", paper.discount],
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
  for (const paper of PAPER_CONFIG) {
    const ws = XLSX.utils.aoa_to_sheet(paperToSheetRows(paper));
    XLSX.utils.book_append_sheet(wb, ws, paper.shortName || paper.name);
  }
  XLSX.writeFile(wb, "KOKALabel1号报价表_" + formatDateFile() + ".xlsx");
  showToast("1号报价表 Excel 已导出");
}

function downloadPaperTemplate() {
  // P1.4: 确保 SheetJS 已加载
  if (typeof XLSX === "undefined") { loadSheetJS().then(() => downloadPaperTemplate()).catch(() => showToast("Excel 库加载失败，请检查网络")); return; }
  const wb = XLSX.utils.book_new();

  // === 模板 = 源 Excel 的 9 张表 1:1 还原 ===
  // 来源：用户提供的「KOKALabel1号报价表.xlsx」真实表头与工艺名（与导入端 100% 一致）。
  // 9 张表顺序：350铜版纸 / 400铜版纸 / 702铜版纸 / 606米白卡 / 500白卡纸 / 海成专用 / 700布纹纸 / 600牛皮纸 / 40C棉麻布
  const TEMPLATE_ROWS = [
    // 1) 350铜版纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-1：350克 A级铜版纸 双面过哑胶（厚度0.38mm）"],
        ["简称", "350铜版纸"],
        ["折扣系数", 1],
        ["备注", "本表暂无附加工艺"]
      ],
      codeHeader: ["代码", "最大含出血面积", "500", "1000", "2000", "2500", "5000", "10000", "20000", "50000"],
      codeRows: [
        ["002", 1999, 20, 25, 45, 50, 95, 190, 350, 800],
        ["003", 3164, 25, 30, 55, 60, 105, 200, 400, 1000],
        ["004", 3999, 30, 35, 65, 70, 120, 240, 480, 1150],
        ["005", 4944, 35, 40, 75, 85, 150, 270, 530, 1300],
        ["055", 5684, 40, 45, 85, 90, 160, 300, 580, 1400],
        ["006", 6264, 50, 55, 95, 100, 170, 340, 660, 1600],
        ["007", 7999, 60, 65, 120, 130, 220, 420, 820, 1900],
        ["008", 7999, 60, 65, 120, 130, 220, 420, 820, 1900],
        ["009", 9999, 80, 90, 160, 170, 290, 560, 1100, 2500],
        ["100", 15999, 110, 130, 230, 250, 420, 820, 1600, 3700]
      ],
      craftHeader: null,
      craftRows: []
    },
    // 2) 400铜版纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-2：400克 A级铜版纸 双面过哑胶（厚度0.45mm）"],
        ["简称", "400铜版纸"],
        ["折扣系数", 0.91],
        ["备注", "工艺名带（单面）/（双面）后缀可区分；空值留空"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000", "20000", "30000", "50000", "100000"],
      codeRows: [
        ["002", 1999, 25, 45, 75, 85, 120, 210, 400, "", 900, ""],
        ["003", 3164, 35, 55, 90, 105, 140, 240, 460, "", 1100, ""],
        ["004", 3999, 40, 70, 105, 130, 170, 285, 550, "", 1300, ""],
        ["005", 4944, 50, 85, 130, 160, 210, 360, 700, "", 1700, ""],
        ["055", 5684, 60, 95, 145, 180, 240, 410, 800, "", 1900, ""],
        ["006", 6264, 65, 105, 160, 200, 260, 440, 880, "", 2050, ""],
        ["007", 7999, 80, 130, 195, 245, 320, 540, 1080, "", 2500, ""],
        ["008", 7999, 80, 130, 195, 245, 320, 540, 1080, "", 2500, ""],
        ["009", 9999, 100, 165, 245, 305, 400, 680, 1380, "", 3150, ""],
        ["100", 15999, 145, 240, 360, 445, 580, 980, 2000, "", 4550, ""]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000", "20000", "30000", "50000", "100000"],
      craftRows: [
        ["烫金（单面）", 90, 110, 130, 140, 160, 200, 280, 350, 450, 800],
        ["烫金（双面）", 110, 130, 150, 160, 200, 240, 380, 500, 800, 1400],
        ["UV", 100, 100, 130, 140, 160, 180, 240, 300, 400, 600],
        ["鸡眼", 40, 60, 90, 120, 150, 300, 600, 800, 1300, 2500],
        ["凹凸", 110, 120, 130, 140, 150, 200, 280, 320, 400, 500]
      ]
    },
    // 3) 702铜版纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-3：702克 A级铜版纸 双面过哑胶（厚度0.85mm）"],
        ["简称", "702铜版纸"],
        ["折扣系数", 1],
        ["备注", "附加工艺中烫金分单/双面，请分别填写"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000", "20000", "50000"],
      codeRows: [
        ["002", 1999, 35, 50, 70, 80, 90, 170, "", ""],
        ["003", 3164, 45, 60, 85, 100, 110, 210, "", ""],
        ["004", 3999, 55, 75, 100, 120, 135, 250, "", ""],
        ["005", 4944, 70, 95, 130, 155, 170, 320, "", ""],
        ["055", 5684, 80, 105, 145, 175, 195, 360, "", ""],
        ["006", 6264, 90, 120, 165, 195, 220, 400, "", ""],
        ["007", 7999, 110, 150, 200, 240, 270, 490, "", ""],
        ["008", 7999, 110, 150, 200, 240, 270, 490, "", ""],
        ["009", 9999, 145, 195, 265, 315, 355, 640, "", ""],
        ["100", 15999, 210, 280, 380, 455, 510, 920, "", ""]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000", "20000", "50000"],
      craftRows: [
        ["无色压凹（单面）", 50, 60, 70, 80, 90, 170, "", ""],
        ["套字压凹（单面）", 80, 90, 100, 110, 120, 230, "", ""],
        ["烫金 小面积（单面）", 50, 60, 70, 80, 90, 170, "", ""],
        ["烫金 小面积（双面）", 70, 80, 90, 100, 120, 230, "", ""],
        ["击凸（单面）", 80, 90, 100, 110, 120, 230, "", ""]
      ]
    },
    // 4) 606米白卡
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-4：米白卡 606克（厚度0.8mm）"],
        ["简称", "606米白卡"],
        ["折扣系数", 1],
        ["备注", "工艺分（单面）/（双面），与代码区批量档对齐"]
      ],
      codeHeader: ["代码", "最大含出血面积", "500", "1000", "2000", "2500", "5000", "7500", "10000", "20000", "30000", "50000"],
      codeRows: [
        ["002", 1999, 20, 30, 40, "", 70, "", 130, "", "", ""],
        ["003", 3164, 25, 35, 50, "", 80, "", 150, "", "", ""],
        ["004", 3999, 30, 40, 60, "", 90, "", 170, "", "", ""],
        ["005", 4944, 35, 50, 70, "", 110, "", 200, "", "", ""],
        ["055", 5684, 40, 55, 80, "", 125, "", 230, "", "", ""],
        ["006", 6264, 45, 60, 85, "", 135, "", 250, "", "", ""],
        ["007", 7999, 55, 75, 105, "", 165, "", 300, "", "", ""],
        ["008", 7999, 55, 75, 105, "", 165, "", 300, "", "", ""],
        ["009", 9999, 70, 95, 135, "", 210, "", 380, "", "", ""],
        ["100", 15999, 100, 140, 195, "", 305, "", 555, "", "", ""]
      ],
      craftHeader: ["工艺名称", "500", "1000", "2000", "2500", "5000", "7500", "10000", "20000", "30000", "50000"],
      craftRows: [
        ["无色压凹（单面）", "", 30, 40, "", 70, "", 130, "", "", ""],
        ["套字压凹（单面）", "", 80, 90, "", 120, "", 230, "", "", ""],
        ["烫金 小面积（单面）", "", 50, 60, "", 90, "", 170, "", "", ""],
        ["烫金 小面积（双面）", "", 70, 80, "", 120, "", 230, "", "", ""],
        ["击凸（单面）", "", 80, 90, "", 120, "", 230, "", "", ""]
      ]
    },
    // 5) 500白卡纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-5：500克 白卡纸（厚度0.55mm）正面过哑胶"],
        ["简称", "500白卡纸"],
        ["折扣系数", 0.75],
        ["备注", "工艺分（单面）/（双面），按实际勾选填写"]
      ],
      codeHeader: ["代码", "最大含出血面积", "500", "1000", "2000", "3000", "4000", "5000", "10000", "20000", "30000", "50000"],
      codeRows: [
        ["002", 1999, 18, 25, 40, 50, 60, 110, 200, "", "", ""],
        ["003", 3164, 22, 30, 50, 60, 75, 130, 240, "", "", ""],
        ["004", 3999, 25, 35, 60, 75, 90, 155, 290, "", "", ""],
        ["005", 4944, 30, 45, 70, 90, 110, 195, 360, "", "", ""],
        ["055", 5684, 35, 50, 80, 100, 120, 215, 400, "", "", ""],
        ["006", 6264, 40, 55, 90, 110, 130, 230, 430, "", "", ""],
        ["007", 7999, 50, 70, 110, 135, 160, 290, 530, "", "", ""],
        ["008", 7999, 50, 70, 110, 135, 160, 290, 530, "", "", ""],
        ["009", 9999, 65, 90, 140, 170, 200, 360, 660, "", "", ""],
        ["100", 15999, 90, 130, 200, 245, 290, 520, 960, "", "", ""]
      ],
      craftHeader: ["工艺名称", "500", "1000", "2000", "2500", "5000", "7500", "10000", "20000", "30000", "50000"],
      craftRows: [
        ["无色压凹（单面）", "", 30, 40, "", 70, "", 130, "", "", ""],
        ["套字压凹（单面）", "", 80, 90, "", 120, "", 230, "", "", ""],
        ["烫金 小面积（单面）", "", 50, 60, "", 90, "", 170, "", "", ""],
        ["烫金 小面积（双面）", "", 70, 80, "", 120, "", 230, "", "", ""],
        ["击凸（单面）", "", 80, 90, "", 120, "", 230, "", "", ""]
      ]
    },
    // 6) 海成专用
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-6：400克米白卡 / 160克半透卡 / 200克合成纸（海成专用 / 撕不烂）"],
        ["简称", "海成专用"],
        ["折扣系数", 0.75],
        ["备注", "鸡眼工艺分（单面）/（对折两张），请按工艺分别报价"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      codeRows: [
        ["002", 1999, 30, 45, 65, 80, 95, 170, ""],
        ["003", 3164, 35, 55, 80, 95, 115, 210, ""],
        ["004", 3999, 45, 65, 95, 115, 135, 250, ""],
        ["005", 4944, 55, 85, 120, 145, 170, 310, ""],
        ["055", 5684, 65, 95, 135, 165, 195, 350, ""],
        ["006", 6264, 70, 105, 150, 180, 215, 385, ""],
        ["007", 7999, 85, 125, 175, 215, 255, 455, ""],
        ["008", 7999, 85, 125, 175, 215, 255, 455, ""],
        ["009", 9999, 110, 160, 225, 275, 325, 580, ""],
        ["100", 15999, 160, 230, 325, 395, 470, 835, ""]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      craftRows: [
        ["烫金 小面积（单面）", 50, 60, 70, 80, 90, 170, ""],
        ["烫金 小面积（双面）", 70, 80, 90, 100, 120, 230, ""],
        ["鸡眼（单面）", 40, 60, 90, 120, 150, 300, ""],
        ["鸡眼 （对折两张）", 60, 120, 180, 240, 300, 550, ""]
      ]
    },
    // 7) 700布纹纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-7：700克 A级布纹纸 双面过光油（厚度0.85mm）"],
        ["简称", "700布纹纸"],
        ["折扣系数", 1],
        ["备注", "烫金小面积后缀（单面）/（双面）必须保留"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      codeRows: [
        ["002", 1999, 30, 45, 65, 80, 95, 170, ""],
        ["003", 3164, 35, 55, 80, 95, 115, 210, ""],
        ["004", 3999, 45, 65, 95, 115, 135, 250, ""],
        ["005", 4944, 55, 85, 120, 145, 170, 310, ""],
        ["055", 5684, 65, 95, 135, 165, 195, 350, ""],
        ["006", 6264, 70, 105, 150, 180, 215, 385, ""],
        ["007", 7999, 85, 125, 175, 215, 255, 455, ""],
        ["008", 7999, 85, 125, 175, 215, 255, 455, ""],
        ["009", 9999, 110, 160, 225, 275, 325, 580, ""],
        ["100", 15999, 160, 230, 325, 395, 470, 835, ""]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      craftRows: [
        ["烫金小面积(单面）", 90, 110, 130, 140, 160, 200, ""],
        ["烫金小面积(双面）", 110, 130, 150, 160, 200, 240, ""],
        ["凹凸", 110, 120, 130, 140, 150, 200, ""],
        ["鸡眼", 40, 60, 90, 120, 150, 300, ""],
        ["无色压凹（单面）", 110, 120, 130, 140, 150, 200, ""],
        ["深凹烫金（单面）", 110, 120, 130, 140, 150, 200, ""]
      ]
    },
    // 8) 600牛皮纸
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-8：600克 A级牛皮纸（厚度0.80mm）"],
        ["简称", "600牛皮纸"],
        ["折扣系数", 0.75],
        ["备注", "烫金与丝印白均区分（单面）/（双面）"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      codeRows: [
        ["002", 1999, 25, 40, 55, 65, 80, 150, ""],
        ["003", 3164, 30, 50, 70, 80, 100, 180, ""],
        ["004", 3999, 35, 60, 85, 95, 120, 215, ""],
        ["005", 4944, 45, 75, 105, 120, 150, 270, ""],
        ["055", 5684, 50, 85, 120, 135, 170, 305, ""],
        ["006", 6264, 55, 95, 130, 150, 185, 335, ""],
        ["007", 7999, 70, 115, 155, 180, 220, 395, ""],
        ["008", 7999, 70, 115, 155, 180, 220, 395, ""],
        ["009", 9999, 90, 145, 200, 230, 280, 505, ""],
        ["100", 15999, 130, 210, 290, 335, 405, 725, ""]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000", "20000"],
      craftRows: [
        ["烫金 小面积（单面）", 50, 60, 70, 80, 90, 170, ""],
        ["烫金 小面积（双面）", 70, 80, 90, 100, 120, 230, ""],
        ["凹凸", 80, 90, 100, 110, 120, 230, ""],
        ["鸡眼", 40, 60, 90, 120, 150, 300, ""],
        ["丝印白（单面）", 30, 40, 50, 60, 70, 130, ""],
        ["丝印白（双面）", 50, 60, 70, 90, 90, 160, ""]
      ]
    },
    // 9) 40C棉麻布
    {
      meta: [
        ["所属小组", "1号小组"],
        ["总报价表", "1号报价表"],
        ["1号报价表全称", "1号报价表-9：棉麻布 40C 米白色"],
        ["简称", "40C棉麻布"],
        ["折扣系数", 0.75],
        ["备注", "工艺均为（单面），单列一栏；如无双面请留空对应行"]
      ],
      codeHeader: ["代码", "最大含出血面积", "1000", "2000", "3000", "4000", "5000", "10000"],
      codeRows: [
        ["002", 1999, 25, 40, 60, 75, 90, 160],
        ["003", 3164, 30, 50, 75, 90, 110, 200],
        ["004", 3999, 35, 60, 90, 110, 130, 240],
        ["005", 4944, 45, 75, 110, 135, 160, 290],
        ["055", 5684, 50, 85, 125, 155, 180, 330],
        ["006", 6264, 55, 95, 135, 170, 195, 360],
        ["007", 7999, 70, 115, 165, 205, 235, 430],
        ["008", 7999, 70, 115, 165, 205, 235, 430],
        ["009", 9999, 90, 145, 215, 260, 300, 545],
        ["100", 15999, 130, 210, 310, 380, 435, 790]
      ],
      craftHeader: ["工艺名称", "1000", "2000", "3000", "4000", "5000", "10000"],
      craftRows: [
        ["专色（单面）", 90, 120, 150, 180, 200, 300],
        ["黑白图案（单面）", 30, 50, 70, 90, 120, 200],
        ["红色（单面）", 30, 50, 70, 90, 120, 200]
      ]
    }
  ];

  // 将「代码 + 工艺 + 单/双面」注入到每张表，sheet 命名为简称
  TEMPLATE_ROWS.forEach((t) => {
    const rows = [];
    t.meta.forEach((m) => rows.push(m));
    rows.push([]);
    rows.push(t.codeHeader);
    t.codeRows.forEach((r) => rows.push(r));
    if (t.craftHeader) {
      rows.push([]);
      rows.push(t.craftHeader);
      t.craftRows.forEach((r) => rows.push(r));
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    // 从 meta 中提取简称作为 Sheet 名（跳过架构层级行）
    const shortNameMeta = t.meta.find(m => m[0] === "简称");
    XLSX.utils.book_append_sheet(wb, ws, (shortNameMeta && shortNameMeta[1]) || ("Sheet" + (TEMPLATE_ROWS.indexOf(t) + 1)));
  });

  XLSX.writeFile(wb, "KOKALabel1号报价表模板_" + formatDateFile() + ".xlsx");
  showToast("模板已下载（9张表1:1，含单/双面工艺）");
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
      if (label === "所属小组" || label === "总报价表") continue; // 架构层级信息，解析时跳过
      if (label === "1号报价表全称" || label === "报价表全称") name = String(val || "").trim();
      else if (label === "简称") shortName = String(val || "").trim();
      else if (label === "折扣系数") discountRaw = val;
    }
    const discount = discountRaw === "" || discountRaw == null ? 1 : Number(discountRaw);

    if (!name) {
      errors.push(`「${sheetName}」缺少1号报价表全称，已跳过`);
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
    const id = defaultPaper ? defaultPaper.id : ("paper" + (sheetIndex + 1));

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
      const paperNum = id.replace(/^paper/, "");
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
        crafts.push({ id: `craft${paperNum}_${crafts.length + 1}`, name: craftName, prices });
      }
    }

    papers.push({
      id,
      name,
      shortName,
      discount: isNaN(discount) || discount <= 0 ? 1 : discount,
      specs
    });
    if (crafts.length) {
      craftsByPaper[id] = crafts;
    }
  });

  if (!papers.length) {
    throw new Error(errors.join("；") || "没有可导入的1号报价表");
  }

  return { papers, crafts: craftsByPaper, errors };
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
      const { papers, crafts, errors } = parsePaperExcel(e.target.result);
      PAPER_CONFIG = papers;
      saveToStorage("paperConfig", PAPER_CONFIG);
      // 合并导入的工艺：保留未在 Excel 中出现的纸张的默认工艺
      CRAFT_CONFIG = { ...DEFAULT_CRAFT_CONFIG, ...crafts };
      saveToStorage("craftConfig", CRAFT_CONFIG);
      rebuildPaperUI();
      onCalculate();
      const craftCount = Object.values(crafts).reduce((sum, arr) => sum + arr.length, 0);
      const successMsg = `成功导入 ${papers.length} 张1号报价表${craftCount ? `（含 ${craftCount} 条工艺）` : ""}：${papers.map(p => p.shortName).join("、")}`;
      const errMsg = errors.length ? `<br><span style="color:var(--danger)">警告：${errors.join("；")}</span>` : "";
      setExcelStatus(successMsg + errMsg, false);
      showToast("Excel 导入成功");
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
  // 重建纸张选择 UI、纸张设置列表、1号报价表查询页
  currentPaperIndex = Math.min(currentPaperIndex, Math.max(0, PAPER_CONFIG.length - 1));
  // 重置纸张设置，保留尺寸数据，默认使用第一张纸
  sheetsState = sheetsState.map(s => ({
    paperId: PAPER_CONFIG.find(p => p.id === s.paperId) ? s.paperId : PAPER_CONFIG[0].id,
    craftIds: [],
    width: s.width || "",
    length: s.length || "",
    sizeType: s.sizeType || "custom"
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
    if (!isNaN(idx) && idx >= 0 && idx < PAPER_CONFIG.length) {
      currentPaperIndex = idx;
      renderPriceTable();
    }
  });

  // 个人主页事件
  if (els.addLevelBtn) els.addLevelBtn.addEventListener("click", addCustomerLevel);
  if (els.resetLevelBtn) els.resetLevelBtn.addEventListener("click", resetCustomerLevels);
  if (els.saveProfileBtn) els.saveProfileBtn.addEventListener("click", saveProfile);
  if (els.exportProfileBtn) els.exportProfileBtn.addEventListener("click", exportProfile);

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
    ["renderHistory", renderHistory]
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

  // 应用个人主页的默认纸张材质设置
  try {
    if (APP_PROFILE.defaultPaperId) {
      const idx = PAPER_CONFIG.findIndex(p => p.id === APP_PROFILE.defaultPaperId);
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