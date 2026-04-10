const state = {
  fields: [],
  regions: []
};
let pendingScrollTargetId = null;
let miniMapViewFieldId = "__none__";
let selectedShelfIds = new Set();
let miniMapPointerState = null;
let miniMapInteractionsInitialized = false;
/** 导出 Excel 列开关：key 为 CORE:表头 或 FIELD:字段id；缺省或未写 false 表示导出 */
let excelExportEnabled = {};
/** 导出列顺序（与弹窗列表一致），元素为 CORE:区域编号 或 FIELD:字段id */
let excelExportOrder = [];
let excelExportDraggedRow = null;
const STORAGE_KEY = "warehouse_location_tool_v6";
let persistTimer = null;

const EXCEL_CORE_HEADERS = ["区域编号", "货架编号", "层级编号", "货道编号"];

const fixedFieldRow = { key: "rows", label: "行数", defaultValue: "7" };
const fixedFieldCol = { key: "cols", label: "列数", defaultValue: "9" };

const defaultFields = [
  {
    id: uid(),
    name: "库区分类",
    options: ["合格区", "不合格区", "退货区", "暂存区"],
    symbols: ["HG", "BHG", "TH", "ZS"]
  },
  {
    id: uid(),
    name: "处方分类",
    options: ["处方药", "非处方药", "其他"],
    symbols: ["RX", "OTC", "QT"]
  },
  {
    id: uid(),
    name: "存储分类",
    options: ["内服药品", "外用药品", "针剂", "保健品", "食品", "医疗器械", "日用"],
    symbols: ["NF", "WY", "ZJ", "BJP", "SP", "YLQX", "RY"]
  }
];

init();

function init() {
  const saved = loadPersistedState();
  if (saved && Array.isArray(saved.fields) && saved.fields.length && Array.isArray(saved.regions) && saved.regions.length) {
    state.fields = saved.fields;
    state.regions = saved.regions;
    miniMapViewFieldId = saved.miniMapViewFieldId || "__none__";
    applyExcelExportPersistence(saved);
    bindEvents();
    if (saved.pdfSettings) applyPdfFormValues(saved.pdfSettings);
    if (saved.miniMapCollapsed) {
      const panel = document.getElementById("miniMapPanel");
      panel.classList.add("collapsed");
      panel.classList.remove("expanded");
      document.body.classList.add("mini-map-collapsed");
      const btn = document.getElementById("toggleMiniMapBtn");
      if (btn) btn.textContent = "展开";
    }
  } else {
    state.fields = structuredClone(defaultFields);
    state.regions = [
      {
        id: uid(),
        name: "A",
        shelves: [buildShelf("01")]
      }
    ];
    excelExportOrder = [];
    excelExportEnabled = {};
    bindEvents();
  }
  renderAll();
}

function migrateLegacyExcelExportKeys(obj) {
  const out = {};
  if (!obj || typeof obj !== "object") return out;
  Object.entries(obj).forEach(([k, v]) => {
    if (k.startsWith("CORE:") || k.startsWith("FIELD:")) {
      out[k] = v;
    } else if (EXCEL_CORE_HEADERS.includes(k)) {
      out[`CORE:${k}`] = v;
    } else {
      out[`FIELD:${k}`] = v;
    }
  });
  return out;
}

function applyExcelExportPersistence(saved) {
  if (!saved) return;
  if (Array.isArray(saved.excelExportOrder)) {
    excelExportOrder = [...saved.excelExportOrder];
  }
  if (saved.excelExportEnabled && typeof saved.excelExportEnabled === "object") {
    excelExportEnabled = migrateLegacyExcelExportKeys(saved.excelExportEnabled);
  }
}

function loadPersistedState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function schedulePersist() {
  if (persistTimer) clearTimeout(persistTimer);
  persistTimer = setTimeout(() => {
    persistTimer = null;
    persistToStorage();
  }, 350);
}

function persistToStorage() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        version: 6,
        fields: state.fields,
        regions: state.regions,
        miniMapViewFieldId,
        excelExportEnabled,
        excelExportOrder,
        pdfSettings: collectPdfFormValues(),
        miniMapCollapsed: document.getElementById("miniMapPanel").classList.contains("collapsed")
      })
    );
  } catch (e) {
    console.warn("localStorage 保存失败", e);
  }
}

function collectPdfFormValues() {
  return {
    labelSizeMode: document.getElementById("labelSizeMode").value,
    paperSizeMode: document.getElementById("paperSizeMode").value,
    labelWidth: document.getElementById("labelWidth").value,
    labelHeight: document.getElementById("labelHeight").value,
    paperWidth: document.getElementById("paperWidth").value,
    paperHeight: document.getElementById("paperHeight").value,
    layoutDirection: document.getElementById("layoutDirection").value,
    verticalTopContent: document.getElementById("verticalTopContent").value,
    horizontalLeftContent: document.getElementById("horizontalLeftContent").value,
    codeFontSize: document.getElementById("codeFontSize").value,
    sepRegionShelf: document.getElementById("sepRegionShelf").value,
    sepShelfLevel: document.getElementById("sepShelfLevel").value,
    sepLevelAisle: document.getElementById("sepLevelAisle").value
  };
}

function applyPdfFormValues(p) {
  if (!p) return;
  const entries = [
    ["labelSizeMode", p.labelSizeMode],
    ["paperSizeMode", p.paperSizeMode],
    ["labelWidth", p.labelWidth],
    ["labelHeight", p.labelHeight],
    ["paperWidth", p.paperWidth],
    ["paperHeight", p.paperHeight],
    ["layoutDirection", p.layoutDirection],
    ["verticalTopContent", p.verticalTopContent],
    ["horizontalLeftContent", p.horizontalLeftContent],
    ["codeFontSize", p.codeFontSize],
    ["sepRegionShelf", p.sepRegionShelf],
    ["sepShelfLevel", p.sepShelfLevel],
    ["sepLevelAisle", p.sepLevelAisle]
  ];
  entries.forEach(([id, v]) => {
    const el = document.getElementById(id);
    if (!el || v === undefined || v === null) return;
    el.value = String(v);
  });
}

function bindEvents() {
  const tutorialDialog = document.getElementById("tutorialDialog");
  document.getElementById("openTutorialBtn").addEventListener("click", () => {
    tutorialDialog.showModal();
  });
  document.getElementById("closeTutorialBtn").addEventListener("click", () => {
    tutorialDialog.close();
  });

  document.getElementById("importExcelBtn").addEventListener("click", () => {
    document.getElementById("importExcelInput").click();
  });
  document.getElementById("importExcelInput").addEventListener("change", handleImportExcel);
  document.getElementById("importReportCloseBtn").addEventListener("click", () => {
    document.getElementById("importReportDialog").close();
  });

  document.getElementById("addFieldBtn").addEventListener("click", () => {
    const newField = {
      id: uid(),
      name: "自定义字段",
      options: ["选项1", "选项2"],
      symbols: ["X1", "X2"]
    };
    state.fields.push(newField);
    state.regions.forEach((region) => {
      region.shelves.forEach((shelf) => {
        shelf.businessValues[newField.id] = newField.options[0];
      });
    });
    renderAll();
  });

  document.getElementById("addRegionBtn").addEventListener("click", () => {
    state.regions.push({
      id: uid(),
      name: nextRegionName(),
      shelves: [buildShelf("01")]
    });
    renderRegions();
  });

  document.getElementById("exportExcelBtn").addEventListener("click", () => {
    renderExcelExportDialog();
    document.getElementById("exportExcelDialog").showModal();
  });
  document.getElementById("cancelExportExcelBtn").addEventListener("click", () => {
    document.getElementById("exportExcelDialog").close();
  });
  document.getElementById("confirmExportExcelBtn").addEventListener("click", () => {
    readExcelExportEnabledFromDialog();
    if (!performExcelExport()) return;
    schedulePersist();
    document.getElementById("exportExcelDialog").close();
  });
  document.getElementById("exportExcelSelectAllBtn").addEventListener("click", () => {
    setAllExcelExportCheckboxes(true);
  });
  document.getElementById("exportExcelSelectNoneBtn").addEventListener("click", () => {
    setAllExcelExportCheckboxes(false);
  });

  const dialog = document.getElementById("pdfSettingsDialog");
  document.getElementById("openPdfSettingsBtn").addEventListener("click", () => {
    dialog.showModal();
    updatePdfPreview();
  });
  document.getElementById("cancelPdfBtn").addEventListener("click", () => dialog.close());
  document.getElementById("pdfSettingsForm").addEventListener("submit", (e) => {
    e.preventDefault();
    exportPdf();
    schedulePersist();
    dialog.close();
  });

  document.getElementById("labelSizeMode").addEventListener("change", (e) => {
    applyPresetSize(e.target.value, "labelWidth", "labelHeight");
    updatePdfPreview();
    schedulePersist();
  });
  document.getElementById("paperSizeMode").addEventListener("change", (e) => {
    applyPresetSize(e.target.value, "paperWidth", "paperHeight");
    updatePdfPreview();
    schedulePersist();
  });

  const pdfPreviewInputIds = [
    "labelWidth",
    "labelHeight",
    "paperWidth",
    "paperHeight",
    "layoutDirection",
    "verticalTopContent",
    "horizontalLeftContent",
    "codeFontSize",
    "sepRegionShelf",
    "sepShelfLevel",
    "sepLevelAisle"
  ];
  pdfPreviewInputIds.forEach((id) => {
    const el = document.getElementById(id);
    el.addEventListener("input", () => {
      updatePdfPreview();
      schedulePersist();
    });
    el.addEventListener("change", () => {
      updatePdfPreview();
      schedulePersist();
    });
  });

  document.getElementById("toggleMiniMapBtn").addEventListener("click", () => {
    const panel = document.getElementById("miniMapPanel");
    panel.classList.toggle("collapsed");
    panel.classList.toggle("expanded");
    const collapsed = panel.classList.contains("collapsed");
    document.body.classList.toggle("mini-map-collapsed", collapsed);
    document.getElementById("toggleMiniMapBtn").textContent = collapsed ? "展开" : "收起";
    schedulePersist();
  });

  document.getElementById("miniMapViewFieldSelect").addEventListener("change", (e) => {
    miniMapViewFieldId = e.target.value;
    renderMiniMap();
    schedulePersist();
  });

  initMiniMapShelfInteractions();
}

function buildShelf(code) {
  const businessValues = {};
  state.fields.forEach((f) => {
    businessValues[f.id] = f.options[0] || "";
  });
  return {
    id: uid(),
    code,
    rows: fixedFieldRow.defaultValue,
    cols: fixedFieldCol.defaultValue,
    businessValues
  };
}

function renderAll() {
  pruneSelectedShelves();
  renderFields();
  renderRegions();
  renderMiniMapFieldOptions();
  renderMiniMap();
  flushPendingScroll();
  schedulePersist();
}

function renderFields() {
  const container = document.getElementById("fieldsContainer");
  container.innerHTML = "";

  state.fields.forEach((field) => {
    const card = document.createElement("div");
    card.className = "field-card";
    card.innerHTML = `
      <div class="field-row">
        <input data-role="field-name" placeholder="字段名称" value="${escapeHtml(field.name)}">
        <input data-role="field-options" placeholder="枚举值(逗号分隔)" value="${escapeHtml(field.options.join(","))}">
        <input data-role="field-symbols" placeholder="枚举符号(逗号分隔)" value="${escapeHtml(field.symbols.join(","))}">
        <button type="button" data-role="delete-field">删除</button>
      </div>
    `;

    card.querySelector('[data-role="field-name"]').addEventListener("input", (e) => {
      field.name = e.target.value.trim() || "未命名字段";
      renderMiniMapFieldOptions();
      renderMiniMap();
      schedulePersist();
    });
    card.querySelector('[data-role="field-options"]').addEventListener("change", (e) => {
      const nextOptions = parseCsv(e.target.value);
      field.options = nextOptions.length ? nextOptions : ["默认值"];
      state.regions.forEach((region) => {
        region.shelves.forEach((shelf) => {
          const current = shelf.businessValues[field.id];
          if (!field.options.includes(current)) {
            shelf.businessValues[field.id] = field.options[0];
          }
        });
      });
      renderAll();
    });
    card.querySelector('[data-role="field-symbols"]').addEventListener("change", (e) => {
      const symbols = parseCsv(e.target.value);
      field.symbols = symbols;
      renderFields();
      schedulePersist();
    });
    card.querySelector('[data-role="delete-field"]').addEventListener("click", () => {
      state.fields = state.fields.filter((f) => f.id !== field.id);
      state.regions.forEach((region) => {
        region.shelves.forEach((shelf) => {
          delete shelf.businessValues[field.id];
        });
      });
      renderAll();
    });

    container.appendChild(card);
  });
}

function renderRegions() {
  const container = document.getElementById("regionsContainer");
  container.innerHTML = "";

  state.regions.forEach((region) => {
    const regionCard = document.createElement("div");
    regionCard.className = "region-card";
    regionCard.id = `region-${region.id}`;
    regionCard.innerHTML = `
      <div class="region-title">
        <h3>
          区域
          <input data-role="region-name" type="text" value="${escapeHtml(region.name)}" style="width:90px;display:inline-block;margin:0 6px;">
          <span class="badge">${region.shelves.length} 个货架</span>
        </h3>
        <div>
          <button type="button" data-role="add-shelf">新增货架</button>
          <button type="button" data-role="copy-region">复制区域</button>
          <button type="button" data-role="delete-region">删除区域</button>
        </div>
      </div>
      <div class="shelves-grid"></div>
    `;

    regionCard.querySelector('[data-role="add-shelf"]').addEventListener("click", () => {
      region.shelves.push(buildShelf(nextShelfCode(region)));
      renderRegions();
    });
    regionCard.querySelector('[data-role="region-name"]').addEventListener("change", (e) => {
      const next = e.target.value.trim();
      if (!next) {
        e.target.value = region.name;
        return;
      }
      region.name = next;
      renderAll();
    });

    regionCard.querySelector('[data-role="copy-region"]').addEventListener("click", () => {
      const copied = structuredClone(region);
      copied.id = uid();
      copied.name = nextRegionName();
      copied.shelves = copied.shelves.map((s) => ({ ...s, id: uid() }));
      state.regions.push(copied);
      pendingScrollTargetId = `region-${copied.id}`;
      renderAll();
    });

    regionCard.querySelector('[data-role="delete-region"]').addEventListener("click", () => {
      if (state.regions.length === 1) {
        alert("至少保留一个区域。");
        return;
      }
      state.regions = state.regions.filter((r) => r.id !== region.id);
      renderAll();
    });

    const shelvesGrid = regionCard.querySelector(".shelves-grid");
    region.shelves.forEach((shelf) => {
      shelvesGrid.appendChild(renderShelf(region, shelf));
    });
    container.appendChild(regionCard);
  });

  renderMiniMap();
  flushPendingScroll();
  schedulePersist();
}

function renderShelf(region, shelf) {
  const card = document.createElement("div");
  card.className = "shelf-card";
  if (selectedShelfIds.has(shelf.id)) card.classList.add("shelf-card-selected");
  card.id = `shelf-${shelf.id}`;

  const businessFieldHtml = state.fields.map((field) => {
    const options = field.options
      .map((op) => `<option value="${escapeHtml(op)}" ${shelf.businessValues[field.id] === op ? "selected" : ""}>${escapeHtml(op)}</option>`)
      .join("");
    return `
      <label>${escapeHtml(field.name)}
        <select data-role="business-select" data-field-id="${field.id}">
          ${options}
        </select>
      </label>
    `;
  }).join("");

  card.innerHTML = `
    <div class="shelf-header">
      <span class="shelf-title">
        货架
        <input data-role="shelf-code" type="text" value="${escapeHtml(shelf.code)}" style="width:70px;display:inline-block;margin-left:6px;">
      </span>
      <div>
        <button type="button" data-role="copy-shelf">复制</button>
        <button type="button" data-role="delete-shelf">删除</button>
      </div>
    </div>
    <div class="shelf-fields">
      <label>${fixedFieldRow.label}
        <input data-role="rows" type="number" min="1" max="99" value="${escapeHtml(String(shelf.rows))}">
      </label>
      <label>${fixedFieldCol.label}
        <input data-role="cols" type="number" min="1" max="99" value="${escapeHtml(String(shelf.cols))}">
      </label>
      ${businessFieldHtml}
    </div>
  `;

  const rowsInput = card.querySelector('[data-role="rows"]');
  const colsInput = card.querySelector('[data-role="cols"]');
  rowsInput.addEventListener("input", () => {
    shelf.rows = sanitizeTwoDigitNum(rowsInput.value, 7);
    rowsInput.value = shelf.rows;
    renderMiniMap();
    schedulePersist();
  });
  colsInput.addEventListener("input", () => {
    shelf.cols = sanitizeTwoDigitNum(colsInput.value, 9);
    colsInput.value = shelf.cols;
    renderMiniMap();
    schedulePersist();
  });
  card.querySelector('[data-role="shelf-code"]').addEventListener("change", (e) => {
    const next = e.target.value.trim();
    if (!next) {
      e.target.value = shelf.code;
      return;
    }
    shelf.code = next;
    renderAll();
  });

  card.querySelectorAll('[data-role="business-select"]').forEach((el) => {
    el.addEventListener("change", (e) => {
      shelf.businessValues[e.target.dataset.fieldId] = e.target.value;
      renderMiniMap();
      schedulePersist();
    });
  });

  card.addEventListener("click", (e) => {
    if (e.target.closest("input, select, button, textarea, label")) return;
    const additive = e.ctrlKey || e.metaKey;
    if (additive) {
      if (selectedShelfIds.has(shelf.id)) selectedShelfIds.delete(shelf.id);
      else selectedShelfIds.add(shelf.id);
    } else {
      selectedShelfIds = new Set([shelf.id]);
    }
    updateSelectionUi();
  });

  card.querySelector('[data-role="copy-shelf"]').addEventListener("click", () => {
    const cloned = structuredClone(shelf);
    cloned.id = uid();
    cloned.code = nextShelfCode(region);
    region.shelves.push(cloned);
    pendingScrollTargetId = `shelf-${cloned.id}`;
    renderAll();
  });

  card.querySelector('[data-role="delete-shelf"]').addEventListener("click", () => {
    if (region.shelves.length === 1) {
      alert("每个区域至少保留一个货架。");
      return;
    }
    region.shelves = region.shelves.filter((s) => s.id !== shelf.id);
    renderAll();
  });

  return card;
}

function excelColumnLetter(oneBasedIndex) {
  let s = "";
  let n = oneBasedIndex;
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function isValidExcelExportToken(t) {
  if (typeof t !== "string") return false;
  if (t.startsWith("CORE:")) return EXCEL_CORE_HEADERS.includes(t.slice(5));
  if (t.startsWith("FIELD:")) return state.fields.some((f) => f.id === t.slice(6));
  return false;
}

function getDefaultExcelExportOrderTokens() {
  return [...EXCEL_CORE_HEADERS.map((h) => `CORE:${h}`), ...state.fields.map((f) => `FIELD:${f.id}`)];
}

function getNormalizedExcelExportOrder() {
  const defaults = getDefaultExcelExportOrderTokens();
  const seen = new Set();
  const out = [];
  excelExportOrder.forEach((t) => {
    if (seen.has(t) || !isValidExcelExportToken(t)) return;
    out.push(t);
    seen.add(t);
  });
  defaults.forEach((t) => {
    if (!seen.has(t)) {
      out.push(t);
      seen.add(t);
    }
  });
  return out;
}

function getOrderedExcelExportSpecs() {
  const specs = [];
  getNormalizedExcelExportOrder().forEach((token) => {
    if (excelExportEnabled[token] === false) return;
    if (token.startsWith("CORE:")) {
      const h = token.slice(5);
      specs.push({ kind: "core", header: h });
    } else {
      const id = token.slice(6);
      const f = state.fields.find((x) => x.id === id);
      if (f) specs.push({ kind: "field", header: f.name, fieldId: f.id });
    }
  });
  return specs;
}

function buildExcelRowFromSpecs(row, specs) {
  const line = {};
  specs.forEach((spec) => {
    if (spec.kind === "core") {
      if (spec.header === "区域编号") line[spec.header] = row.region;
      else if (spec.header === "货架编号") line[spec.header] = row.shelf;
      else if (spec.header === "层级编号") line[spec.header] = row.level;
      else line[spec.header] = row.aisle;
    } else {
      line[spec.header] = row.business[spec.fieldId] || "";
    }
  });
  return line;
}

function renderExcelExportDialog() {
  const container = document.getElementById("exportExcelColumnList");
  container.querySelectorAll(".export-excel-row").forEach((r) => r.remove());

  let indicator = container.querySelector(".export-excel-drop-indicator");
  if (!indicator) {
    indicator = document.createElement("div");
    indicator.className = "export-excel-drop-indicator";
    indicator.setAttribute("aria-hidden", "true");
    indicator.hidden = true;
    container.appendChild(indicator);
  }

  getNormalizedExcelExportOrder().forEach((token) => {
    container.insertBefore(createExcelExportRowFromToken(token), indicator);
  });

  updateExcelExportColumnBadges(container);
  bindExportExcelListInteractions(container);
}

function getExportDropInsertInfo(clientY, container, dragRow) {
  const rows = [...container.querySelectorAll(".export-excel-row")].filter((r) => r !== dragRow);
  const cRect = container.getBoundingClientRect();
  const scrollTop = container.scrollTop;
  if (!rows.length) {
    return { insertBefore: null, lineTop: 6 };
  }
  for (let i = 0; i < rows.length; i += 1) {
    const r = rows[i];
    const rect = r.getBoundingClientRect();
    if (clientY < rect.top + rect.height / 2) {
      const lineTop = rect.top - cRect.top + scrollTop - 2;
      return { insertBefore: r, lineTop: Math.max(4, lineTop) };
    }
  }
  const last = rows[rows.length - 1];
  const lr = last.getBoundingClientRect();
  const lineTop = lr.bottom - cRect.top + scrollTop - 2;
  return { insertBefore: null, lineTop: Math.max(4, lineTop) };
}

function exportRowIsAlreadyAtTarget(container, dragRow, insertBefore) {
  const indicator = container.querySelector(".export-excel-drop-indicator");
  if (insertBefore) {
    return dragRow.nextElementSibling === insertBefore;
  }
  if (indicator) {
    return dragRow.nextElementSibling === indicator;
  }
  const rows = container.querySelectorAll(".export-excel-row");
  return rows.length > 0 && rows[rows.length - 1] === dragRow;
}

function bindExportExcelListInteractions(container) {
  if (container.dataset.exportDragBound === "1") return;
  container.dataset.exportDragBound = "1";

  const getIndicator = () => container.querySelector(".export-excel-drop-indicator");

  container.addEventListener("dragover", (e) => {
    const dragRow = excelExportDraggedRow;
    if (!dragRow || !container.contains(dragRow)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    const indicator = getIndicator();
    if (!indicator) return;
    const { lineTop } = getExportDropInsertInfo(e.clientY, container, dragRow);
    indicator.style.top = `${lineTop}px`;
    indicator.hidden = false;
  });

  container.addEventListener("drop", (e) => {
    e.preventDefault();
    const dragRow = excelExportDraggedRow;
    const indicator = getIndicator();
    if (indicator) indicator.hidden = true;
    if (!dragRow || !container.contains(dragRow)) return;

    const { insertBefore } = getExportDropInsertInfo(e.clientY, container, dragRow);
    if (exportRowIsAlreadyAtTarget(container, dragRow, insertBefore)) {
      updateExcelExportColumnBadges(container);
      return;
    }

    reorderExportRowsAnimated(container, dragRow, insertBefore);
    readExcelExportEnabledFromDialog();
    updateExcelExportColumnBadges(container);
    schedulePersist();
  });

  container.addEventListener("dragleave", (e) => {
    if (e.target !== container) return;
    const rel = e.relatedTarget;
    if (rel && container.contains(rel)) return;
    const ind = getIndicator();
    if (ind) ind.hidden = true;
  });

  if (!window.__excelExportDocDragEnd) {
    window.__excelExportDocDragEnd = true;
    document.addEventListener(
      "dragend",
      () => {
        const ind = document.querySelector("#exportExcelColumnList .export-excel-drop-indicator");
        if (ind) ind.hidden = true;
      },
      true
    );
  }
}

function reorderExportRowsAnimated(container, dragRow, insertBeforeEl) {
  const indicator = container.querySelector(".export-excel-drop-indicator");
  const getRows = () => [...container.querySelectorAll(".export-excel-row")];

  const beforeRects = new Map();
  getRows().forEach((r) => beforeRects.set(r, r.getBoundingClientRect()));

  if (insertBeforeEl && insertBeforeEl !== dragRow) {
    container.insertBefore(dragRow, insertBeforeEl);
  } else if (indicator) {
    container.insertBefore(dragRow, indicator);
  } else {
    container.appendChild(dragRow);
  }

  requestAnimationFrame(() => {
    const rowList = getRows();
    const animated = new Set();
    rowList.forEach((r) => {
      const fr = beforeRects.get(r);
      if (!fr) return;
      const lr = r.getBoundingClientRect();
      const dy = fr.top - lr.top;
      if (Math.abs(dy) < 0.5) return;
      animated.add(r);
      r.style.transition = "none";
      r.style.transform = `translateY(${dy}px)`;
    });
    requestAnimationFrame(() => {
      rowList.forEach((r, i) => {
        if (!animated.has(r)) return;
        const delayMs = Math.min(i * 10, 72);
        r.style.transition = `transform 0.3s cubic-bezier(0.22, 1, 0.36, 1) ${delayMs}ms`;
        r.style.transform = "";
      });
    });
    const clear = () => {
      getRows().forEach((r) => {
        r.style.transition = "";
        r.style.removeProperty("transform");
      });
    };
    window.setTimeout(clear, 420);
  });
}

function createExcelExportRowFromToken(token) {
  const row = document.createElement("div");
  row.className = "export-excel-row";
  row.dataset.exportToken = token;

  const handle = document.createElement("span");
  handle.className = "export-excel-drag-handle";
  handle.draggable = true;
  handle.title = "拖动调整导出列顺序";
  handle.textContent = "⠿";
  handle.addEventListener("dragstart", (e) => {
    e.stopPropagation();
    excelExportDraggedRow = row;
    row.classList.add("export-excel-row--dragging");
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", token);
  });
  handle.addEventListener("dragend", () => {
    row.classList.remove("export-excel-row--dragging");
    excelExportDraggedRow = null;
  });

  let labelText;
  if (token.startsWith("CORE:")) {
    labelText = token.slice(5);
  } else {
    const f = state.fields.find((x) => x.id === token.slice(6));
    labelText = f ? f.name : token;
  }

  const checked = excelExportEnabled[token] !== false;

  const badge = document.createElement("span");
  badge.className = "export-excel-col-badge";
  badge.dataset.role = "col-badge";

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = checked;
  cb.dataset.exportKey = token;
  cb.addEventListener("change", () => {
    updateExcelExportColumnBadges(row.parentElement);
  });

  const lab = document.createElement("label");
  lab.appendChild(cb);
  const span = document.createElement("span");
  span.textContent = labelText;
  lab.appendChild(span);

  row.appendChild(handle);
  row.appendChild(lab);
  row.appendChild(badge);

  return row;
}

function updateExcelExportColumnBadges(container) {
  if (!container) return;
  let col = 1;
  container.querySelectorAll(".export-excel-row").forEach((row) => {
    const cb = row.querySelector('input[type="checkbox"]');
    const badge = row.querySelector('[data-role="col-badge"]');
    if (!cb || !badge) return;
    if (cb.checked) {
      const letter = excelColumnLetter(col);
      badge.textContent = `${letter}列`;
      col += 1;
      badge.style.opacity = "1";
    } else {
      badge.textContent = "不导出";
      badge.style.opacity = "0.55";
    }
  });
}

function readExcelExportEnabledFromDialog() {
  const container = document.getElementById("exportExcelColumnList");
  excelExportOrder = [...container.querySelectorAll(".export-excel-row[data-export-token]")].map((r) => r.dataset.exportToken);
  container.querySelectorAll("input[type=checkbox][data-export-key]").forEach((cb) => {
    excelExportEnabled[cb.dataset.exportKey] = cb.checked;
  });
}

function setAllExcelExportCheckboxes(checked) {
  const container = document.getElementById("exportExcelColumnList");
  container.querySelectorAll('input[type=checkbox][data-export-key]').forEach((cb) => {
    cb.checked = checked;
  });
  updateExcelExportColumnBadges(container);
}

function performExcelExport() {
  const specs = getOrderedExcelExportSpecs();
  if (!specs.length) {
    alert("请至少勾选一列再导出。");
    return false;
  }
  const rows = buildLocationRows();
  const wsData = rows.map((row) => buildExcelRowFromSpecs(row, specs));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "货位编码模板");
  XLSX.writeFile(wb, `货位码导入模板_${nowDateText()}.xlsx`);
  return true;
}

async function exportPdf() {
  const rows = buildLocationRows();
  if (!rows.length) {
    alert("暂无可导出的货位编码。");
    return;
  }

  const { jsPDF } = window.jspdf;
  const paper = readMmSize("paper");
  const label = readMmSize("label");
  const pdfSettings = readPdfSettingsForExport();

  const pdf = new jsPDF({
    orientation: paper.w >= paper.h ? "landscape" : "portrait",
    unit: "mm",
    format: [paper.w, paper.h]
  });

  for (let i = 0; i < rows.length; i += 1) {
    if (i > 0) pdf.addPage([paper.w, paper.h], paper.w >= paper.h ? "landscape" : "portrait");
    const row = rows[i];
    const codeForBarcode = buildLocationCode(row, pdfSettings.separators, "barcode");
    const codeForText = buildLocationCode(row, pdfSettings.separators, "text");
    const barcodeDataUrl = makeBarcode(codeForBarcode);
    const offsetX = (paper.w - label.w) / 2;
    const offsetY = (paper.h - label.h) / 2;
    drawSingleLabelOnPdf(pdf, offsetX, offsetY, label.w, label.h, codeForText, barcodeDataUrl, pdfSettings);
  }

  pdf.save(`货位码标签_${nowDateText()}.pdf`);
}

function readPdfSettingsForExport() {
  const raw = document.getElementById("codeFontSize").value;
  const fontSize = Number.parseFloat(raw);
  const fontSizePt = Number.isFinite(fontSize) ? fontSize : 24;
  return {
    fontSizePt,
    layoutDirection: document.getElementById("layoutDirection").value,
    verticalTopContent: document.getElementById("verticalTopContent").value,
    horizontalLeftContent: document.getElementById("horizontalLeftContent").value,
    separators: {
      rs: document.getElementById("sepRegionShelf").value || "",
      sl: document.getElementById("sepShelfLevel").value || "",
      la: document.getElementById("sepLevelAisle").value || ""
    }
  };
}

function getLabelContentFractions(fontSizePt, layoutDirection) {
  const textWeight = Math.min(2, Math.max(0.6, fontSizePt / 20));
  if (layoutDirection === "vertical") {
    const textRatio = Math.min(0.55, Math.max(0.24, 0.24 + textWeight * 0.08));
    return { textRatio, barcodeRatio: 1 - textRatio };
  }
  const textRatio = Math.min(0.58, Math.max(0.24, 0.24 + textWeight * 0.09));
  return { textRatio, barcodeRatio: 1 - textRatio };
}

function drawSingleLabelOnPdf(pdf, offsetX, offsetY, labelW, labelH, code, barcodeDataUrl, settings) {
  const { fontSizePt, layoutDirection, verticalTopContent, horizontalLeftContent, separators } = settings;
  const { textRatio, barcodeRatio } = getLabelContentFractions(fontSizePt, layoutDirection);
  pdf.setDrawColor(180);
  pdf.rect(offsetX, offsetY, labelW, labelH);

  if (layoutDirection === "vertical") {
    if (verticalTopContent === "barcode") {
      drawBarcode(pdf, barcodeDataUrl, offsetX, offsetY, labelW, labelH * barcodeRatio);
      drawCodeTextFixedPt(pdf, code, offsetX, offsetY + labelH * barcodeRatio, labelW, labelH * textRatio, fontSizePt, separators);
    } else {
      drawCodeTextFixedPt(pdf, code, offsetX, offsetY, labelW, labelH * textRatio, fontSizePt, separators);
      drawBarcode(pdf, barcodeDataUrl, offsetX, offsetY + labelH * textRatio, labelW, labelH * barcodeRatio);
    }
  } else if (horizontalLeftContent === "barcode") {
    drawBarcode(pdf, barcodeDataUrl, offsetX, offsetY, labelW * barcodeRatio, labelH);
    drawCodeTextFixedPt(pdf, code, offsetX + labelW * barcodeRatio, offsetY, labelW * textRatio, labelH, fontSizePt, separators);
  } else {
    drawCodeTextFixedPt(pdf, code, offsetX, offsetY, labelW * textRatio, labelH, fontSizePt, separators);
    drawBarcode(pdf, barcodeDataUrl, offsetX + labelW * textRatio, offsetY, labelW * barcodeRatio, labelH);
  }
}

function buildLocationRows() {
  const out = [];
  state.regions.forEach((region) => {
    region.shelves.forEach((shelf) => {
      const rows = sanitizeTwoDigitNum(shelf.rows, 7);
      const cols = sanitizeTwoDigitNum(shelf.cols, 9);
      for (let r = 1; r <= Number(rows); r += 1) {
        for (let c = 1; c <= Number(cols); c += 1) {
          out.push({
            region: region.name,
            shelf: shelf.code,
            level: pad2(r),
            aisle: pad2(c),
            business: structuredClone(shelf.businessValues)
          });
        }
      }
    });
  });
  return out;
}

function drawBarcode(pdf, dataUrl, x, y, w, h) {
  pdf.addImage(dataUrl, "PNG", x + 2, y + 2, Math.max(w - 4, 10), Math.max(h - 8, 10));
}

function computeCodeLines(pdf, text, wMm, hMm, fontSizePt, separators) {
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(fontSizePt);
  const paddedW = Math.max(8, wMm - 4);
  const paddedHmm = Math.max(8, hMm - 4);
  const lines = wrapCodeByWidth(pdf, text, paddedW, separators);
  const ptToMm = 25.4 / 72;
  const lineHeightMm = fontSizePt * ptToMm * 1.2;
  return { lines, lineHeightMm, paddedHmm };
}

function drawCodeTextFixedPt(pdf, text, x, y, w, h, fontSizePt, separators) {
  const { lines, lineHeightMm, paddedHmm } = computeCodeLines(pdf, text, w, h, fontSizePt, separators);
  const ptToMm = 25.4 / 72;
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(fontSizePt);
  const centerX = x + w / 2;
  const topY = y + 2;
  const contentHeight = lines.length * lineHeightMm;
  const startY = topY + Math.max(0, (paddedHmm - contentHeight) / 2) + fontSizePt * ptToMm * 0.88;
  lines.forEach((line, idx) => {
    pdf.text(line, centerX, startY + idx * lineHeightMm, { align: "center" });
  });
}

function wrapCodeByWidth(pdf, text, maxWidth, separators) {
  const hardBreakLines = String(text).split("\n");
  const lines = [];
  hardBreakLines.forEach((partLine) => {
    const rawParts = splitByCustomSeparators(partLine, separators);
    let current = "";
    rawParts.forEach((token) => {
      const candidate = `${current}${token}`;
      if (!current || pdf.getTextWidth(candidate) <= maxWidth) {
        current = candidate;
      } else {
        lines.push(current);
        current = token;
      }
    });
    if (current) lines.push(current);
  });
  return lines.length ? lines : [""];
}

function splitByCustomSeparators(text, separators) {
  const symbols = [separators.rs, separators.sl, separators.la].filter(Boolean);
  if (!symbols.length) return [text];
  const escaped = symbols.map((s) => escapeRegExp(s));
  const regex = new RegExp(`(${escaped.join("|")})`);
  return text.split(regex).filter((part) => part.length > 0);
}

/** 条码编码内容固定为四段 + 短横线，与「标签连接符」设置无关（连接符仅用于可见文字）。 */
function buildLocationCode(row, separators, mode = "barcode") {
  if (mode === "barcode") {
    return `${row.region}-${row.shelf}-${row.level}-${row.aisle}`;
  }
  const mapSep = (s) => {
    if (/\s/.test(s || "")) return "\n";
    return s || "";
  };
  const rs = mapSep(separators.rs);
  const sl = mapSep(separators.sl);
  const la = mapSep(separators.la);
  return `${row.region}${rs}${row.shelf}${sl}${row.level}${la}${row.aisle}`;
}

function getSampleLocationRow() {
  const r0 = state.regions[0];
  if (!r0) return { region: "A", shelf: "01", level: "01", aisle: "01" };
  const s0 = r0.shelves[0];
  if (!s0) return { region: r0.name, shelf: "01", level: "01", aisle: "01" };
  return { region: r0.name, shelf: s0.code, level: "01", aisle: "01" };
}

function getPreviewSampleRow() {
  const rows = buildLocationRows();
  if (rows.length) {
    const r = rows[0];
    return { region: r.region, shelf: r.shelf, level: r.level, aisle: r.aisle };
  }
  return getSampleLocationRow();
}

function updatePdfPreview() {
  const host = document.getElementById("pdfPreviewHost");
  if (!host) return;

  const paper = readMmSize("paper");
  const label = readMmSize("label");
  const settings = readPdfSettingsForExport();
  const sample = getPreviewSampleRow();
  const codeForBarcode = buildLocationCode(sample, settings.separators, "barcode");
  const codeForText = buildLocationCode(sample, settings.separators, "text");

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: paper.w >= paper.h ? "landscape" : "portrait",
    unit: "mm",
    format: [paper.w, paper.h]
  });
  const barcodeUrl = makeBarcode(codeForBarcode);
  const offsetX = (paper.w - label.w) / 2;
  const offsetY = (paper.h - label.h) / 2;
  drawSingleLabelOnPdf(pdf, offsetX, offsetY, label.w, label.h, codeForText, barcodeUrl, settings);

  host.innerHTML = "";
  const frame = document.createElement("iframe");
  frame.className = "pdf-preview-frame";
  frame.title = "PDF打印预览";
  frame.src = pdf.output("datauristring");
  host.appendChild(frame);
}

function makeBarcode(code) {
  const canvas = document.createElement("canvas");
  JsBarcode(canvas, code, {
    format: "CODE128",
    lineColor: "#000",
    width: 2,
    height: 80,
    displayValue: false,
    margin: 0
  });
  return canvas.toDataURL("image/png");
}

function nextRegionName() {
  let num = 0;
  const set = new Set(state.regions.map((r) => r.name));
  while (true) {
    const candidate = numberToLetters(num);
    if (!set.has(candidate)) return candidate;
    num += 1;
  }
}

function nextShelfCode(region) {
  const maxCode = region.shelves.reduce((max, s) => {
    const num = Number(s.code);
    if (!Number.isFinite(num)) return max;
    return Math.max(max, num);
  }, 0);
  return pad2(maxCode + 1);
}

function numberToLetters(num) {
  let n = num;
  let result = "";
  do {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return result;
}

function applyPresetSize(mode, widthId, heightId) {
  if (mode === "custom") return;
  const preset = mode === "a3" ? { w: 297, h: 420 } : { w: 210, h: 297 };
  document.getElementById(widthId).value = preset.w;
  document.getElementById(heightId).value = preset.h;
}

function readMmSize(prefix) {
  const mode = document.getElementById(`${prefix}SizeMode`).value;
  let w = Number(document.getElementById(`${prefix}Width`).value);
  let h = Number(document.getElementById(`${prefix}Height`).value);
  if (mode === "a4") ({ w, h } = { w: 210, h: 297 });
  if (mode === "a3") ({ w, h } = { w: 297, h: 420 });
  return { w: Math.max(10, w), h: Math.max(10, h) };
}

function sanitizeTwoDigitNum(value, fallback) {
  const num = Number(value);
  if (!Number.isFinite(num)) return String(fallback);
  const clamped = Math.min(99, Math.max(1, Math.floor(num)));
  return String(clamped);
}

function parseCsv(raw) {
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function uid() {
  return `${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function pad2(num) {
  return String(num).padStart(2, "0");
}

function nowDateText() {
  const d = new Date();
  const y = d.getFullYear();
  const m = pad2(d.getMonth() + 1);
  const day = pad2(d.getDate());
  const hh = pad2(d.getHours());
  const mm = pad2(d.getMinutes());
  return `${y}${m}${day}_${hh}${mm}`;
}

function renderMiniMap() {
  const container = document.getElementById("miniMapContent");
  container.innerHTML = "";
  const viewField = state.fields.find((f) => f.id === miniMapViewFieldId) || null;
  state.regions.forEach((region) => {
    const regionWrap = document.createElement("div");
    regionWrap.className = "mini-region";

    const title = document.createElement("div");
    title.className = "mini-region-title";
    title.textContent = `区域 ${region.name}`;
    title.addEventListener("click", () => smoothScrollTo(`region-${region.id}`));
    regionWrap.appendChild(title);

    const shelfGrid = document.createElement("div");
    shelfGrid.className = "mini-shelf-grid";
    region.shelves.forEach((shelf) => {
      const shelfNode = document.createElement("div");
      shelfNode.className = "mini-shelf";
      if (selectedShelfIds.has(shelf.id)) shelfNode.classList.add("mini-shelf-selected");
      shelfNode.dataset.shelfId = shelf.id;
      const labelText = viewField ? (shelf.businessValues[viewField.id] || "-") : "";
      const color = getShelfColor(viewField, labelText);
      shelfNode.style.background = color;
      const rowN = Number(sanitizeTwoDigitNum(shelf.rows, 7));
      const colN = Number(sanitizeTwoDigitNum(shelf.cols, 9));
      const slotCount = rowN * colN;
      shelfNode.innerHTML = `
        <div class="mini-shelf-top">${escapeHtml(shelf.code)}</div>
        <div class="mini-shelf-count">${slotCount}货位</div>
        <div class="mini-shelf-bottom">${escapeHtml(labelText || "未选择视图字段")}</div>
      `;
      shelfNode.title = `区域${region.name} 货架${shelf.code}（左键选中，拖拽框选，Ctrl/Cmd+点击加选，右键批量改「视图字段」）`;
      shelfGrid.appendChild(shelfNode);
    });
    regionWrap.appendChild(shelfGrid);
    container.appendChild(regionWrap);
  });
}

function renderMiniMapFieldOptions() {
  const select = document.getElementById("miniMapViewFieldSelect");
  const oldValue = miniMapViewFieldId;
  select.innerHTML = "";

  const noneOption = document.createElement("option");
  noneOption.value = "__none__";
  noneOption.textContent = "不启用";
  select.appendChild(noneOption);

  state.fields.forEach((field) => {
    const op = document.createElement("option");
    op.value = field.id;
    op.textContent = field.name;
    select.appendChild(op);
  });

  const availableValues = new Set(["__none__", ...state.fields.map((f) => f.id)]);
  miniMapViewFieldId = availableValues.has(oldValue) ? oldValue : "__none__";
  select.value = miniMapViewFieldId;
}

function getShelfColor(viewField, optionLabel) {
  if (!viewField || miniMapViewFieldId === "__none__") return "#d1d5db";
  const idx = viewField.options.findIndex((op) => op === optionLabel);
  const palette = [
    "#fca5a5",
    "#93c5fd",
    "#fde68a",
    "#86efac",
    "#f9a8d4",
    "#c4b5fd",
    "#fdba74",
    "#67e8f9",
    "#d9f99d",
    "#e5e7eb"
  ];
  if (idx < 0) return "#d1d5db";
  return palette[idx % palette.length];
}

function flushPendingScroll() {
  if (!pendingScrollTargetId) return;
  smoothScrollTo(pendingScrollTargetId);
  pendingScrollTargetId = null;
}

function smoothScrollTo(elementId) {
  const el = document.getElementById(elementId);
  if (!el) return;
  el.scrollIntoView({ behavior: "smooth", block: "start", inline: "nearest" });
}

function pruneSelectedShelves() {
  const valid = new Set();
  state.regions.forEach((r) => {
    r.shelves.forEach((s) => valid.add(s.id));
  });
  selectedShelfIds = new Set([...selectedShelfIds].filter((id) => valid.has(id)));
}

function updateSelectionUi() {
  renderMiniMap();
  syncWorkspaceShelfHighlight();
}

function syncWorkspaceShelfHighlight() {
  document.querySelectorAll(".shelf-card").forEach((card) => {
    const id = card.id.startsWith("shelf-") ? card.id.slice("shelf-".length) : "";
    if (!id) return;
    card.classList.toggle("shelf-card-selected", selectedShelfIds.has(id));
  });
}

function hideMiniMapContextMenu() {
  const menu = document.getElementById("miniMapContextMenu");
  if (menu) menu.hidden = true;
}

function showMiniMapContextMenu(clientX, clientY) {
  const menu = document.getElementById("miniMapContextMenu");
  const field = state.fields.find((f) => f.id === miniMapViewFieldId);
  menu.innerHTML = "";
  if (!field) {
    const empty = document.createElement("div");
    empty.className = "context-menu-empty";
    empty.textContent = "未找到视图字段";
    menu.appendChild(empty);
  } else {
    const title = document.createElement("div");
    title.className = "context-menu-title";
    title.textContent = `批量设置「${field.name}」· 已选 ${selectedShelfIds.size} 个货架`;
    menu.appendChild(title);
    field.options.forEach((opt) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = opt;
      btn.addEventListener("click", () => {
        applyBatchViewField(field.id, opt);
        hideMiniMapContextMenu();
      });
      menu.appendChild(btn);
    });
  }
  menu.hidden = false;
  menu.style.left = `${clientX}px`;
  menu.style.top = `${clientY}px`;
  requestAnimationFrame(() => {
    const r = menu.getBoundingClientRect();
    let x = clientX;
    let y = clientY;
    if (r.right > window.innerWidth - 8) x = window.innerWidth - r.width - 8;
    if (r.bottom > window.innerHeight - 8) y = window.innerHeight - r.height - 8;
    menu.style.left = `${Math.max(8, x)}px`;
    menu.style.top = `${Math.max(8, y)}px`;
  });
}

function applyBatchViewField(fieldId, value) {
  state.regions.forEach((region) => {
    region.shelves.forEach((shelf) => {
      if (selectedShelfIds.has(shelf.id)) {
        shelf.businessValues[fieldId] = value;
      }
    });
  });
  renderAll();
}

function positionMiniMapRubberBand(x1, y1, x2, y2) {
  const el = document.getElementById("miniMapRubberBand");
  const l = Math.min(x1, x2);
  const t = Math.min(y1, y2);
  const w = Math.abs(x2 - x1);
  const h = Math.abs(y2 - y1);
  el.style.left = `${l}px`;
  el.style.top = `${t}px`;
  el.style.width = `${w}px`;
  el.style.height = `${h}px`;
}

function getMiniShelfIdsInScreenRect(x1, y1, x2, y2) {
  const left = Math.min(x1, x2);
  const right = Math.max(x1, x2);
  const top = Math.min(y1, y2);
  const bottom = Math.max(y1, y2);
  const ids = [];
  document.querySelectorAll(".mini-shelf[data-shelf-id]").forEach((el) => {
    const r = el.getBoundingClientRect();
    if (r.width === 0 && r.height === 0) return;
    const intersect = !(r.right < left || r.left > right || r.bottom < top || r.top > bottom);
    if (intersect) ids.push(el.dataset.shelfId);
  });
  return ids;
}

function onMiniMapDocPointerMove(e) {
  if (!miniMapPointerState) return;
  miniMapPointerState.cx = e.clientX;
  miniMapPointerState.cy = e.clientY;
  const dx = miniMapPointerState.cx - miniMapPointerState.sx;
  const dy = miniMapPointerState.cy - miniMapPointerState.sy;
  if (!miniMapPointerState.dragging && dx * dx + dy * dy > 36) {
    miniMapPointerState.dragging = true;
    const rubber = document.getElementById("miniMapRubberBand");
    rubber.hidden = false;
  }
  if (miniMapPointerState.dragging) {
    positionMiniMapRubberBand(miniMapPointerState.sx, miniMapPointerState.sy, miniMapPointerState.cx, miniMapPointerState.cy);
  }
}

function onMiniMapDocPointerUp() {
  document.removeEventListener("pointermove", onMiniMapDocPointerMove);
  document.removeEventListener("pointerup", onMiniMapDocPointerUp);
  document.removeEventListener("pointercancel", onMiniMapDocPointerUp);
  const rubber = document.getElementById("miniMapRubberBand");
  if (rubber) rubber.hidden = true;
  if (!miniMapPointerState) return;
  const st = miniMapPointerState;
  miniMapPointerState = null;

  if (st.dragging) {
    const ids = getMiniShelfIdsInScreenRect(st.sx, st.sy, st.cx, st.cy);
    if (st.ctrl) {
      ids.forEach((id) => selectedShelfIds.add(id));
    } else {
      selectedShelfIds = new Set(ids);
    }
    updateSelectionUi();
    return;
  }

  if (st.shelfEl) {
    const id = st.shelfEl.dataset.shelfId;
    if (!id) return;
    if (st.ctrl) {
      if (selectedShelfIds.has(id)) selectedShelfIds.delete(id);
      else selectedShelfIds.add(id);
    } else {
      selectedShelfIds = new Set([id]);
    }
    updateSelectionUi();
    smoothScrollTo(`shelf-${id}`);
  }
}

function initMiniMapShelfInteractions() {
  if (miniMapInteractionsInitialized) return;
  miniMapInteractionsInitialized = true;

  const content = document.getElementById("miniMapContent");
  const menu = document.getElementById("miniMapContextMenu");

  content.addEventListener("pointerdown", (e) => {
    hideMiniMapContextMenu();
    if (e.button !== 0) return;
    if (e.target.closest(".mini-region-title")) return;

    miniMapPointerState = {
      sx: e.clientX,
      sy: e.clientY,
      cx: e.clientX,
      cy: e.clientY,
      shelfEl: e.target.closest(".mini-shelf"),
      ctrl: e.ctrlKey || e.metaKey,
      dragging: false
    };
    document.addEventListener("pointermove", onMiniMapDocPointerMove);
    document.addEventListener("pointerup", onMiniMapDocPointerUp);
    document.addEventListener("pointercancel", onMiniMapDocPointerUp);
  });

  content.addEventListener("contextmenu", (e) => {
    if (selectedShelfIds.size === 0) return;
    e.preventDefault();
    if (miniMapViewFieldId === "__none__") {
      alert("请先在「视图字段」中选择一个业务字段，才能批量修改该字段。");
      return;
    }
    showMiniMapContextMenu(e.clientX, e.clientY);
  });

  menu.addEventListener("click", (ev) => {
    ev.stopPropagation();
  });

  document.addEventListener(
    "pointerdown",
    (ev) => {
      if (ev.target.closest("#miniMapContextMenu")) return;
      hideMiniMapContextMenu();
    },
    true
  );

  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") hideMiniMapContextMenu();
  });
}

function showImportReport(report) {
  const dlg = document.getElementById("importReportDialog");
  const title = document.getElementById("importReportTitle");
  const summary = document.getElementById("importReportSummary");
  const issuesEl = document.getElementById("importReportIssues");
  const issuesWrap = document.getElementById("importReportIssuesWrap");

  title.textContent = report.success ? "导入结果校验" : "导入未通过校验";
  summary.className = `import-report-summary ${report.success ? "import-report-ok" : "import-report-error"}`;
  summary.innerHTML = report.summaryHtml;

  issuesEl.innerHTML = "";
  const lines = report.issueLines || [];
  if (lines.length) {
    issuesWrap.style.display = "";
    lines.forEach(({ type, text }) => {
      const li = document.createElement("li");
      li.className = type === "warn" ? "issue-warn" : "issue-skip";
      li.textContent = text;
      issuesEl.appendChild(li);
    });
  } else {
    issuesWrap.style.display = report.success ? "none" : "";
    if (report.success) {
      /* 无明细区块 */
    } else if (!lines.length) {
      const li = document.createElement("li");
      li.className = "issue-skip";
      li.textContent = "无更多明细。";
      issuesEl.appendChild(li);
    }
  }

  dlg.showModal();
}

async function handleImportExcel(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const baseCols = ["区域编号", "货架编号", "层级编号", "货道编号"];

  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) {
      showImportReport({
        success: false,
        summaryHtml: `<strong>工作表为空。</strong><br>文件：<code>${escapeHtml(file.name)}</code>，工作表：<code>${escapeHtml(sheetName)}</code>`,
        issueLines: []
      });
      return;
    }

    const headerKeys = Object.keys(rows[0]);
    const missingHeaders = baseCols.filter((k) => !headerKeys.includes(k));
    if (missingHeaders.length) {
      showImportReport({
        success: false,
        summaryHtml: `<strong>表头缺少必要列，无法导入。</strong><br>缺少：${escapeHtml(missingHeaders.join("、"))}<br>请使用导出的模板列名（区域编号、货架编号、层级编号、货道编号）。`,
        issueLines: [{ type: "skip", text: `当前识别到的列：${headerKeys.slice(0, 20).join("、")}${headerKeys.length > 20 ? "…" : ""}` }]
      });
      return;
    }

    const skipped = [];
    const warnings = [];
    const seenSlots = new Set();
    const validRows = [];

    const parseDim = (raw, colName, excelRow) => {
      const s = String(raw ?? "").trim();
      let n = Number(s);
      if (!Number.isFinite(n)) {
        warnings.push({
          type: "warn",
          text: `第 ${excelRow} 行：${colName}「${s || "(空)"}」无法解析为数字，已按 1 处理`
        });
        n = 1;
      }
      n = Math.floor(n);
      if (n < 1) {
        warnings.push({ type: "warn", text: `第 ${excelRow} 行：${colName} 数值 ${n} 小于 1，已修正为 1` });
        n = 1;
      }
      if (n > 99) {
        warnings.push({ type: "warn", text: `第 ${excelRow} 行：${colName} 数值超过 99，已截断为 99` });
        n = 99;
      }
      return n;
    };

    rows.forEach((r, idx) => {
      const excelRow = idx + 2;
      if (isImportRowEmpty(r)) {
        skipped.push({ type: "skip", text: `第 ${excelRow} 行：空行已忽略` });
        return;
      }
      const missing = baseCols.filter((k) => String(r[k] ?? "").trim() === "");
      if (missing.length) {
        skipped.push({ type: "skip", text: `第 ${excelRow} 行：缺少必填 ${missing.join("、")}，已跳过` });
        return;
      }
      const regionName = String(r["区域编号"]).trim();
      const shelfCode = String(r["货架编号"]).trim();
      if (!regionName || !shelfCode) {
        skipped.push({
          type: "skip",
          text: `第 ${excelRow} 行：区域编号或货架编号为空，已跳过`
        });
        return;
      }
      const level = parseDim(r["层级编号"], "层级编号", excelRow);
      const aisle = parseDim(r["货道编号"], "货道编号", excelRow);
      const slotKey = `${regionName}\t${shelfCode}\t${level}\t${aisle}`;
      if (seenSlots.has(slotKey)) {
        warnings.push({
          type: "warn",
          text: `第 ${excelRow} 行：与文件中已出现的货位重复（${regionName} / ${shelfCode} / ${pad2(level)} / ${pad2(aisle)}），行数据仍参与统计`
        });
      }
      seenSlots.add(slotKey);
      validRows.push(r);
    });

    if (!validRows.length) {
      showImportReport({
        success: false,
        summaryHtml: `<strong>没有可导入的有效数据行。</strong><br>共读取 ${rows.length} 行（不含表头），请检查必填列是否填写完整。`,
        issueLines: [...skipped, ...warnings].slice(0, 200)
      });
      return;
    }

    const customFieldNames = Object.keys(validRows[0]).filter((k) => !baseCols.includes(k));
    const importedFields = customFieldNames.map((name) => {
      const options = [...new Set(validRows.map((row) => String(row[name] || "").trim()).filter(Boolean))];
      return {
        id: uid(),
        name,
        options: options.length ? options : ["默认值"],
        symbols: []
      };
    });
    const fieldIdByName = Object.fromEntries(importedFields.map((f) => [f.name, f.id]));

    const regionMap = new Map();
    validRows.forEach((r) => {
      const regionName = String(r["区域编号"]).trim();
      const shelfCode = String(r["货架编号"]).trim();
      const level = parseDimSilent(r["层级编号"]);
      const aisle = parseDimSilent(r["货道编号"]);
      if (!regionMap.has(regionName)) regionMap.set(regionName, new Map());
      const shelfMap = regionMap.get(regionName);
      if (!shelfMap.has(shelfCode)) {
        shelfMap.set(shelfCode, { maxLevel: 1, maxAisle: 1, rows: [] });
      }
      const slot = shelfMap.get(shelfCode);
      slot.maxLevel = Math.max(slot.maxLevel, level);
      slot.maxAisle = Math.max(slot.maxAisle, aisle);
      slot.rows.push(r);
    });

    const importedRegions = [...regionMap.entries()].map(([regionName, shelfMap]) => {
      const shelves = [...shelfMap.entries()].map(([shelfCode, slot]) => {
        const businessValues = {};
        importedFields.forEach((f) => {
          const counts = new Map();
          slot.rows.forEach((row) => {
            const v = String(row[f.name] || "").trim();
            if (!v) return;
            counts.set(v, (counts.get(v) || 0) + 1);
          });
          const top = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
          businessValues[fieldIdByName[f.name]] = top || f.options[0] || "";
        });
        return {
          id: uid(),
          code: shelfCode,
          rows: String(Math.max(1, Math.min(99, slot.maxLevel))),
          cols: String(Math.max(1, Math.min(99, slot.maxAisle))),
          businessValues
        };
      }).sort((a, b) => naturalCompare(a.code, b.code));

      return {
        id: uid(),
        name: regionName,
        shelves
      };
    }).sort((a, b) => naturalCompare(a.name, b.name));

    state.fields = importedFields;
    state.regions = importedRegions.length ? importedRegions : [{
      id: uid(),
      name: "A",
      shelves: [buildShelf("01")]
    }];
    miniMapViewFieldId = "__none__";
    renderAll();

    const shelfTotal = state.regions.reduce((n, r) => n + r.shelves.length, 0);
    const issueLines = [...skipped, ...warnings];
    const maxShow = 250;
    const truncated = issueLines.length > maxShow;
    const shown = truncated ? issueLines.slice(0, maxShow) : issueLines;
    if (truncated) {
      shown.push({ type: "skip", text: `… 另有 ${issueLines.length - maxShow} 条未显示，请优先修正 Excel 源文件` });
    }

    showImportReport({
      success: true,
      summaryHtml: [
        `<strong>导入成功。</strong>`,
        `文件：<code>${escapeHtml(file.name)}</code>，工作表：<code>${escapeHtml(sheetName)}</code>`,
        `有效行：<strong>${validRows.length}</strong> / 读取行：${rows.length}；区域：<strong>${state.regions.length}</strong>；货架：<strong>${shelfTotal}</strong>`,
        customFieldNames.length ? `识别业务字段：<strong>${customFieldNames.length}</strong> 个` : `未识别到业务字段列（仅四列核心字段）`,
        skipped.length ? `<br><span style="color:#b45309">跳过行：${skipped.length}</span>` : "",
        warnings.length ? `<br><span style="color:#b45309">警告：${warnings.length}</span>` : ""
      ].join("<br>"),
      issueLines: shown
    });
  } catch (err) {
    showImportReport({
      success: false,
      summaryHtml: `<strong>导入失败。</strong><br>${escapeHtml(err?.message || "文件格式错误")}`,
      issueLines: []
    });
  } finally {
    e.target.value = "";
  }
}

function isImportRowEmpty(r) {
  return Object.values(r).every((v) => String(v ?? "").trim() === "");
}

function parseDimSilent(v) {
  let n = Number(String(v ?? "").trim());
  if (!Number.isFinite(n)) n = 1;
  n = Math.floor(n);
  return Math.min(99, Math.max(1, n));
}

function naturalCompare(a, b) {
  return String(a).localeCompare(String(b), "zh-Hans-CN", { numeric: true, sensitivity: "base" });
}
