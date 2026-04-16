const state = {
  fields: [],
  regions: []
};
let pendingScrollTargetId = null;
let miniMapViewFieldId = "__none__";
/** 缩略图视图：虚拟字段，与业务字段右键交互一致 */
const MINI_MAP_VIEW_ROWS = "__rows__";
const MINI_MAP_VIEW_COLS = "__cols__";
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

/** 条码/标识在整张纸上的位置（mm，左上角为原点）；宽高由表单输入，此处仅存 x,y。改表单参数不重置位置，仅做边界夹紧。 */
let pdfPrintLayoutState = null;
let pdfEditSelected = null;
/** @type {{ key: string; startClientX: number; startClientY: number; startPos: { x: number; y: number }; paperRect: DOMRect; paperWmm: number; paperHmm: number } | null} */
let pdfEditorDrag = null;
let pdfPreviewRenderRaf = null;
/** 递增以取消过期的异步栅格化，避免旧结果盖住新布局 */
let pdfPreviewRenderGen = 0;
let pdfPreviewResizeObserver = null;
let pdfPreviewResizeDebounceTimer = null;
/** 拖动中节流栅格化，避免每帧重画 PDF 导致闪烁 */
let pdfPreviewDragRasterTimer = null;

function cancelPdfPreviewDragRasterize() {
  if (pdfPreviewDragRasterTimer) {
    clearTimeout(pdfPreviewDragRasterTimer);
    pdfPreviewDragRasterTimer = null;
  }
}

function schedulePdfPreviewDragRasterize(canvasEl, stackEl) {
  if (pdfPreviewDragRasterTimer) return;
  pdfPreviewDragRasterTimer = setTimeout(() => {
    pdfPreviewDragRasterTimer = null;
    schedulePdfPreviewCanvasRefresh(canvasEl, stackEl);
  }, 140);
}

function disposePdfLivePreviewObservers() {
  if (pdfPreviewResizeObserver) {
    pdfPreviewResizeObserver.disconnect();
    pdfPreviewResizeObserver = null;
  }
  if (pdfPreviewResizeDebounceTimer) {
    clearTimeout(pdfPreviewResizeDebounceTimer);
    pdfPreviewResizeDebounceTimer = null;
  }
  cancelPdfPreviewDragRasterize();
}

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
    paperSizeMode: document.getElementById("paperSizeMode").value,
    paperWidth: document.getElementById("paperWidth").value,
    paperHeight: document.getElementById("paperHeight").value,
    barcodeBlockWidthMm: document.getElementById("barcodeBlockWidthMm").value,
    barcodeBlockHeightMm: document.getElementById("barcodeBlockHeightMm").value,
    textBlockWidthMm: document.getElementById("textBlockWidthMm").value,
    textBlockHeightMm: document.getElementById("textBlockHeightMm").value,
    codeFontSize: document.getElementById("codeFontSize").value,
    sepRegionShelf: document.getElementById("sepRegionShelf").value,
    sepShelfLevel: document.getElementById("sepShelfLevel").value,
    sepLevelAisle: document.getElementById("sepLevelAisle").value,
    printLayout: pdfPrintLayoutState
      ? {
          barcode: { x: pdfPrintLayoutState.barcode.x, y: pdfPrintLayoutState.barcode.y },
          text: { x: pdfPrintLayoutState.text.x, y: pdfPrintLayoutState.text.y }
        }
      : null
  };
}

function migratePrintLayoutFromSaved(p) {
  if (p.printLayout && p.printLayout.barcode && p.printLayout.text) {
    return {
      barcode: { x: Number(p.printLayout.barcode.x) || 0, y: Number(p.printLayout.barcode.y) || 0 },
      text: { x: Number(p.printLayout.text.x) || 0, y: Number(p.printLayout.text.y) || 0 }
    };
  }
  if (p.labelLayout && p.labelLayout.barcode && p.labelLayout.text) {
    const pw = Math.max(10, Number(p.paperWidth) || 210);
    const ph = Math.max(10, Number(p.paperHeight) || 297);
    const lw = Math.max(10, Number(p.labelWidth) || 90);
    const lh = Math.max(10, Number(p.labelHeight) || 60);
    const ox = (pw - lw) / 2;
    const oy = (ph - lh) / 2;
    return {
      barcode: {
        x: ox + (Number(p.labelLayout.barcode.x) || 0),
        y: oy + (Number(p.labelLayout.barcode.y) || 0)
      },
      text: {
        x: ox + (Number(p.labelLayout.text.x) || 0),
        y: oy + (Number(p.labelLayout.text.y) || 0)
      }
    };
  }
  return null;
}

function applyPdfFormValues(p) {
  if (!p) return;
  const entries = [
    ["paperSizeMode", p.paperSizeMode],
    ["paperWidth", p.paperWidth],
    ["paperHeight", p.paperHeight],
    ["barcodeBlockWidthMm", p.barcodeBlockWidthMm],
    ["barcodeBlockHeightMm", p.barcodeBlockHeightMm],
    ["textBlockWidthMm", p.textBlockWidthMm],
    ["textBlockHeightMm", p.textBlockHeightMm],
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
  if (p.labelLayout && p.labelLayout.barcode && !p.barcodeBlockWidthMm) {
    const bwEl = document.getElementById("barcodeBlockWidthMm");
    const bhEl = document.getElementById("barcodeBlockHeightMm");
    const twEl = document.getElementById("textBlockWidthMm");
    const thEl = document.getElementById("textBlockHeightMm");
    if (bwEl && p.labelLayout.barcode.w != null) {
      bwEl.value = String(Math.max(8, Math.round(Number(p.labelLayout.barcode.w))));
    }
    if (bhEl && p.labelLayout.barcode.h != null) {
      bhEl.value = String(Math.max(6, Math.round(Number(p.labelLayout.barcode.h))));
    }
    if (twEl && p.labelLayout.text?.w != null) {
      twEl.value = String(Math.max(8, Math.round(Number(p.labelLayout.text.w))));
    }
    if (thEl && p.labelLayout.text?.h != null) {
      thEl.value = String(Math.max(6, Math.round(Number(p.labelLayout.text.h))));
    }
  }
  const migrated = migratePrintLayoutFromSaved(p);
  pdfPrintLayoutState = migrated;
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

  document.getElementById("paperSizeMode").addEventListener("change", (e) => {
    applyPresetSize(e.target.value, "paperWidth", "paperHeight");
    updatePdfPreview();
    schedulePersist();
  });

  const pdfPreviewInputIds = [
    "paperWidth",
    "paperHeight",
    "barcodeBlockWidthMm",
    "barcodeBlockHeightMm",
    "textBlockWidthMm",
    "textBlockHeightMm",
    "codeFontSize",
    "sepRegionShelf",
    "sepShelfLevel",
    "sepLevelAisle"
  ];
  const onPdfFormInput = () => {
    updatePdfPreview();
    schedulePersist();
  };
  pdfPreviewInputIds.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", onPdfFormInput);
    el.addEventListener("change", onPdfFormInput);
  });

  const resetPdfBtn = document.getElementById("resetPdfLayoutBtn");
  if (resetPdfBtn) {
    resetPdfBtn.addEventListener("click", () => {
      const paper = readPaperMmSize();
      const bw = readPdfBarcodeSizeMm();
      const td = readPdfTextBlockSizeMm();
      const def = computeDefaultPrintPositionsMm(paper, bw.w, bw.h, td.w, td.h);
      pdfPrintLayoutState = {
        barcode: { ...def.barcode },
        text: { ...def.text }
      };
      updatePdfPreview();
      schedulePersist();
    });
  }

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
  const paper = readPaperMmSize();
  ensurePdfPrintLayoutState(paper);
  const pdfSettings = readPdfSettingsForExport();

  const pdf = new jsPDF({
    orientation: paper.w >= paper.h ? "landscape" : "portrait",
    unit: "mm",
    format: [paper.w, paper.h]
  });

  for (let i = 0; i < rows.length; i += 1) {
    if (i > 0) pdf.addPage([paper.w, paper.h], paper.w >= paper.h ? "landscape" : "portrait");
    appendLocationLabelPageToDoc(pdf, rows[i], pdfSettings);
  }

  pdf.save(`货位码标签_${nowDateText()}.pdf`);
}

function readPaperMmSize() {
  const mode = document.getElementById("paperSizeMode").value;
  let w = Number(document.getElementById("paperWidth").value);
  let h = Number(document.getElementById("paperHeight").value);
  if (mode === "a4") ({ w, h } = { w: 210, h: 297 });
  if (mode === "a3") ({ w, h } = { w: 297, h: 420 });
  return { w: Math.max(10, w), h: Math.max(10, h) };
}

function readPdfBarcodeSizeMm() {
  const w = Number.parseFloat(document.getElementById("barcodeBlockWidthMm").value);
  const h = Number.parseFloat(document.getElementById("barcodeBlockHeightMm").value);
  return {
    w: Math.min(500, Math.max(8, Number.isFinite(w) ? w : 72)),
    h: Math.min(500, Math.max(6, Number.isFinite(h) ? h : 28))
  };
}

function readPdfTextBlockSizeMm() {
  const w = Number.parseFloat(document.getElementById("textBlockWidthMm").value);
  const h = Number.parseFloat(document.getElementById("textBlockHeightMm").value);
  return {
    w: Math.min(500, Math.max(8, Number.isFinite(w) ? w : 80)),
    h: Math.min(500, Math.max(6, Number.isFinite(h) ? h : 22))
  };
}

function clampPosOnPaper(pos, w, h, pw, ph) {
  return {
    x: Math.max(0, Math.min(pos.x, pw - w)),
    y: Math.max(0, Math.min(pos.y, ph - h))
  };
}

function computeDefaultPrintPositionsMm(paper, bw, bh, tw, th) {
  const { w: pw, h: ph } = paper;
  const gap = 6;
  const bx = (pw - bw) / 2;
  const by = Math.max(8, Math.min(28, (ph - bh - gap - th) * 0.28));
  const tx = (pw - tw) / 2;
  const ty = Math.min(ph - th - 10, by + bh + gap);
  return {
    barcode: clampPosOnPaper({ x: bx, y: by }, bw, bh, pw, ph),
    text: clampPosOnPaper({ x: tx, y: ty }, tw, th, pw, ph)
  };
}

function ensurePdfPrintLayoutState(paper) {
  const bw = readPdfBarcodeSizeMm();
  const td = readPdfTextBlockSizeMm();
  const def = computeDefaultPrintPositionsMm(paper, bw.w, bw.h, td.w, td.h);
  if (!pdfPrintLayoutState) {
    pdfPrintLayoutState = {
      barcode: { ...def.barcode },
      text: { ...def.text }
    };
    return;
  }
  pdfPrintLayoutState.barcode = clampPosOnPaper(pdfPrintLayoutState.barcode, bw.w, bw.h, paper.w, paper.h);
  pdfPrintLayoutState.text = clampPosOnPaper(pdfPrintLayoutState.text, td.w, td.h, paper.w, paper.h);
}

function buildPrintElementRectsMm() {
  const paper = readPaperMmSize();
  ensurePdfPrintLayoutState(paper);
  const bw = readPdfBarcodeSizeMm();
  const td = readPdfTextBlockSizeMm();
  return {
    barcode: {
      x: pdfPrintLayoutState.barcode.x,
      y: pdfPrintLayoutState.barcode.y,
      w: bw.w,
      h: bw.h
    },
    text: {
      x: pdfPrintLayoutState.text.x,
      y: pdfPrintLayoutState.text.y,
      w: td.w,
      h: td.h
    }
  };
}

function readPdfSettingsForExport() {
  const raw = document.getElementById("codeFontSize").value;
  const fontSize = Number.parseFloat(raw);
  const fontSizePt = Number.isFinite(fontSize) ? fontSize : 14;
  const paper = readPaperMmSize();
  ensurePdfPrintLayoutState(paper);
  return {
    fontSizePt,
    separators: {
      rs: document.getElementById("sepRegionShelf").value || "",
      sl: document.getElementById("sepShelfLevel").value || "",
      la: document.getElementById("sepLevelAisle").value || ""
    },
    elementRects: buildPrintElementRectsMm(),
    paperW: paper.w,
    paperH: paper.h
  };
}

function drawPrintPageOnPdf(pdf, paperW, paperH, code, barcodeAsset, settings) {
  const { fontSizePt, separators, elementRects } = settings;
  pdf.setDrawColor(210);
  pdf.setLineWidth(0.15);
  pdf.rect(0.5, 0.5, paperW - 1, paperH - 1);
  const b = elementRects.barcode;
  const t = elementRects.text;
  drawBarcode(pdf, barcodeAsset, b.x, b.y, b.w, b.h);
  drawCodeTextFixedPt(pdf, code, t.x, t.y, t.w, t.h, fontSizePt, separators);
}

/** 单页标签绘制：导出与预览共用，保证所见即所得。 */
function appendLocationLabelPageToDoc(pdf, row, settings) {
  const codeForBarcode = buildLocationCode(row, settings.separators, "barcode");
  const codeForText = buildLocationCode(row, settings.separators, "text");
  const barcodeAsset = makeBarcode(
    codeForBarcode,
    settings.elementRects.barcode.w,
    settings.elementRects.barcode.h
  );
  drawPrintPageOnPdf(pdf, settings.paperW, settings.paperH, codeForText, barcodeAsset, settings);
}

function buildSinglePagePreviewJsPdf() {
  const paper = readPaperMmSize();
  ensurePdfPrintLayoutState(paper);
  const settings = readPdfSettingsForExport();
  const sample = getPreviewSampleRow();
  if (!window.jspdf) return null;
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: paper.w >= paper.h ? "landscape" : "portrait",
    unit: "mm",
    format: [paper.w, paper.h]
  });
  appendLocationLabelPageToDoc(pdf, sample, settings);
  return pdf;
}

function createPreviewPdfArrayBuffer() {
  const pdf = buildSinglePagePreviewJsPdf();
  return pdf ? pdf.output("arraybuffer") : null;
}

function ensurePdfjsWorker() {
  if (typeof pdfjsLib === "undefined") return false;
  if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  }
  return true;
}

/**
 * 用 PDF.js 将导出同源单页栅格化到 canvas，铺满栈区域（避免 iframe 内置查看器留白/缩放导致与拖动层错位）。
 */
async function rasterizePreviewPdfToCanvas(canvasEl, stackEl) {
  const startGen = pdfPreviewRenderGen;
  const ab = createPreviewPdfArrayBuffer();
  if (!ab || !canvasEl || !stackEl) return;
  if (!ensurePdfjsWorker()) {
    console.warn("pdf.js 未加载，标签预览无法栅格化");
    return;
  }
  const rect = stackEl.getBoundingClientRect();
  const w = Math.max(8, rect.width);
  const h = Math.max(8, rect.height);
  let pdfDoc;
  try {
    pdfDoc = await pdfjsLib.getDocument({ data: ab }).promise;
  } catch (err) {
    console.error(err);
    return;
  }
  if (startGen !== pdfPreviewRenderGen) return;
  let page;
  try {
    page = await pdfDoc.getPage(1);
  } catch (err) {
    console.error(err);
    return;
  }
  if (startGen !== pdfPreviewRenderGen) return;
  const vp1 = page.getViewport({ scale: 1 });
  const sCss = Math.min(w / vp1.width, h / vp1.height);
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  const viewport = page.getViewport({ scale: sCss * dpr });
  canvasEl.width = Math.max(1, Math.floor(viewport.width));
  canvasEl.height = Math.max(1, Math.floor(viewport.height));
  canvasEl.style.width = "100%";
  canvasEl.style.height = "100%";
  const ctx = canvasEl.getContext("2d");
  if (!ctx) return;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvasEl.width, canvasEl.height);
  try {
    await page.render({ canvasContext: ctx, viewport }).promise;
  } catch (err) {
    console.error(err);
  }
  if (startGen !== pdfPreviewRenderGen) return;
}

function schedulePdfPreviewCanvasRefresh(canvasEl, stackEl) {
  if (!canvasEl || !stackEl) return;
  if (pdfPreviewRenderRaf) cancelAnimationFrame(pdfPreviewRenderRaf);
  pdfPreviewRenderRaf = requestAnimationFrame(() => {
    pdfPreviewRenderRaf = null;
    rasterizePreviewPdfToCanvas(canvasEl, stackEl).catch((err) => console.error(err));
  });
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

/**
 * 条码位图已由 makeBarcode 裁边并按区块 mm 宽高比铺满，此处直接贴入 mm 矩形，无额外留白。
 * @param {{ dataUrl: string; widthPx: number; heightPx: number }} asset
 */
function drawBarcode(pdf, asset, x, y, w, h) {
  const dataUrl = asset?.dataUrl || asset;
  pdf.addImage(dataUrl, "PNG", x, y, w, h);
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

  const paper = readPaperMmSize();
  ensurePdfPrintLayoutState(paper);

  pdfEditorDrag = null;
  cancelPdfPreviewDragRasterize();
  pdfPreviewRenderGen += 1;
  disposePdfLivePreviewObservers();
  if (pdfPreviewRenderRaf) {
    cancelAnimationFrame(pdfPreviewRenderRaf);
    pdfPreviewRenderRaf = null;
  }
  host.innerHTML = "";
  renderPdfLivePreview(host, paper);
}

function stylePdfEditElOnPaper(el, rectMm, paper) {
  el.style.left = `${(rectMm.x / paper.w) * 100}%`;
  el.style.top = `${(rectMm.y / paper.h) * 100}%`;
  el.style.width = `${(rectMm.w / paper.w) * 100}%`;
  el.style.height = `${(rectMm.h / paper.h) * 100}%`;
}

/**
 * 预览区：内嵌与导出同一套 jsPDF 单页，叠透明拖动层，保证换行、字号、条码与 PDF 完全一致。
 */
function renderPdfLivePreview(host, paper) {
  const wrap = document.createElement("div");
  wrap.className = "pdf-edit-wrap";

  const stage = document.createElement("div");
  stage.className = "pdf-edit-stage";

  const stackEl = document.createElement("div");
  stackEl.className = "pdf-preview-stack";
  stackEl.style.aspectRatio = `${paper.w} / ${paper.h}`;
  stackEl.style.width = "100%";

  const canvas = document.createElement("canvas");
  canvas.className = "pdf-preview-canvas";
  canvas.setAttribute("aria-hidden", "true");

  const layer = document.createElement("div");
  layer.className = "pdf-preview-drag-layer";

  const barEl = document.createElement("div");
  barEl.className = "pdf-edit-el pdf-edit-barcode pdf-edit-overlay-hit";
  barEl.dataset.role = "barcode";
  const barHandle = document.createElement("div");
  barHandle.className = "pdf-edit-drag-handle";
  barHandle.title = "拖动手柄";
  barEl.appendChild(barHandle);

  const textEl = document.createElement("div");
  textEl.className = "pdf-edit-el pdf-edit-text pdf-edit-overlay-hit";
  textEl.dataset.role = "text";
  const textHandle = document.createElement("div");
  textHandle.className = "pdf-edit-drag-handle";
  textHandle.title = "拖动手柄";
  textEl.appendChild(textHandle);

  const applyRectsToDom = (opts = {}) => {
    const refreshPreview = opts.refreshPreview !== false;
    const st = pdfPrintLayoutState;
    if (!st) return;
    const bw = readPdfBarcodeSizeMm();
    const td = readPdfTextBlockSizeMm();
    stylePdfEditElOnPaper(barEl, { ...st.barcode, w: bw.w, h: bw.h }, paper);
    stylePdfEditElOnPaper(textEl, { ...st.text, w: td.w, h: td.h }, paper);
    barEl.classList.toggle("is-selected", pdfEditSelected === "barcode");
    textEl.classList.toggle("is-selected", pdfEditSelected === "text");
    if (refreshPreview) schedulePdfPreviewCanvasRefresh(canvas, stackEl);
  };

  layer.appendChild(barEl);
  layer.appendChild(textEl);
  stackEl.appendChild(canvas);
  stackEl.appendChild(layer);
  stage.appendChild(stackEl);
  wrap.appendChild(stage);
  host.appendChild(wrap);

  stage.addEventListener("pointerdown", (e) => {
    if (!e.target.closest(".pdf-edit-el")) {
      pdfEditSelected = null;
      applyRectsToDom({ refreshPreview: false });
    }
  });

  const beginDrag = (key, e) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    pdfEditSelected = key;
    applyRectsToDom({ refreshPreview: false });
    const r = stackEl.getBoundingClientRect();
    pdfEditorDrag = {
      key,
      startClientX: e.clientX,
      startClientY: e.clientY,
      startPos: { x: pdfPrintLayoutState[key].x, y: pdfPrintLayoutState[key].y },
      paperRect: r,
      paperWmm: paper.w,
      paperHmm: paper.h
    };
    e.currentTarget.setPointerCapture(e.pointerId);
  };

  barEl.addEventListener("pointerdown", (e) => beginDrag("barcode", e));
  textEl.addEventListener("pointerdown", (e) => beginDrag("text", e));

  const onMove = (e) => {
    if (!pdfEditorDrag) return;
    const { key, startClientX, startClientY, startPos, paperRect, paperWmm: pw, paperHmm: ph } = pdfEditorDrag;
    const dims = key === "barcode" ? readPdfBarcodeSizeMm() : readPdfTextBlockSizeMm();
    const prw = paperRect.width || 1;
    const prh = paperRect.height || 1;
    const dxMm = ((e.clientX - startClientX) / prw) * pw;
    const dyMm = ((e.clientY - startClientY) / prh) * ph;
    const next = clampPosOnPaper({ x: startPos.x + dxMm, y: startPos.y + dyMm }, dims.w, dims.h, pw, ph);
    pdfPrintLayoutState[key] = next;
    applyRectsToDom({ refreshPreview: false });
    schedulePdfPreviewDragRasterize(canvas, stackEl);
  };

  const endDrag = (e) => {
    if (pdfEditorDrag) {
      try {
        e.currentTarget.releasePointerCapture(e.pointerId);
      } catch {
        /* ignore */
      }
      pdfEditorDrag = null;
      schedulePersist();
      cancelPdfPreviewDragRasterize();
      schedulePdfPreviewCanvasRefresh(canvas, stackEl);
    }
  };

  barEl.addEventListener("pointermove", onMove);
  barEl.addEventListener("pointerup", endDrag);
  barEl.addEventListener("pointercancel", endDrag);
  textEl.addEventListener("pointermove", onMove);
  textEl.addEventListener("pointerup", endDrag);
  textEl.addEventListener("pointercancel", endDrag);

  applyRectsToDom({ refreshPreview: true });

  disposePdfLivePreviewObservers();
  const ro = new ResizeObserver(() => {
    if (pdfPreviewResizeDebounceTimer) clearTimeout(pdfPreviewResizeDebounceTimer);
    pdfPreviewResizeDebounceTimer = setTimeout(() => {
      pdfPreviewResizeDebounceTimer = null;
      schedulePdfPreviewCanvasRefresh(canvas, stackEl);
    }, 80);
  });
  pdfPreviewResizeObserver = ro;
  ro.observe(stackEl);
}

const makeBarcodeAssetCache = new Map();
const MAKE_BARCODE_CACHE_MAX = 48;

/** 去掉 JsBarcode 生成的透明外边，减小 PDF 内「假留白」。 */
function trimCanvasTransparent(source) {
  const w0 = source.width;
  const h0 = source.height;
  if (!w0 || !h0) return source;
  const ctx = source.getContext("2d");
  if (!ctx) return source;
  const { data } = ctx.getImageData(0, 0, w0, h0);
  let minX = w0;
  let minY = h0;
  let maxX = -1;
  let maxY = -1;
  const stride = 4;
  for (let y = 0; y < h0; y += 1) {
    const row = y * w0 * stride;
    for (let x = 0; x < w0; x += 1) {
      const a = data[row + x * stride + 3];
      if (a > 12) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < minX) return source;
  const cw = maxX - minX + 1;
  const ch = maxY - minY + 1;
  const out = document.createElement("canvas");
  out.width = cw;
  out.height = ch;
  out.getContext("2d").drawImage(source, minX, minY, cw, ch, 0, 0, cw, ch);
  return out;
}

/**
 * 将裁好的条码按区块 mm 的宽高比做「铺满」合成（必要时左右或上下裁掉多余条宽），
 * 使 drawBarcode 用 (w,h) mm 贴图时四周无应用层留白。
 */
function composeBarcodeToBoxAspect(source, boxWmm, boxHmm) {
  const trimmed = trimCanvasTransparent(source);
  const tw = trimmed.width;
  const th = trimmed.height;
  if (!tw || !th) return trimmed;
  const targetAr = Math.max(0.02, boxWmm / boxHmm);
  const srcAr = tw / th;
  const longPx = 1600;
  let outW;
  let outH;
  if (targetAr >= 1) {
    outW = longPx;
    outH = Math.max(1, Math.round(longPx / targetAr));
  } else {
    outH = longPx;
    outW = Math.max(1, Math.round(longPx * targetAr));
  }
  const out = document.createElement("canvas");
  out.width = outW;
  out.height = outH;
  const octx = out.getContext("2d");
  if (!octx) return trimmed;
  octx.fillStyle = "#ffffff";
  octx.fillRect(0, 0, outW, outH);
  let sx;
  let sy;
  let sw;
  let sh;
  if (srcAr > targetAr) {
    sh = th;
    sw = th * targetAr;
    sx = (tw - sw) / 2;
    sy = 0;
  } else {
    sw = tw;
    sh = tw / targetAr;
    sx = 0;
    sy = (th - sh) / 2;
  }
  octx.drawImage(trimmed, sx, sy, sw, sh, 0, 0, outW, outH);
  return out;
}

/**
 * 生成与条码区块 mm 同宽高比的 PNG，供 PDF 整格贴入；宽窄条比例在生成阶段保持，仅裁掉库自带透明边。
 */
function makeBarcode(code, boxWmm, boxHmm) {
  const W = Math.max(8, Number(boxWmm) || 72);
  const H = Math.max(6, Number(boxHmm) || 28);
  const cacheKey = `${code}\t${Math.round(W * 4) / 4}\t${Math.round(H * 4) / 4}`;
  if (makeBarcodeAssetCache.has(cacheKey)) {
    return makeBarcodeAssetCache.get(cacheKey);
  }

  const targetAr = W / H;
  let bestCanvas = null;
  let bestScore = -1;

  for (let lineW = 1; lineW <= 8; lineW += 1) {
    for (let hPx = 24; hPx <= 240; hPx += 4) {
      const canvas = document.createElement("canvas");
      try {
        JsBarcode(canvas, code, {
          format: "CODE128",
          lineColor: "#000",
          width: lineW,
          height: hPx,
          displayValue: false,
          margin: 0
        });
      } catch {
        continue;
      }
      const cw = canvas.width;
      const ch = canvas.height;
      if (!cw || !ch) continue;
      const ar = cw / ch;
      const fit = Math.min(targetAr / ar, ar / targetAr);
      const score = fit + cw * ch * 1e-9;
      if (score > bestScore) {
        bestScore = score;
        bestCanvas = canvas;
      }
    }
  }

  if (!bestCanvas) {
    const c = document.createElement("canvas");
    JsBarcode(c, code, {
      format: "CODE128",
      lineColor: "#000",
      width: 2,
      height: 80,
      displayValue: false,
      margin: 0
    });
    bestCanvas = c;
  }

  const composed = composeBarcodeToBoxAspect(bestCanvas, W, H);
  const out = {
    dataUrl: composed.toDataURL("image/png"),
    widthPx: composed.width,
    heightPx: composed.height
  };
  if (makeBarcodeAssetCache.size >= MAKE_BARCODE_CACHE_MAX) {
    const firstKey = makeBarcodeAssetCache.keys().next().value;
    makeBarcodeAssetCache.delete(firstKey);
  }
  makeBarcodeAssetCache.set(cacheKey, out);
  return out;
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

function getShelfDimColor(valueStr) {
  const n = Number(String(valueStr).replace(/\D/g, "")) || 0;
  const palette = ["#93c5fd", "#86efac", "#fde68a", "#fca5a5", "#c4b5fd", "#67e8f9", "#f9a8d4", "#d9f99d"];
  return palette[n % palette.length];
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
      const rowN = Number(sanitizeTwoDigitNum(shelf.rows, 7));
      const colN = Number(sanitizeTwoDigitNum(shelf.cols, 9));
      const slotCount = rowN * colN;
      let bottomText = "";
      let bg = "#4b5563";
      if (miniMapViewFieldId === MINI_MAP_VIEW_ROWS) {
        bottomText = `${shelf.rows} 行`;
        bg = getShelfDimColor(shelf.rows);
      } else if (miniMapViewFieldId === MINI_MAP_VIEW_COLS) {
        bottomText = `${shelf.cols} 列`;
        bg = getShelfDimColor(shelf.cols);
      } else if (miniMapViewFieldId === "__none__") {
        bottomText = "未选视图";
        bg = "#374151";
      } else if (viewField) {
        bottomText = shelf.businessValues[viewField.id] || "-";
        bg = getShelfColor(viewField, bottomText);
      } else {
        bottomText = "-";
      }
      shelfNode.style.background = bg;
      shelfNode.innerHTML = `
        <div class="mini-shelf-top">${escapeHtml(shelf.code)}</div>
        <div class="mini-shelf-count">${slotCount}货位</div>
        <div class="mini-shelf-bottom">${escapeHtml(bottomText)}</div>
      `;
      shelfNode.title = `区域${region.name} 货架${shelf.code}（左键选中，拖拽框选，Ctrl/Cmd+点击加选；有选中时右键按当前视图字段批量修改）`;
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

  const rowsOp = document.createElement("option");
  rowsOp.value = MINI_MAP_VIEW_ROWS;
  rowsOp.textContent = "行数（层级）";
  select.appendChild(rowsOp);

  const colsOp = document.createElement("option");
  colsOp.value = MINI_MAP_VIEW_COLS;
  colsOp.textContent = "列数（货道）";
  select.appendChild(colsOp);

  state.fields.forEach((field) => {
    const op = document.createElement("option");
    op.value = field.id;
    op.textContent = field.name;
    select.appendChild(op);
  });

  const availableValues = new Set([
    "__none__",
    MINI_MAP_VIEW_ROWS,
    MINI_MAP_VIEW_COLS,
    ...state.fields.map((f) => f.id)
  ]);
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

function applyRowsToSelectedShelves(rowsStr) {
  const rows = sanitizeTwoDigitNum(rowsStr, 7);
  state.regions.forEach((region) => {
    region.shelves.forEach((shelf) => {
      if (selectedShelfIds.has(shelf.id)) shelf.rows = rows;
    });
  });
  renderAll();
}

function applyColsToSelectedShelves(colsStr) {
  const cols = sanitizeTwoDigitNum(colsStr, 9);
  state.regions.forEach((region) => {
    region.shelves.forEach((shelf) => {
      if (selectedShelfIds.has(shelf.id)) shelf.cols = cols;
    });
  });
  renderAll();
}

function appendMiniMapContextNumericBatch(menu, mode) {
  const wrap = document.createElement("div");
  wrap.className = "context-menu-numeric";
  const lab = document.createElement("label");
  lab.className = "context-menu-numeric-label";
  lab.textContent = mode === "row" ? "自定义行数" : "自定义列数";
  const inp = document.createElement("input");
  inp.type = "number";
  inp.min = 1;
  inp.max = 99;
  inp.className = "context-menu-numeric-input";
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "context-menu-numeric-btn";
  btn.textContent = "应用";
  btn.addEventListener("click", () => {
    if (mode === "row") applyRowsToSelectedShelves(inp.value);
    else applyColsToSelectedShelves(inp.value);
    hideMiniMapContextMenu();
  });
  lab.appendChild(inp);
  wrap.appendChild(lab);
  wrap.appendChild(btn);
  menu.appendChild(wrap);
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
  const n = selectedShelfIds.size;
  menu.innerHTML = "";

  if (miniMapViewFieldId === MINI_MAP_VIEW_ROWS) {
    const title = document.createElement("div");
    title.className = "context-menu-title";
    title.textContent = `批量设置行数 · 已选 ${n} 个货架`;
    menu.appendChild(title);
    [5, 6, 7, 8, 9, 10, 12, 15].forEach((v) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = `设为 ${v} 行`;
      btn.addEventListener("click", () => {
        applyRowsToSelectedShelves(String(v));
        hideMiniMapContextMenu();
      });
      menu.appendChild(btn);
    });
    appendMiniMapContextNumericBatch(menu, "row");
  } else if (miniMapViewFieldId === MINI_MAP_VIEW_COLS) {
    const title = document.createElement("div");
    title.className = "context-menu-title";
    title.textContent = `批量设置列数 · 已选 ${n} 个货架`;
    menu.appendChild(title);
    [6, 7, 8, 9, 10, 12, 15, 20].forEach((v) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = `设为 ${v} 列`;
      btn.addEventListener("click", () => {
        applyColsToSelectedShelves(String(v));
        hideMiniMapContextMenu();
      });
      menu.appendChild(btn);
    });
    appendMiniMapContextNumericBatch(menu, "col");
  } else if (field && miniMapViewFieldId !== "__none__") {
    const title = document.createElement("div");
    title.className = "context-menu-title";
    title.textContent = `批量设置「${field.name}」· 已选 ${n} 个货架`;
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
  } else {
    const hint = document.createElement("div");
    hint.className = "context-menu-hint";
    hint.textContent = "请先在「视图字段」中选择：行数、列数或某个业务字段，再右键批量修改。";
    menu.appendChild(hint);
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
