// =============================================
// CertifyPro - Bulk Certificate Generator
// =============================================

// ---- State ----
const state = {
  templateImg: null,
  templateFile: null,
  csvData: [],
  csvHeaders: [],
  fields: [],
  selectedFieldId: null,
  outputFormat: "both",
  generatedBlobs: [],
  stopGeneration: false,
};

let fieldIdCounter = 0;

// ---- DOM Helpers ----
const qs = id => document.getElementById(id);
const toast = (msg, type="") => {
  const el = qs("toast");
  el.textContent = msg;
  el.className = "toast show" + (type ? " " + type : "");
  clearTimeout(el._t);
  el._t = setTimeout(() => el.className = "toast", 3000);
};

// ---- Step Management ----
let currentStep = 1;

function goToStep(n) {
  qs("section-" + currentStep).classList.remove("active");
  qs("step-indicator-" + currentStep).classList.remove("active");
  for (let i = 1; i < n; i++) {
    qs("step-indicator-" + i).classList.add("completed");
    qs("step-indicator-" + i).classList.remove("active");
  }
  currentStep = n;
  qs("section-" + currentStep).classList.add("active");
  qs("step-indicator-" + currentStep).classList.remove("completed");
  qs("step-indicator-" + currentStep).classList.add("active");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// =============================================
// STEP 1 - Template Upload
// =============================================
const templateDropZone = qs("template-drop-zone");
const templateFileInput = qs("template-file");

function handleTemplateFile(file) {
  if (!file || !file.type.startsWith("image/")) {
    toast("Please upload a PNG or JPG image.", "error"); return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    const img = new Image();
    img.onload = () => {
      state.templateImg = img;
      state.templateFile = file;
      drawTemplatePreview(img);
      qs("template-preview-wrap").classList.remove("hidden");
      templateDropZone.classList.add("hidden");
      qs("step1-next").disabled = false;
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function drawTemplatePreview(img) {
  const canvas = qs("template-canvas");
  const MAX_W = 860;
  const scale = Math.min(1, MAX_W / img.width);
  canvas.width = img.width;
  canvas.height = img.height;
  canvas.style.maxWidth = MAX_W + "px";
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0);
}

templateDropZone.addEventListener("dragover", e => { e.preventDefault(); templateDropZone.classList.add("drag-over"); });
templateDropZone.addEventListener("dragleave", () => templateDropZone.classList.remove("drag-over"));
templateDropZone.addEventListener("drop", e => { e.preventDefault(); templateDropZone.classList.remove("drag-over"); handleTemplateFile(e.dataTransfer.files[0]); });
templateDropZone.addEventListener("click", () => templateFileInput.click());
templateFileInput.addEventListener("change", e => handleTemplateFile(e.target.files[0]));

qs("change-template-btn").addEventListener("click", () => {
  qs("template-preview-wrap").classList.add("hidden");
  templateDropZone.classList.remove("hidden");
  qs("step1-next").disabled = true;
  state.templateImg = null; state.templateFile = null;
  templateFileInput.value = "";
});

qs("step1-next").addEventListener("click", () => goToStep(2));

// =============================================
// STEP 2 - CSV/Excel Upload
// =============================================
const dataDropZone = qs("data-drop-zone");
const dataFileInput = qs("data-file");

function handleDataFile(file) {
  if (!file) return;
  const ext = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(ext)) {
    toast("Please upload a CSV or Excel file.", "error"); return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (json.length < 2) { toast("File must have at least 1 header row and 1 data row.", "error"); return; }
      state.csvHeaders = json[0].map(String);
      state.csvData = json.slice(1).filter(row => row.some(c => c !== "")).map(row => {
        const obj = {};
        state.csvHeaders.forEach((h, i) => obj[h] = String(row[i] !== undefined ? row[i] : ""));
        return obj;
      });
      renderDataTable();
      qs("data-preview-wrap").classList.remove("hidden");
      dataDropZone.classList.add("hidden");
      qs("step2-next").disabled = false;
      qs("row-count-label").textContent = state.csvData.length + " rows";
    } catch(err) {
      toast("Failed to parse file: " + err.message, "error");
    }
  };
  reader.readAsBinaryString(file);
}

function renderDataTable() {
  const table = qs("data-table");
  const previewRows = state.csvData.slice(0, 5);
  let html = "<thead><tr>" + state.csvHeaders.map(h => "<th>" + escHtml(h) + "</th>").join("") + "</tr></thead>";
  html += "<tbody>" + previewRows.map(row => "<tr>" + state.csvHeaders.map(h => "<td>" + escHtml(row[h] || "") + "</td>").join("") + "</tr>").join("") + "</tbody>";
  table.innerHTML = html;
}

dataDropZone.addEventListener("dragover", e => { e.preventDefault(); dataDropZone.classList.add("drag-over"); });
dataDropZone.addEventListener("dragleave", () => dataDropZone.classList.remove("drag-over"));
dataDropZone.addEventListener("drop", e => { e.preventDefault(); dataDropZone.classList.remove("drag-over"); handleDataFile(e.dataTransfer.files[0]); });
dataDropZone.addEventListener("click", () => dataFileInput.click());
dataFileInput.addEventListener("change", e => handleDataFile(e.target.files[0]));

qs("change-data-btn").addEventListener("click", () => {
  qs("data-preview-wrap").classList.add("hidden");
  dataDropZone.classList.remove("hidden");
  qs("step2-next").disabled = true;
  state.csvData = []; state.csvHeaders = [];
  dataFileInput.value = "";
});

qs("step2-back").addEventListener("click", () => goToStep(1));
qs("step2-next").addEventListener("click", () => { initStep3(); goToStep(3); });

// =============================================
// STEP 3 - Field Mapping
// =============================================
function initStep3() {
  drawMainCanvas();
  renderFieldList();
  renderDragFields();
  populateColumnSelects();
  // Auto-detect {{placeholders}} from template - hint fields
  if (state.fields.length === 0) {
    autoAddFields();
  }
}

function autoAddFields() {
  const canvas = qs("main-canvas");
  const total = state.csvHeaders.length;
  const startY = total === 1 ? 44 : 38;
  const step = total <= 1 ? 0 : Math.min(16, 50 / total);
  state.csvHeaders.forEach((header, i) => {
    addField({
      id: ++fieldIdCounter,
      label: header,
      column: header,
      xPct: 50,
      yPct: startY + i * step,
      fontSize: 48,
      fontFamily: "Georgia, serif",
      fontWeight: "bold",
      color: "#1a2b5a",
      align: "center",
      mask: "auto",
      maskWidthPct: 38,
    });
  });
  renderFieldList();
  renderDragFields();
  redrawMainCanvas();
}

function drawMainCanvas() {
  if (!state.templateImg) return;
  const canvas = qs("main-canvas");
  const wrapper = qs("canvas-wrapper");
  const MAX_W = wrapper.clientWidth || 700;
  const scale = Math.min(1, MAX_W / state.templateImg.width);
  canvas.width = state.templateImg.width;
  canvas.height = state.templateImg.height;
  canvas.style.maxWidth = MAX_W + "px";
  canvas.style.width = "100%";
  const ctx = canvas.getContext("2d");
  ctx.drawImage(state.templateImg, 0, 0);
}

function getSampledBgColor(ctx, sampleX, sampleY, imgW, imgH) {
  try {
    const sx = Math.max(2, Math.min(imgW - 3, Math.round(sampleX)));
    const sy = Math.max(2, Math.min(imgH - 3, Math.round(sampleY)));
    const p = ctx.getImageData(sx, sy, 1, 1).data;
    return "rgb(" + p[0] + "," + p[1] + "," + p[2] + ")";
  } catch(e) {
    return "#f5f2e9";
  }
}

function drawSingleFieldOnCanvas(ctx, field, text, imgW, imgH) {
  const x = (field.xPct / 100) * imgW;
  const y = (field.yPct / 100) * imgH;
  ctx.font = field.fontWeight + " " + field.fontSize + "px " + field.fontFamily;

  if (field.mask && field.mask !== "none") {
    const metrics = ctx.measureText(text);
    const textWidth = metrics.width;
    const textHeight = field.fontSize * 1.3;
    const padX = field.fontSize * 0.3;

    const pctWidth = (field.maskWidthPct || 35) / 100 * imgW;
    const maskW = Math.max(pctWidth, textWidth + padX * 2);
    const maskH = textHeight * 1.25;

    let maskX = x;
    if (field.align === "center") maskX = x - maskW / 2;
    else if (field.align === "right") maskX = x - maskW;
    else maskX = x - padX;

    const maskY = y - maskH / 2;

    let bgColor = "#ffffff";
    if (field.mask === "auto") {
      // Sample 35px above text center where background pattern is clean
      bgColor = getSampledBgColor(ctx, x, y - field.fontSize * 1.1, imgW, imgH);
    } else if (field.mask === "white") {
      bgColor = "#ffffff";
    } else if (field.mask === "custom") {
      bgColor = field.maskColor || "#ffffff";
    }

    ctx.fillStyle = bgColor;
    ctx.fillRect(maskX, maskY, maskW, maskH);
  }

  ctx.fillStyle = field.color;
  ctx.textAlign = field.align;
  ctx.textBaseline = "middle";
  ctx.fillText(text, x, y);
}

function redrawMainCanvas() {
  if (!state.templateImg) return;
  const canvas = qs("main-canvas");
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(state.templateImg, 0, 0);

  const sampleRow = state.csvData[0] || {};
  state.fields.forEach(field => {
    const text = field.column ? (sampleRow[field.column] || field.label) : field.label;
    drawSingleFieldOnCanvas(ctx, field, text, canvas.width, canvas.height);
  });
}

function addField(opts = {}) {
  const field = {
    id: opts.id || ++fieldIdCounter,
    label: opts.label || "Field " + fieldIdCounter,
    column: opts.column || (state.csvHeaders[0] || ""),
    xPct: opts.xPct !== undefined ? opts.xPct : 50,
    yPct: opts.yPct !== undefined ? opts.yPct : 50,
    fontSize: opts.fontSize || 42,
    fontFamily: opts.fontFamily || "Georgia, serif",
    fontWeight: opts.fontWeight || "bold",
    color: opts.color || "#1a2b5a",
    align: opts.align || "center",
    mask: opts.mask || "auto",
    maskColor: opts.maskColor || "#ffffff",
    maskWidthPct: opts.maskWidthPct || 35,
  };
  state.fields.push(field);
  return field;
}

function getFieldById(id) { return state.fields.find(f => f.id === id); }

function renderFieldList() {
  const list = qs("field-list");
  list.innerHTML = "";
  state.fields.forEach(field => {
    const el = document.createElement("div");
    el.className = "field-item" + (state.selectedFieldId === field.id ? " selected" : "");
    el.innerHTML = "<div><div class='field-item-name'>" + escHtml(field.label) + "</div><div class='field-item-col'>" + (field.column ? "→ " + escHtml(field.column) : "No column") + "</div></div>";
    el.addEventListener("click", () => selectField(field.id));
    list.appendChild(el);
  });
}

function getTransformForAlign(align) {
  if (align === "left") return "translate(0%, -50%)";
  if (align === "right") return "translate(-100%, -50%)";
  return "translate(-50%, -50%)";
}

function renderDragFields() {
  const layer = qs("drag-fields-layer");
  layer.innerHTML = "";
  const canvas = qs("main-canvas");
  // Use getBoundingClientRect for accurate rendered size (respects CSS scaling)
  const rect = canvas.getBoundingClientRect();
  const cw = rect.width || canvas.offsetWidth || canvas.width;
  const ch = rect.height || canvas.offsetHeight || canvas.height;
  const scaleX = cw / canvas.width;

  state.fields.forEach(field => {
    const el = document.createElement("div");
    el.className = "drag-field-el" + (state.selectedFieldId === field.id ? " selected-field" : "");
    el.setAttribute("data-col", field.column || "");
    el.dataset.fieldId = field.id;
    const px = (field.xPct / 100) * cw;
    const py = (field.yPct / 100) * ch;
    el.style.left = px + "px";
    el.style.top = py + "px";
    el.style.transform = getTransformForAlign(field.align);
    const scaledFontSize = Math.max(8, Math.round(field.fontSize * scaleX));
    el.style.fontSize = scaledFontSize + "px";
    el.style.fontFamily = field.fontFamily;
    el.style.fontWeight = field.fontWeight;
    el.style.color = field.color;
    el.style.textAlign = field.align;
    if (field.mask && field.mask !== "none") {
      const bgColor = field.mask === "white" ? "#ffffff" : (field.maskColor || "#ffffff");
      el.style.background = bgColor;
    }
    const previewVal = field.column ? (state.csvData[0] ? state.csvData[0][field.column] || field.label : field.label) : field.label;
    el.textContent = previewVal;
    el.title = field.column ? ("Column: " + field.column) : "No column bound";
    // Drag logic
    makeDraggable(el, field, canvas, layer);
    el.addEventListener("click", e => { e.stopPropagation(); selectField(field.id); });
    layer.appendChild(el);
  });
  redrawMainCanvas();
}

function makeDraggable(el, field, canvas, layer) {
  let startX, startY, startLeft, startTop, dragging = false;
  el.addEventListener("mousedown", e => {
    if (e.button !== 0) return;
    e.preventDefault();
    dragging = true;
    startX = e.clientX; startY = e.clientY;
    startLeft = parseFloat(el.style.left); startTop = parseFloat(el.style.top);
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
  function onMove(e) {
    if (!dragging) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    const rect = canvas.getBoundingClientRect();
    const cw = rect.width || canvas.offsetWidth;
    const ch = rect.height || canvas.offsetHeight;
    const newLeft = Math.max(0, Math.min(cw, startLeft + dx));
    const newTop = Math.max(0, Math.min(ch, startTop + dy));
    el.style.left = newLeft + "px";
    el.style.top = newTop + "px";
    field.xPct = (newLeft / cw) * 100;
    field.yPct = (newTop / ch) * 100;
    redrawMainCanvas();
  }
  function onUp() {
    dragging = false;
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
  }
  // Touch support
  el.addEventListener("touchstart", e => {
    const touch = e.touches[0];
    startX = touch.clientX; startY = touch.clientY;
    startLeft = parseFloat(el.style.left); startTop = parseFloat(el.style.top);
    dragging = true;
    e.preventDefault();
  }, { passive: false });
  el.addEventListener("touchmove", e => {
    if (!dragging) return;
    const touch = e.touches[0];
    const dx = touch.clientX - startX;
    const dy = touch.clientY - startY;
    const rect = canvas.getBoundingClientRect();
    const cw = rect.width || canvas.offsetWidth;
    const ch = rect.height || canvas.offsetHeight;
    const newLeft = Math.max(0, Math.min(cw, startLeft + dx));
    const newTop = Math.max(0, Math.min(ch, startTop + dy));
    el.style.left = newLeft + "px";
    el.style.top = newTop + "px";
    field.xPct = (newLeft / cw) * 100;
    field.yPct = (newTop / ch) * 100;
    redrawMainCanvas();
    e.preventDefault();
  }, { passive: false });
  el.addEventListener("touchend", () => { dragging = false; });
}

function selectField(id) {
  state.selectedFieldId = id;
  renderFieldList();
  renderDragFields();
  openFieldEditor(id);
}

function populateColumnSelects() {
  const sel = qs("field-column");
  sel.innerHTML = "<option value=''>-- None --</option>" + state.csvHeaders.map(h => "<option value='" + escAttr(h) + "'>" + escHtml(h) + "</option>").join("");
}

function openFieldEditor(id) {
  const field = getFieldById(id);
  if (!field) return;
  qs("field-editor").classList.remove("hidden");
  qs("field-column").value = field.column;
  qs("field-label").value = field.label;
  qs("field-size").value = field.fontSize;
  const rangeEl = qs("field-size-range");
  if (rangeEl) rangeEl.value = Math.min(150, field.fontSize);
  qs("field-font").value = field.fontFamily;
  qs("field-weight").value = field.fontWeight;
  qs("field-color").value = field.color;
  ["left","center","right"].forEach(a => {
    const btn = qs("align-" + a);
    btn.classList.toggle("active", field.align === a);
  });
  qs("field-mask").value = field.mask || "auto";
  const maskColorEl = qs("field-mask-color");
  if (maskColorEl) {
    maskColorEl.value = field.maskColor || "#ffffff";
    maskColorEl.classList.toggle("hidden", field.mask !== "custom");
  }
  const maskWidthRange = qs("field-mask-width-range");
  const maskWidthRow = qs("mask-width-row");
  if (maskWidthRange && maskWidthRow) {
    maskWidthRow.classList.toggle("hidden", field.mask === "none");
    maskWidthRange.value = field.maskWidthPct || 35;
    const valLabel = qs("mask-width-val");
    if (valLabel) valLabel.textContent = (field.maskWidthPct || 35) + "%";
  }
}


function updateSelectedField() {
  const id = state.selectedFieldId;
  const field = getFieldById(id);
  if (!field) return;
  field.column = qs("field-column").value;
  field.label = qs("field-label").value || field.column || "Field";
  field.fontSize = parseInt(qs("field-size").value) || 36;
  field.fontFamily = qs("field-font").value;
  field.fontWeight = qs("field-weight").value;
  field.color = qs("field-color").value;
  field.mask = qs("field-mask").value;
  field.maskColor = qs("field-mask-color").value;
  const maskWidthRange = qs("field-mask-width-range");
  if (maskWidthRange) {
    field.maskWidthPct = parseInt(maskWidthRange.value) || 35;
    const valLabel = qs("mask-width-val");
    if (valLabel) valLabel.textContent = field.maskWidthPct + "%";
  }
  const maskColorEl = qs("field-mask-color");
  if (maskColorEl) {
    maskColorEl.classList.toggle("hidden", field.mask !== "custom");
  }
  const maskWidthRow = qs("mask-width-row");
  if (maskWidthRow) {
    maskWidthRow.classList.toggle("hidden", field.mask === "none");
  }
  renderFieldList();
  renderDragFields();
}

// Editor event listeners
["field-column","field-label","field-size","field-font","field-weight","field-color","field-mask","field-mask-color","field-mask-width-range"].forEach(id => {
  const el = qs(id);
  if (el) {
    el.addEventListener("input", updateSelectedField);
    el.addEventListener("change", updateSelectedField);
  }
});

// Font change async loader for HTML5 Canvas
if (document.fonts) {
  document.fonts.ready.then(() => {
    redrawMainCanvas();
  });
}
qs("field-font").addEventListener("change", async () => {
  const fontVal = qs("field-font").value;
  if (document.fonts) {
    try {
      await document.fonts.load("48px " + fontVal);
    } catch(e) {}
  }
  updateSelectedField();
});

// Sync font size slider <-> number input
const sizeRange = qs("field-size-range");
const sizeNum = qs("field-size");
if (sizeRange) {
  sizeRange.addEventListener("input", () => {
    sizeNum.value = sizeRange.value;
    updateSelectedField();
  });
  sizeNum.addEventListener("input", () => {
    sizeRange.value = Math.min(150, Math.max(8, parseInt(sizeNum.value) || 36));
  });
}


["left","center","right"].forEach(a => {
  qs("align-" + a).addEventListener("click", () => {
    const field = getFieldById(state.selectedFieldId);
    if (field) { field.align = a; }
    ["left","center","right"].forEach(b => qs("align-" + b).classList.toggle("active", b === a));
    renderDragFields();
  });
});

qs("editor-close").addEventListener("click", () => {
  qs("field-editor").classList.add("hidden");
  state.selectedFieldId = null;
  renderFieldList(); renderDragFields();
});

qs("delete-field-btn").addEventListener("click", () => {
  const id = state.selectedFieldId;
  state.fields = state.fields.filter(f => f.id !== id);
  state.selectedFieldId = null;
  qs("field-editor").classList.add("hidden");
  renderFieldList(); renderDragFields();
});

qs("add-field-btn-2").addEventListener("click", () => {
  const canvas = qs("main-canvas");
  const f = addField({ xPct: 50, yPct: 50 });
  state.selectedFieldId = f.id;
  renderFieldList(); renderDragFields(); openFieldEditor(f.id);
});

// Output format
document.querySelectorAll("input[name=outputFormat]").forEach(radio => {
  radio.addEventListener("change", e => { state.outputFormat = e.target.value; });
});

qs("step3-back").addEventListener("click", () => goToStep(2));
qs("step3-next").addEventListener("click", () => {
  if (state.fields.length === 0) { toast("Please add at least one text field.", "error"); return; }
  goToStep(4);
  generateCertificates();
});

// =============================================
// STEP 4 - Generate Certificates
// =============================================
async function generateCertificates() {
  const format = state.outputFormat;
  const total = state.csvData.length;
  const progressBar = qs("progress-bar");
  const progressLabel = qs("progress-label");
  const previewGrid = qs("preview-grid");
  const genCanvas = qs("gen-canvas");
  const ctx = genCanvas.getContext("2d");
  const imgW = state.templateImg.width;
  const imgH = state.templateImg.height;
  genCanvas.width = imgW;
  genCanvas.height = imgH;

  state.stopGeneration = false;
  state.generatedBlobs = [];
  previewGrid.innerHTML = "";
  qs("download-section").classList.add("hidden");

  // Show stop button, hide back button during generation
  qs("stop-gen-btn").classList.remove("hidden");
  qs("gen-back-nav").classList.add("hidden");

  const zip = new JSZip();
  const pdfFolder = format !== "png" ? zip.folder("pdf") : null;
  const pngFolder = format !== "pdf" ? zip.folder("png") : null;

  for (let i = 0; i < total; i++) {
    if (state.stopGeneration) {
      // Stopped early — wrap up what we have
      break;
    }
    const row = state.csvData[i];
    ctx.clearRect(0, 0, imgW, imgH);
    ctx.drawImage(state.templateImg, 0, 0);

    state.fields.forEach(field => {
      const text = field.column ? (row[field.column] || field.label) : field.label;
      const x = (field.xPct / 100) * imgW;
      const y = (field.yPct / 100) * imgH;
      ctx.font = field.fontWeight + " " + field.fontSize + "px " + field.fontFamily;
      ctx.fillStyle = field.color;
      ctx.textAlign = field.align;
      ctx.textBaseline = "middle";
      ctx.fillText(text, x, y);
    });

    let pngDataUrl;
    try {
      pngDataUrl = genCanvas.toDataURL("image/png");
    } catch(e) {
      toast("Canvas security error - make sure images are from the same origin.", "error");
      return;
    }
    const safeName = (row[state.csvHeaders[0]] || "cert_" + (i+1)).replace(/[^a-z0-9_\-]/gi, "_");

    if (format !== "pdf") {
      const pngBlob = await dataUrlToBlob(pngDataUrl);
      pngFolder.file(safeName + ".png", pngBlob);
    }

    if (format !== "png") {
      const { jsPDF } = window.jspdf;
      const orientation = imgW > imgH ? "landscape" : "portrait";
      const pdf = new jsPDF({ orientation, unit: "px", format: [imgW, imgH], compress: true });
      pdf.addImage(pngDataUrl, "PNG", 0, 0, imgW, imgH);
      const pdfBlob = pdf.output("blob");
      pdfFolder.file(safeName + ".pdf", pdfBlob);
    }

    // Add to preview
    const card = document.createElement("div");
    card.className = "preview-card";
    card.innerHTML = "<img src='" + pngDataUrl + "' alt='Certificate'><div class='preview-card-name'>" + escHtml(row[state.csvHeaders[0]] || "#" + (i+1)) + "</div>";
    previewGrid.appendChild(card);

    // Update progress
    const pct = Math.round(((i+1) / total) * 100);
    progressBar.style.width = pct + "%";
    progressLabel.textContent = (i+1) + " / " + total;
    await sleep(0);
  }

  // Wrap up
  const doneCount = previewGrid.querySelectorAll(".preview-card").length;
  const wasStopped = state.stopGeneration;
  qs("stop-gen-btn").classList.add("hidden");
  qs("gen-back-nav").classList.remove("hidden");

  const zipName = (qs("zip-name").value || "certificates").replace(/[^a-z0-9_\-]/gi, "_");

  if (doneCount > 0) {
    const zipBlob = await zip.generateAsync({ type: "blob" });
    state.zipBlob = zipBlob;
    state.zipName = zipName;

    if (wasStopped) {
      qs("gen-subtitle").textContent = "Stopped after " + doneCount + " of " + total + " certificates.";
      qs("download-title").textContent = doneCount + " Certificates Ready (Partial)";
      qs("download-desc").textContent = "Generation was stopped early. " + doneCount + " certificates are packed and ready to download.";
    } else {
      qs("gen-subtitle").textContent = "All " + total + " certificates have been generated!";
      qs("download-title").textContent = total + " Certificates Ready!";
      qs("download-desc").textContent = "All certificates are packed in a ZIP file ready to download.";
    }
    qs("download-section").classList.remove("hidden");
    toast((wasStopped ? "Stopped at " + doneCount : "All " + total) + " certificates done!", wasStopped ? "" : "success");
  } else {
    qs("gen-subtitle").textContent = "No certificates generated.";
    toast("No certificates were generated.", "error");
  }
  state.stopGeneration = false;
}

// Stop generation
qs("stop-gen-btn").addEventListener("click", () => {
  state.stopGeneration = true;
  qs("stop-gen-btn").disabled = true;
  qs("stop-gen-btn").textContent = "Stopping...";
  toast("Stopping after current certificate...", "");
});

qs("download-zip-btn").addEventListener("click", () => {

  if (state.zipBlob) {
    saveAs(state.zipBlob, state.zipName + ".zip");
    toast("Downloading ZIP...", "success");
  }
});

qs("restart-btn").addEventListener("click", () => {
  location.reload();
});

qs("step4-back").addEventListener("click", () => {
  qs("preview-grid").innerHTML = "";
  qs("download-section").classList.add("hidden");
  qs("progress-bar").style.width = "0%";
  qs("progress-label").textContent = "0 / 0";
  qs("gen-subtitle").textContent = "Please wait while we create your certificates...";
  goToStep(3);
  // Re-render canvas and fields
  setTimeout(() => { drawMainCanvas(); renderDragFields(); }, 100);
});

// =============================================
// Utils
// =============================================
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function dataUrlToBlob(dataUrl) {
  return fetch(dataUrl).then(r => r.blob());
}

function escHtml(str) {
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function escAttr(str) {
  return String(str).replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}
