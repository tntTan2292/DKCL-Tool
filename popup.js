const API_URLS = {
  phatThanhCongBuuCuc: "https://dkcl.vnpost.vn/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc",
  thuGomLienTinh: "https://dkcl.vnpost.vn/kpi/chat-luong-thu-gom-cua-nvtg",
  noiTinhF11: "https://dkcl.vnpost.vn/kpi/chat-luong-toan-trinh-buu-gui-noi-tinh",
  thuGomLienTinhF12: "https://dkcl.vnpost.vn/kpi/chat-luong-thu-gom-buu-lien-tinh",
  phatLienTinhF13: "https://dkcl.vnpost.vn/kpi/chat-luong-phat-buu-gui-lien-tinh"
};

const EXCEL_COLUMNS = [
  "STT",
  "Mã tỉnh",
  "Tên tỉnh",
  "Sản lượng PTC/ Nộp tiền/ CH",
  "Sản lượng bưu gửi PTC tại bưu cục trong thời gian 8 tiếng (có quét TMS)",
  "Tỷ lệ TMĐT",
  "TT TMĐT",
  "Sản lượng PTC/ Nộp tiền/ CH",
  "Sản lượng bưu gửi PTC tại bưu cục trong thời gian 8 tiếng (có quét TMS)",
  "Tỷ lệ Truyền thống",
  "TT Truyền thống",
  "Sản lượng thu gom thành công",
  "Sản lượng thu gom thành công đúng QĐ (≤6 giờ)",
  "Tỷ lệ Thu gom",
  "TT Thu gom",
  "Tỷ lệ bình quân đạt được",
  "TT bình quân"
];

const TEMPLATE_EXCEL_COLUMNS = [
  "STT",
  "Mã tỉnh chấp nhận",
  "Tên tỉnh chấp nhận",
  "Sản lượng PTC/ Nộp tiền/ CH",
  "Sản lượng bưu gửi PTC tại bưu cục trong thời gian 6 tiếng (có quét TMS)",
  "Tỷ lệ gửi PTC tại bưu cục trong thời gian 6 tiếng (có quét TMS)",
  "TT mạng lưới",
  "Sản lượng PTC/ Nộp tiền/ CH",
  "Sản lượng bưu gửi PTC tại bưu cục trong thời gian 6 tiếng (có quét TMS)",
  "Tỷ lệ gửi PTC tại bưu cục trong thời gian 6 tiếng (có quét TMS)",
  "TT mạng lưới",
  "Sản lượng thu gom thành công",
  "Sản lượng thu gom thành công đúng QĐ (≤6 giờ)",
  "Tỷ lệ TG bưu gửi đúng thời gian QĐ (≤6 giờ)",
  "TT mạng lưới",
  "Tỷ lệ BQ",
  "TT"
];

const V2_EXCEL_COLUMNS = [
  ...TEMPLATE_EXCEL_COLUMNS,
  "Sản lượng PTC/ Nộp tiền/ CH",
  "Sản lượng bưu gửi PTC tại bưu cục trong thời gian 8 tiếng (có quét TMS)",
  "Tỷ lệ gửi PTC tại bưu cục trong thời gian 8 tiếng (có quét TMS)",
  "TT mạng lưới",
  "SL bưu gửi phát thành công/Nộp tiền/chuyển hoàn",
  "Sản lượng bưu gửi PTC/nộp tiền",
  "Sản lượng bưu gửi PTC/nộp tiền đúng QĐ <= 24h",
  "Tỷ lệ bưu gửi PTC/Nộp tiền đúng QĐ",
  "TT mạng lưới",
  "SL bưu gửi có BĐ10 đóng đi tại khai thác tỉnh chấp nhận",
  "SL bưu gửi có BD10 quét TMS lên tại Khai thác tỉnh hoặc BC khớp nối",
  "SL đúng thời gian quy định <= 10 giờ",
  "Tỷ lệ đúng thời gian quy định",
  "TT mạng lưới",
  "SL bưu gửi phát thành công/Nộp tiền/CH",
  "Sản lượng bưu gửi PTC/nộp tiền đúng thời gian QĐ 2026",
  "Tỷ lệ bưu gửi PTC/Nộp tiền đúng QĐ theo chi tiêu 2026",
  "TT mạng lưới",
  "Tỷ lệ BQ( TGLT+ PTC LT)",
  "TT"
];

const TEMPLATE_GROUP_HEADERS = [
  { label: "", colspan: 3 },
  { label: "Phát thành công TMĐT", colspan: 4 },
  { label: "Phát thành công Truyền thống", colspan: 4 },
  { label: "Thu gom bưu gửi", colspan: 4 },
  { label: "Tỷ lệ trung bình chung", colspan: 2 }
];

const V2_GROUP_HEADERS = [
  ...TEMPLATE_GROUP_HEADERS,
  { label: "Chất lượng phát thành công tại bưu cục (F4.1)", colspan: 4 },
  { label: "Nội tỉnh F1.1", colspan: 5 },
  { label: "Thu gom bưu gửi đi liên tỉnh (F1.2)", colspan: 5 },
  { label: "Chất lượng phát bưu gửi liên tỉnh (F1.3)", colspan: 4 },
  { label: "Tỷ lệ trung bình chung chất lượng (TGLT,PTC LT)", colspan: 2 }
];

const GROUP_HEADERS = [
  { label: "", colspan: 3 },
  { label: "Phát thành công TMĐT", colspan: 4 },
  { label: "Phát thành công Truyền thống", colspan: 4 },
  { label: "Thu gom", colspan: 4 },
  { label: "Bình quân", colspan: 2 }
];

const SOURCE_CONFIGS = [
  {
    key: "phatThanhCongBuuCuc",
    name: "Phát thành công TMĐT",
    status: "Đang gọi curl Phát thành công TMĐT...",
    buildParams: buildPhatThanhCongTmdtParams,
    columns: [
      { source: "total", target: 3 },
      { source: "success8h", target: 4 },
      { source: "rate", target: 5, type: "percent" },
      { source: "rank", target: 6 }
    ]
  },
  {
    key: "phatThanhCongBuuCuc",
    name: "Phát thành công Truyền thống",
    status: "Đang gọi curl Phát thành công Truyền thống...",
    buildParams: buildPhatThanhCongTruyenThongParams,
    columns: [
      { source: "total", target: 7 },
      { source: "success8h", target: 8 },
      { source: "rate", target: 9, type: "percent" },
      { source: "rank", target: 10 }
    ]
  },
  {
    key: "thuGomLienTinh",
    name: "Thu gom bưu gửi đi liên tỉnh",
    status: "Đang gọi curl Thu gom bưu gửi đi liên tỉnh...",
    buildParams: buildThuGomParams,
    columns: [
      { source: "total", target: 11 },
      { source: "onTime", target: 12 }
    ]
  }
];

const V2_SOURCE_CONFIGS = [
  { key: "phatThanhCongBuuCuc", name: "Chất lượng phát thành công tại bưu cục (F4.1)", status: "Đang gọi curl V2 F4.1...", buildParams: buildPhatThanhCongV2Params, values: { total: 10, onTime: 25 }, targets: { total: 17, onTime: 18, rate: 19, rank: 20 } },
  { key: "noiTinhF11", name: "Nội tỉnh F1.1", status: "Đang gọi curl V2 Nội tỉnh F1.1...", buildParams: buildNoiTinhF11Params, codeIndex: 1, nameIndex: 2, values: { total: 10, subTotal: 11, onTime: 12 }, targets: { total: 21, subTotal: 22, onTime: 23, rate: 24, rank: 25 } },
  { key: "thuGomLienTinhF12", name: "Thu gom bưu gửi đi liên tỉnh (F1.2)", status: "Đang gọi curl V2 Thu gom liên tỉnh F1.2...", buildParams: buildThuGomLienTinhF12Params, codeIndex: 1, nameIndex: 2, values: { total: 5, subTotal: 6, onTime: 7 }, targets: { total: 26, subTotal: 27, onTime: 28, rate: 29, rank: 30 } },
  { key: "phatLienTinhF13", name: "Chất lượng phát bưu gửi liên tỉnh (F1.3)", status: "Đang gọi curl V2 Phát liên tỉnh F1.3...", buildParams: buildPhatLienTinhF13Params, codeIndex: 1, nameIndex: 2, values: { total: 8, onTime: 16 }, targets: { total: 31, onTime: 32, rate: 33, rank: 34 } }
];

const PROVINCE_CODE_INDEX = 1;
const PROVINCE_NAME_INDEX = 2;
const form = document.getElementById("report-form");
const fromDateInput = document.getElementById("from-date");
const toDateInput = document.getElementById("to-date");
const nextFromDateInput = document.getElementById("next-from-date");
const nextToDateInput = document.getElementById("next-to-date");
const dateDisplayMap = new Map([
  [fromDateInput, document.getElementById("from-date-display")],
  [toDateInput, document.getElementById("to-date-display")],
  [nextFromDateInput, document.getElementById("next-from-date-display")],
  [nextToDateInput, document.getElementById("next-to-date-display")]
]);
const displayDateMap = new Map([...dateDisplayMap.entries()].map(([nativeInput, displayInput]) => [displayInput, nativeInput]));
const statusBox = document.getElementById("status");
const reportPanel = document.getElementById("report-panel");
const reportBadge = document.getElementById("report-badge");
const emptyState = document.getElementById("empty-state");
const reportSummary = document.getElementById("report-summary");
const reportText = document.getElementById("report-text");
const weightTmdtInput = document.getElementById("weight-tmdt");
const weightTraditionalInput = document.getElementById("weight-traditional");
const weightPickupInput = document.getElementById("weight-pickup");
const templateReportButton = document.getElementById("export-province-template");
const v2Buttons = [...document.querySelectorAll("button[data-v2-group]")];
const buttons = [...document.querySelectorAll("button[data-group]")];
const datePickerButtons = [...document.querySelectorAll(".date-picker-btn")];
const allActionButtons = templateReportButton ? [...buttons, templateReportButton, ...v2Buttons] : [...buttons, ...v2Buttons];
const pinUiButton = document.getElementById("pin-ui");
const PIN_STORAGE_KEY = "dkclSidebarPinned";

const modeDateRadio = document.getElementById("mode-date");
const modeMonthRadio = document.getElementById("mode-month");
const dateModeContainer = document.getElementById("date-mode-container");
const monthModeContainer = document.getElementById("month-mode-container");
const modeOptionDate = document.getElementById("mode-option-date");
const modeOptionMonth = document.getElementById("mode-option-month");

function updateReportModeUi() {
  const isMonthMode = modeMonthRadio && modeMonthRadio.checked;
  if (dateModeContainer) dateModeContainer.hidden = isMonthMode;
  if (monthModeContainer) monthModeContainer.hidden = !isMonthMode;
  if (modeOptionDate) modeOptionDate.classList.toggle("active", !isMonthMode);
  if (modeOptionMonth) modeOptionMonth.classList.toggle("active", isMonthMode);
}

if (modeDateRadio && modeMonthRadio) {
  modeDateRadio.addEventListener("change", updateReportModeUi);
  modeMonthRadio.addEventListener("change", updateReportModeUi);
  updateReportModeUi();
}

let pendingExportAction = null;

const metricSelectionPanel = document.getElementById("metric-selection-panel");
const selectedReportTitle = document.getElementById("selected-report-title");
const metricsCheckboxList = document.getElementById("metrics-checkbox-list");
const tickAllMetrics = document.getElementById("tick-all-metrics");
const confirmExportBtn = document.getElementById("confirm-export-btn");

const AVAILABLE_METRICS = [
  "F4.1 – Chất lượng phát thành công tại bưu cục",
  "F1.3 – Chất lượng phát bưu gửi liên tỉnh",
  "F1.1 – Nội tỉnh",
  "F1.2 – Thu gom bưu gửi đi liên tỉnh"
];

function populateMetricsList() {
  if (!metricsCheckboxList) return;
  metricsCheckboxList.innerHTML = AVAILABLE_METRICS.map((m) => `
    <label class="metric-chk-item" style="display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 500; color: #e2e8f0; cursor: pointer; padding: 6px 10px; background: rgba(255,255,255,0.05); border-radius: 6px;">
      <input type="checkbox" class="metric-item-chk" value="${escapeHtml(m)}" checked />
      <span>${escapeHtml(m)}</span>
    </label>
  `).join("");
}

if (tickAllMetrics) {
  tickAllMetrics.addEventListener("change", () => {
    const chks = document.querySelectorAll(".metric-item-chk");
    chks.forEach((c) => (c.checked = tickAllMetrics.checked));
  });
}

function openMetricSelectionPanel(v2Group, v2Mode) {
  pendingExportAction = { v2Group, v2Mode };
  populateMetricsList();
  if (tickAllMetrics) tickAllMetrics.checked = true;

  const isMonthMode = modeMonthRadio && modeMonthRadio.checked;
  const groupLabel = v2Group === "BC" ? "Bưu cục" : "34 Tỉnh/TP";
  const modeLabel = isMonthMode ? "Lũy kế đa tháng" : "So sánh 2 kỳ (Ngày)";

  if (selectedReportTitle) {
    selectedReportTitle.textContent = `Chọn chỉ tiêu xuất Báo cáo V2 ${groupLabel} (${modeLabel})`;
  }

  if (metricSelectionPanel) {
    metricSelectionPanel.hidden = false;
    metricSelectionPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

if (confirmExportBtn) {
  confirmExportBtn.addEventListener("click", () => {
    const selectedLabels = getSelectedMetricLabels();
    if (!selectedLabels || selectedLabels.length === 0) {
      setStatus("Vui lòng chọn ít nhất 1 chỉ tiêu chất lượng.", "error");
      return;
    }

    if (!pendingExportAction) {
      pendingExportAction = { v2Group: "TINH", v2Mode: "compare" };
    }

    const isMonthMode = modeMonthRadio && modeMonthRadio.checked;
    if (isMonthMode) {
      exportReportV2MultiMonth(pendingExportAction.v2Group);
    } else {
      exportReportV2(pendingExportAction.v2Group, pendingExportAction.v2Mode || "compare");
    }
  });
}

initDefaultDates();
buttons.forEach((button) => button.addEventListener("click", () => exportReport(button.dataset.group)));
v2Buttons.forEach((button) => {
  button.addEventListener("click", () => {
    openMetricSelectionPanel(button.dataset.v2Group, button.dataset.v2Mode || "compare");
  });
});
if (templateReportButton) templateReportButton.addEventListener("click", exportProvinceTemplateReport);
if (pinUiButton) {
  initPinnedUi();
  pinUiButton.addEventListener("click", togglePinnedUi);
}
[fromDateInput, toDateInput, nextFromDateInput, nextToDateInput].forEach((input) => {
  input.addEventListener("change", () => syncDateDisplay(input));
  input.addEventListener("input", () => syncDateDisplay(input));
});
displayDateMap.forEach((nativeInput, displayInput) => {
  displayInput.addEventListener("input", () => syncNativeDateFromDisplay(displayInput));
  displayInput.addEventListener("change", () => syncNativeDateFromDisplay(displayInput, { showError: true }));
  displayInput.addEventListener("blur", () => syncNativeDateFromDisplay(displayInput, { showError: true, normalize: true }));
});
datePickerButtons.forEach((button) => {
  button.addEventListener("click", (event) => {
    event.preventDefault();
    const nativeInput = button.parentElement?.querySelector(".date-native");
    if (!nativeInput) return;
    nativeInput.focus({ preventScroll: true });
    if (typeof nativeInput.showPicker === "function") nativeInput.showPicker();
    else nativeInput.click();
  });
});

function initPinnedUi() {
  const pinned = localStorage.getItem(PIN_STORAGE_KEY) === "true";
  updatePinUiButton(pinned);
  notifyParentPinnedState(pinned);
}

function togglePinnedUi() {
  const pinned = pinUiButton.getAttribute("aria-pressed") !== "true";
  localStorage.setItem(PIN_STORAGE_KEY, String(pinned));
  updatePinUiButton(pinned);
  notifyParentPinnedState(pinned);
  setStatus(pinned ? "Đã ghim giao diện: sidebar sẽ luôn mở." : "Đã bỏ ghim giao diện: có thể ẩn/hiện sidebar bằng nút DKCL.", "ok");
}

function updatePinUiButton(pinned) {
  pinUiButton.classList.toggle("is-pinned", pinned);
  pinUiButton.setAttribute("aria-pressed", String(pinned));
  pinUiButton.title = pinned ? "Bỏ ghim giao diện" : "Ghim giao diện luôn mở";
  const text = pinUiButton.querySelector(".pin-text");
  if (text) text.textContent = pinned ? "Đã ghim" : "Ghim";
}

function notifyParentPinnedState(pinned) {
  if (window.parent === window) return;
  window.parent.postMessage({ source: "dkcl-report-popup", type: "pin-state", pinned }, "*");
}

function initDefaultDates() {
  const today = new Date();
  const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  const nextFirstDay = addDays(today, 1);
  const nextLastDay = new Date(nextFirstDay.getFullYear(), nextFirstDay.getMonth() + 1, 0);

  fromDateInput.value = toInputDate(firstDay);
  toDateInput.value = toInputDate(today);
  nextFromDateInput.value = toInputDate(nextFirstDay);
  nextToDateInput.value = toInputDate(nextLastDay);
  syncAllDateDisplays();

  const fromMonthInput = document.getElementById("from-month");
  const toMonthInput = document.getElementById("to-month");
  if (fromMonthInput && !fromMonthInput.value) {
    const currentYear = today.getFullYear();
    fromMonthInput.value = `${currentYear}-01`;
  }
  if (toMonthInput && !toMonthInput.value) {
    const currentYear = today.getFullYear();
    const currentMonthStr = pad(today.getMonth() + 1);
    toMonthInput.value = `${currentYear}-${currentMonthStr}`;
  }
}

async function exportReport(tuyChonGR) {
  if (!form.reportValidity()) return;

  const from = parseControlDate(fromDateInput.value);
  const to = parseControlDate(toDateInput.value);
  const nextFrom = parseControlDate(nextFromDateInput.value);
  const nextTo = parseControlDate(nextToDateInput.value);
  if (!validateRequiredDates(from, to, nextFrom, nextTo) || !validatePeriodDates(from, to, nextFrom, nextTo)) return;

  setBusy(true);
  setStatus(`Đang lấy dữ liệu tổng hợp theo đúng mẫu filebaocao.xlsx (${tuyChonGR})...`);

  try {
    const rows = await fetchCombinedReport(tuyChonGR, from, to);
    if (!rows.length) {
      throw new Error("Không tìm thấy dữ liệu trong các curl. Vui lòng kiểm tra đăng nhập hoặc khoảng ngày.");
    }

    const weights = readWeights();
    const finalRows = finalizeRows(rows, weights, tuyChonGR);
    console.log(`[DKCL][${tuyChonGR}] joined rows`, rows.length, "final rows", finalRows.length, finalRows.slice(0, 5));
    if (!finalRows.length) {
      throw new Error("Đã join được dữ liệu nhưng tất cả dòng bị lọc vì tỷ lệ bình quân bằng 0. Kiểm tra lại index cột sản lượng trong log sample cells.");
    }
    const fileName = `filebaocao_${tuyChonGR}_${toFileDate(from)}_${toFileDate(to)}.xls`;
    renderReportPreview({ rows: finalRows, tuyChonGR, from, to, fileName, weights });
    downloadExcel({ rows: finalRows, tuyChonGR, from, to, fileName });
    setStatus(`Hoàn tất: đã gọi ${SOURCE_CONFIGS.length} curl, hiển thị ${finalRows.length} dòng và xuất ${fileName}.`, "ok");
  } catch (error) {
    setStatus(error.message || "Có lỗi khi lấy dữ liệu.", "error");
  } finally {
    setBusy(false);
  }
}

async function exportProvinceTemplateReport() {
  if (!form.reportValidity()) return;

  const from = parseControlDate(fromDateInput.value);
  const to = parseControlDate(toDateInput.value);
  const compareFrom = parseControlDate(nextFromDateInput.value);
  const compareTo = parseControlDate(nextToDateInput.value);
  if (!validateRequiredDates(from, to, compareFrom, compareTo) || !validatePeriodDates(from, to, compareFrom, compareTo)) return;

  setBusy(true);
  setStatus("Đang lấy dữ liệu Tỉnh cho kỳ báo cáo và kỳ so sánh...");

  try {
    const tuyChonGR = "TINH";
    const weights = readWeights();

    setStatus("Đang gọi dữ liệu Tỉnh: kỳ báo cáo...");
    const currentRows = finalizeRows(await fetchCombinedReport(tuyChonGR, from, to, "kỳ báo cáo"), weights, tuyChonGR);

    setStatus("Đang gọi dữ liệu Tỉnh: kỳ so sánh...");
    const compareRows = finalizeRows(await fetchCombinedReport(tuyChonGR, compareFrom, compareTo, "kỳ so sánh"), weights, tuyChonGR);

    if (!currentRows.length || !compareRows.length) {
      throw new Error("Không đủ dữ liệu cấp Tỉnh cho 2 kỳ để lập báo cáo so sánh.");
    }

    const comparisonRows = buildProvinceComparisonRows(currentRows, compareRows);
    if (!comparisonRows.length) {
      throw new Error("Không tìm thấy tỉnh trùng mã giữa kỳ báo cáo và kỳ so sánh.");
    }

    const provinceCode = "10";
    const provinceComparison = comparisonRows.find((row) => String(row.code) === provinceCode);
    if (!provinceComparison) {
      throw new Error("Không tìm thấy dữ liệu mã tỉnh = 10 (BĐ Hà Nội) trong 2 kỳ báo cáo.");
    }

    const fileName = `baocao_so_sanh_ky_tinh_${toFileDate(from)}_${toFileDate(to)}_vs_${toFileDate(compareFrom)}_${toFileDate(compareTo)}.xls`;
    renderTemplateReportPreview({ rows: comparisonRows, provinceComparison, from, to, compareFrom, compareTo, fileName, weights });
    downloadProvinceTemplateExcel({
      currentRows,
      compareRows,
      comparisonRows,
      from,
      to,
      compareFrom,
      compareTo,
      fileName
    });
    setStatus(`Hoàn tất báo cáo so sánh kỳ Tỉnh: ${comparisonRows.length} đơn vị, đã xuất ${fileName}.`, "ok");
  } catch (error) {
    setStatus(error.message || "Có lỗi khi lấy dữ liệu báo cáo so sánh kỳ Tỉnh.", "error");
  } finally {
    setBusy(false);
  }
}

async function exportReportV2(tuyChonGR, mode = "data") {
  if (!form.reportValidity()) return;

  const from = parseControlDate(fromDateInput.value);
  const to = parseControlDate(toDateInput.value);
  const compareFrom = parseControlDate(nextFromDateInput.value);
  const compareTo = parseControlDate(nextToDateInput.value);
  if (!validateRequiredDates(from, to, compareFrom, compareTo) || !validatePeriodDates(from, to, compareFrom, compareTo)) return;

  setBusy(true);
  setStatus(mode === "compare" ? "Đang lấy dữ liệu V2 cho 2 kỳ so sánh..." : `Đang lấy dữ liệu filebaocaov2.xlsx (${tuyChonGR})...`);

  try {
    const weights = readWeights();
    setStatus("Đang gọi dữ liệu V2: kỳ báo cáo...");
    const currentRows = await fetchV2ReportRows(tuyChonGR, from, to, weights, mode === "compare" ? "kỳ báo cáo" : "");

    if (mode === "compare") {
      setStatus("Đang gọi dữ liệu V2: kỳ so sánh...");
      const compareRows = await fetchV2ReportRows(tuyChonGR, compareFrom, compareTo, weights, "kỳ so sánh");
      const comparisonRows = buildV2ComparisonRows(currentRows, compareRows);
      const provinceComparison = comparisonRows.find((row) => String(row.code) === "10");
      if (!provinceComparison) {
        throw new Error("Không tìm thấy dữ liệu mã tỉnh = 10 (BĐ Hà Nội) trong 2 kỳ báo cáo V2.");
      }
      const fileName = `filebaocaov2_so_sanh_ky_${toFileDate(from)}_${toFileDate(to)}_vs_${toFileDate(compareFrom)}_${toFileDate(compareTo)}.xls`;
      renderV2ReportPreview({ rows: comparisonRows, tuyChonGR, from, to, compareFrom, compareTo, fileName, mode: "compare", provinceComparison });
      downloadV2ComparisonExcel({ currentRows, compareRows, comparisonRows, tuyChonGR, from, to, compareFrom, compareTo, fileName });
      setStatus(`Hoàn tất Báo cáo so sánh kỳ V2: ${comparisonRows.length} dòng, đã xuất ${fileName}.`, "ok");
      return;
    }

    const reportName = tuyChonGR === "TINH" ? "tinh" : "buu_cuc";
    const fileName = `filebaocaov2_${reportName}_${toFileDate(from)}_${toFileDate(to)}.xls`;
    renderV2ReportPreview({ rows: currentRows, tuyChonGR, from, to, fileName, mode: "data" });
    downloadV2Excel({ rows: currentRows, tuyChonGR, from, to, fileName });
    setStatus(`Hoàn tất Báo cáo ${tuyChonGR === "TINH" ? "theo tỉnh" : "theo bưu cục"} V2: ${currentRows.length} dòng, đã xuất ${fileName}.`, "ok");
  } catch (error) {
    setStatus(error.message || "Có lỗi khi lấy dữ liệu V2.", "error");
  } finally {
    setBusy(false);
  }
}

async function fetchCombinedReport(tuyChonGR, from, to, periodLabel = "") {
  const sourceTables = [];

  for (const config of SOURCE_CONFIGS) {
    setStatus(config.status);
    const requestUrl = buildApiRequestUrl(config.key, config.buildParams(tuyChonGR, from, to));
    const payload = await fetchApi(config.key, config.buildParams(tuyChonGR, from, to));
    const records = parseApiRows(payload.data || "", config, tuyChonGR);
    console.log(`[DKCL][${tuyChonGR}] ${config.name}: parsed records`, records.length, records.slice(0, 3));
    if (!records.length) {
      throw new Error(buildMissingSourceDataMessage(config.name, periodLabel, from, to, requestUrl));
    }
    sourceTables.push({ config, records });
  }

  return joinProvinceTables(sourceTables, tuyChonGR);
}

function buildMissingSourceDataMessage(sourceName, periodLabel, from, to, requestUrl = "") {
  const periodText = periodLabel ? `${periodLabel} ` : "";
  const curlText = requestUrl ? `\nLink curl: ${requestUrl}` : "";
  return `Biểu "${sourceName}" không có dữ liệu cho ${periodText}(${toApiDate(from)} - ${toApiDate(to)}).${curlText}\nVui lòng kiểm tra đăng nhập, khoảng ngày hoặc điều kiện lọc.`;
}

function getActiveV2SourceConfigs() {
  const selected = getSelectedMetricLabels();
  return V2_SOURCE_CONFIGS.filter((config) => {
    if (config.key === "phatThanhCongBuuCuc") return selected.includes("F4.1 – Chất lượng phát thành công tại bưu cục");
    if (config.key === "noiTinhF11") return selected.includes("F1.1 – Nội tỉnh");
    if (config.key === "thuGomLienTinhF12") return selected.includes("F1.2 – Thu gom bưu gửi đi liên tỉnh");
    if (config.key === "phatLienTinhF13") return selected.includes("F1.3 – Chất lượng phát bưu gửi liên tỉnh");
    return true;
  });
}

async function fetchV2ReportRows(tuyChonGR, from, to, weights, periodLabel = "", selectedBcProvCode = null, selectedBcProvName = null) {
  let baseRows = [];

  if (tuyChonGR === "TINH") {
    const PROVINCE_NAMES = {
      10: "TP. Hà Nội", 16: "Hưng Yên", 18: "TP. Hải Phòng", 20: "Quảng Ninh", 22: "Bắc Ninh",
      24: "Lạng Sơn", 25: "Thái Nguyên", 27: "Cao Bằng", 29: "Phú Thọ", 30: "Tuyên Quang",
      33: "Lào Cai", 36: "Sơn La", 38: "Điện Biên", 39: "Lai Châu", 43: "Ninh Bình",
      44: "Thanh Hóa", 46: "Nghệ An", 48: "Hà Tĩnh", 52: "Quảng Trị", 53: "Thừa Thiên Huế",
      55: "TP. Đà Nẵng", 57: "Quảng Ngãi", 60: "Gia Lai", 63: "Đắc Lắk", 65: "Khánh Hòa",
      67: "Lâm Đồng", 70: "TP. Hồ Chí Minh", 81: "Đồng Nai", 84: "Tây Ninh", 87: "Đồng Tháp",
      88: "An Giang", 89: "Vĩnh Long", 90: "TP. Cần Thơ", 97: "Cà Mau"
    };

    baseRows = Object.entries(PROVINCE_NAMES).map(([code, name], idx) => {
      const v2Row = Array(V2_EXCEL_COLUMNS.length).fill("");
      v2Row[0] = String(idx + 1);
      v2Row[PROVINCE_CODE_INDEX] = code;
      v2Row[PROVINCE_NAME_INDEX] = name;
      return v2Row;
    });
  }

  const activeConfigs = getActiveV2SourceConfigs();
  if (!activeConfigs.length) {
    throw new Error("Vui lòng chọn ít nhất 1 chỉ tiêu chất lượng.");
  }

  let rowMap = new Map(baseRows.map((row) => [String(row[PROVINCE_CODE_INDEX]), row]));
  let joinCodeSet = new Set(rowMap.keys());

  for (const config of activeConfigs) {
    setStatus(config.status);
    const requestUrl = buildApiRequestUrl(config.key, config.buildParams(tuyChonGR, from, to));
    const payload = await fetchApi(config.key, config.buildParams(tuyChonGR, from, to));
    const records = parseV2ApiRows(payload.data || "", config, tuyChonGR, joinCodeSet);
    const aggregatedRecords = aggregateV2RecordsByCode(records, config);
    console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: parsed records`, records.length, records.slice(0, 3));
    console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: aggregated records before join`, aggregatedRecords.length, aggregatedRecords.slice(0, 5));

    if (!baseRows.length && aggregatedRecords.length) {
      aggregatedRecords.forEach((rec, idx) => {
        const v2Row = Array(V2_EXCEL_COLUMNS.length).fill("");
        v2Row[0] = String(idx + 1);
        v2Row[PROVINCE_CODE_INDEX] = String(rec.provinceCodeInt);
        v2Row[PROVINCE_NAME_INDEX] = rec.provinceName || `Đơn vị ${rec.provinceCodeInt}`;
        baseRows.push(v2Row);
        rowMap.set(String(rec.provinceCodeInt), v2Row);
      });
      joinCodeSet = new Set(rowMap.keys());
    }

    if (aggregatedRecords.length) {
      for (const record of aggregatedRecords) {
        if (!Number.isInteger(record.provinceCodeInt)) continue;
        if (tuyChonGR === "TINH" && [1, 8].includes(record.provinceCodeInt)) continue;
        let row = rowMap.get(String(record.provinceCodeInt));
        if (!row) {
          row = Array(V2_EXCEL_COLUMNS.length).fill("");
          row[0] = String(baseRows.length + 1);
          row[PROVINCE_CODE_INDEX] = String(record.provinceCodeInt);
          row[PROVINCE_NAME_INDEX] = record.provinceName || `Đơn vị ${record.provinceCodeInt}`;
          baseRows.push(row);
          rowMap.set(String(record.provinceCodeInt), row);
        }
        Object.entries(config.targets).forEach(([field, target]) => {
          if (field === "rate" || field === "rank") return;
          row[target] = zeroIfBlank(record[field]);
        });
      }
    }
  }

  if (!baseRows.length) {
    throw new Error(buildMissingSourceDataMessage(activeConfigs.map((c) => c.name).join(" / "), periodLabel, from, to));
  }

  finalizeV2Rows(baseRows);
  console.log(`[DKCL][V2][${tuyChonGR}] final rows after V2 join`, baseRows.length, baseRows.slice(0, 10));
  return baseRows;
}

function parseV2ApiRows(html, config, tuyChonGR, joinCodeSet = null) {
  const documentHtml = new DOMParser().parseFromString(`<table><tbody>${html}</tbody></table>`, "text/html");
  const rows = [...documentHtml.querySelectorAll("tr")]
    .filter((tr) => tr.children.length > 2 && !tr.classList.contains("tr_tong"));
  const cellRows = rows.map((tr) => [...tr.children].map((td) => cleanText(td.textContent)));
  console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: raw rows`, rows.length, "sample cells", cellRows.slice(0, 10));

  return cellRows
    .map((cells) => {
      const provinceCodeText = pickV2JoinCodeText(cells, config, tuyChonGR, joinCodeSet);
      const provinceCodeInt = toIntegerCode(provinceCodeText);
      const provinceName = pickReportName(cells, config, tuyChonGR);
      return {
        cells,
        provinceCodeText,
        provinceCodeInt,
        provinceName,
        total: cells[config.values.total] || "",
        subTotal: Number.isInteger(config.values.subTotal) ? cells[config.values.subTotal] || "" : "",
        onTime: cells[config.values.onTime] || ""
      };
    })
    .filter((record) => Number.isInteger(record.provinceCodeInt));
}

function pickV2JoinCodeText(cells, config, tuyChonGR, joinCodeSet) {
  if (tuyChonGR === "BC" && joinCodeSet?.size) {
    const matchedCode = cells
      .map((cell) => String(toIntegerCode(cell)))
      .find((code) => code !== "10" && joinCodeSet.has(code));
    if (matchedCode) return matchedCode;
  }

  return pickReportCodeText(cells, config, tuyChonGR);
}

function aggregateV2RecordsByCode(records, config) {
  const recordMap = new Map();

  for (const record of records) {
    const key = String(record.provinceCodeInt);
    if (!recordMap.has(key)) {
      recordMap.set(key, {
        ...record,
        total: "0",
        subTotal: Number.isInteger(config.values.subTotal) ? "0" : "",
        onTime: "0"
      });
    }

    const aggregate = recordMap.get(key);
    aggregate.total = sumNumberText(aggregate.total, record.total);
    aggregate.onTime = sumNumberText(aggregate.onTime, record.onTime);
    if (Number.isInteger(config.values.subTotal)) {
      aggregate.subTotal = sumNumberText(aggregate.subTotal, record.subTotal);
    }
  }

  return [...recordMap.values()];
}

function sumNumberText(a, b) {
  return String(toNumberValue(a) + toNumberValue(b));
}

function toNumberValue(value) {
  const number = Number(cleanNumber(value));
  return Number.isFinite(number) ? number : 0;
}

function finalizeV2Rows(rows) {
  const activeConfigs = getActiveV2SourceConfigs();
  activeConfigs.forEach((config) => {
    rows.forEach((row) => {
      const denominator = zeroIfBlank(row[config.targets.total]);
      const numerator = zeroIfBlank(row[config.targets.onTime]);
      row[config.targets.total] = denominator;
      row[config.targets.onTime] = numerator;
      if (Number.isInteger(config.targets.subTotal)) row[config.targets.subTotal] = zeroIfBlank(row[config.targets.subTotal]);
      row[config.targets.rate] = percentFrom(numerator, denominator);
    });
    assignRankByColumn(rows, config.targets.rate, config.targets.rank);
  });
  rows.forEach((row) => {
    const adVal = toNumberValue(row[29]);
    const ahVal = toNumberValue(row[33]);
    row[35] = String((adVal + ahVal) / 2);
  });
  assignRankByColumn(rows, 35, 36);
  rows.forEach((row, index) => (row[0] = String(index + 1)));
}

function buildPhatThanhCongV2Params(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhPhat: tuyChonGR === "BC" ? "10" : "ALL",
    stMaLoaiBCPhat: tuyChonGR === "BC" ? "ALL" : "NULL",
    stMaBuuCucPhat: "ALL",
    stLoaiDichVu: "ALL",
    stNhomLoaiKH: "ALL",
    stPhamViTinh: "NULL",
    stLoaiTuyenPhat: "NULL",
    stLoaiPhuongXa: "NULL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: "50000",
    iPage: "1"
  };
}

function buildNoiTinhF11Params(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhChapNhan: tuyChonGR === "BC" ? "10" : "ALL",
    stMaBuuCucNhan: "NULL",
    stMaTinhPhat: "ALL",
    stMaBCKTTinhChapNhan: "NULL",
    stMaBCKTTinhPhat: "NULL",
    stMaLoaiBCKT: "NULL",
    stMaBuuCucPhat: "NULL",
    stLoaiDichVu: "ALL",
    stNhomLoaiKH: "ALL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: "50000",
    iPage: "1"
  };
}

function buildThuGomLienTinhF12Params(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhNhan: tuyChonGR === "BC" ? "10" : "ALL",
    stMaBuuCucNhan: "NULL",
    stMaBCKTTinhNhan: "ALL",
    stLoaiDichVu: "ALL",
    stNhomLoaiKH: "ALL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: "50000",
    iPage: "1"
  };
}

function buildPhatLienTinhF13Params(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhPhat: tuyChonGR === "BC" ? "10" : "ALL",
    stMaBCKTTinhPhat: "ALL",
    stMaBuuCucPhat: "ALL",
    stLoaiDichVu: "ALL",
    stNhomLoaiKH: "ALL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: "50000",
    iPage: "1"
  };
}

function joinProvinceTables(sourceTables, tuyChonGR) {
  const rowMap = new Map();

  for (const { config, records } of sourceTables) {
    for (const record of records) {
      if (!Number.isInteger(record.provinceCodeInt)) continue;
      if (tuyChonGR === "TINH" && [1, 8].includes(record.provinceCodeInt)) continue;
      if (!cleanText(record.provinceName)) continue;

      const row = ensureReportRow(rowMap, record.provinceCodeInt, record.provinceCodeText, record.provinceName);
      config.columns.forEach(({ source, target, type }) => {
        row[target] = normalizeCell(record[source], type);
      });
    }
  }

  return [...rowMap.values()];
}

function buildApiRequestUrl(apiKey, paramsObject) {
  const params = new URLSearchParams(paramsObject);
  return `${API_URLS[apiKey]}?${params.toString()}`;
}

async function fetchApi(apiKey, paramsObject) {
  const requestUrl = buildApiRequestUrl(apiKey, paramsObject);
  console.log(`[DKCL] Request ${apiKey}:`, requestUrl, paramsObject);

  const response = await fetch(requestUrl, {
    method: "GET",
    credentials: "include",
    headers: {
      accept: "*/*",
      "accept-language": "en-US,en;q=0.9,vi;q=0.8",
      referer: "https://dkcl.vnpost.vn/",
      "x-requested-with": "XMLHttpRequest"
    }
  });

  if (!response.ok) {
    throw new Error(`${apiKey}: API trả về lỗi HTTP ${response.status}. Hãy đăng nhập dkcl.vnpost.vn rồi thử lại.`);
  }

  const text = await response.text();
  console.log(`[DKCL] Response ${apiKey}:`, {
    status: response.status,
    ok: response.ok,
    length: text.length,
    preview: text.slice(0, 500)
  });

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`${apiKey}: Response không phải JSON. Có thể phiên đăng nhập đã hết hạn.`);
  }
}


function buildPhatThanhCongTmdtParams(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaHuyenPhat: "",
    stMaTinhPhat: tuyChonGR === "BC" ? "10" : "ALL",
    stMaLoaiBCPhat: tuyChonGR === "BC" ? "ALL" : "NULL",
    stMaBuuCucPhat: "ALL",
    stLoaiDichVu: "ALL",
    "stNhomLoaiBuuGui[]": "TMĐT",
    stNhomLoaiKH: "ALL",
    stPhamViTinh: "NULL",
    stLoaiTuyenPhat: "NULL",
    stLoaiPhuongXa: "NULL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: tuyChonGR === "BC" ? "50000" : "10000",
    iPage: "1"
  };
}

function buildPhatThanhCongTruyenThongParams(tuyChonGR, from, to) {
  const params = buildPhatThanhCongTmdtParams(tuyChonGR, from, to);
  params["stNhomLoaiBuuGui[]"] = "Truyền thống";
  return params;
}

function buildThuGomParams(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhThuGom: tuyChonGR === "BC" ? "10" : "ALL",
    stMaBuuCucThuGom: "NULL",
    stMaBCKTTinhNhan: "NULL",
    stLoaiDichVu: "ALL",
    stNhomLoaiKH: "ALL",
    stPhamViTinh: "NULL",
    iFrom: toApiDate(from),
    iTo: toApiDate(to),
    iPageSize: "50000",
    iPage: "1"
  };
}

function parseApiRows(html, config, tuyChonGR) {
  const documentHtml = new DOMParser().parseFromString(`<table><tbody>${html}</tbody></table>`, "text/html");
  const rows = [...documentHtml.querySelectorAll("tr")]
    .filter((tr) => tr.children.length > 2 && !tr.classList.contains("tr_tong"));
  const cellRows = rows.map((tr) => [...tr.children].map((td) => cleanText(td.textContent)));

  console.log(`[DKCL][${tuyChonGR}] ${config.name}: raw rows`, rows.length, "sample cells", cellRows.slice(0, 3));

  return cellRows
    .map((cells) => mapRecordBySource(cells, config, tuyChonGR))
    .filter((record) => Number.isInteger(record.provinceCodeInt));
}

function mapRecordBySource(cells, config, tuyChonGR) {
  const provinceCodeText = pickReportCodeText(cells, config, tuyChonGR);
  const provinceCodeInt = toIntegerCode(provinceCodeText);
  const provinceName = pickReportName(cells, config, tuyChonGR);

  if (config.key === "thuGomLienTinh") {
    return mapThuGomRecord(cells, provinceCodeText, provinceCodeInt, provinceName);
  }

  return mapPhatThanhCongRecord(cells, provinceCodeText, provinceCodeInt, provinceName);
}

function mapPhatThanhCongRecord(cells, provinceCodeText, provinceCodeInt, provinceName) {
  return {
    cells,
    provinceCodeText,
    provinceCodeInt,
    provinceName,
    total: cells[10] || "",
    success8h: cells[25] || ""
  };
}

function mapThuGomRecord(cells, provinceCodeText, provinceCodeInt, provinceName) {
  return {
    cells,
    provinceCodeText,
    provinceCodeInt,
    provinceName,
    total: cells[8] || "",
    onTime: cells[9] || ""
  };
}

function ensureReportRow(rowMap, provinceCodeInt, provinceCodeText, provinceName) {
  if (!rowMap.has(provinceCodeInt)) {
    const row = Array(EXCEL_COLUMNS.length).fill("");
    row[PROVINCE_CODE_INDEX] = provinceCodeInt;
    row[PROVINCE_NAME_INDEX] = provinceName;
    rowMap.set(provinceCodeInt, row);
  }
  const row = rowMap.get(provinceCodeInt);
  if (!row[PROVINCE_NAME_INDEX] && provinceName) row[PROVINCE_NAME_INDEX] = provinceName;
  if (!row[PROVINCE_CODE_INDEX] && provinceCodeText) row[PROVINCE_CODE_INDEX] = provinceCodeInt;
  return row;
}

function finalizeRows(rows, weights, tuyChonGR) {
  rows.forEach((row) => {
    row[3] = zeroIfBlank(row[3]);
    row[4] = zeroIfBlank(row[4]);
    row[7] = zeroIfBlank(row[7]);
    row[8] = zeroIfBlank(row[8]);
    row[11] = zeroIfBlank(row[11]);
    row[12] = zeroIfBlank(row[12]);
    row[5] = percentFrom(row[4], row[3]);
    row[9] = percentFrom(row[8], row[7]);
    row[13] = percentFrom(row[12], row[11]);
    row[15] = weightedAveragePercent([row[5], row[9], row[13]], weights);
  });

  const finalRows = tuyChonGR === "BC" ? rows : rows.filter(hasAllPositiveRateColumns);

  assignRankByColumn(finalRows, 5, 6);
  assignRankByColumn(finalRows, 9, 10);
  assignRankByColumn(finalRows, 13, 14);
  assignRankByColumn(finalRows, 15, 16);

  finalRows.sort((a, b) => compareNumberOrText(a[PROVINCE_CODE_INDEX], b[PROVINCE_CODE_INDEX]));
  finalRows.forEach((row, index) => {
    row[0] = String(index + 1);
  });

  return finalRows;
}

function hasAllPositiveRateColumns(row) {
  return [5, 9, 13, 15].every((index) => Number(row[index]) > 0);
}

function renderReportPreview({ rows, tuyChonGR, from, to, fileName, weights }) {
  const reportName = tuyChonGR === "TINH" ? "Theo tỉnh" : "Theo bưu cục";

  reportPanel.classList.remove("is-empty");
  reportBadge.textContent = reportName;
  emptyState.hidden = true;
  reportSummary.hidden = false;
  reportText.hidden = false;

  reportSummary.innerHTML = [
    metricTemplate("Loại báo cáo", reportName),
    metricTemplate("Số dòng", rows.length.toLocaleString("vi-VN")),
    metricTemplate("Trọng số", `TMĐT ${weights.tmdt}% · TT ${weights.traditional}% · TG ${weights.pickup}%`),
    metricTemplate("File Excel", fileName)
  ].join("");

  reportText.textContent = `Đã lấy ${rows.length.toLocaleString("vi-VN")} dòng dữ liệu ${reportName.toLowerCase()}.\nKỳ báo cáo: ${toApiDate(from)} - ${toApiDate(to)}\nFile Excel: ${fileName}`;
}

function renderTemplateReportPreview({ rows, provinceComparison, from, to, compareFrom, compareTo, fileName, weights }) {
  reportPanel.classList.remove("is-empty");
  reportBadge.textContent = "So sánh kỳ Tỉnh";
  emptyState.hidden = true;
  reportSummary.hidden = false;
  reportText.hidden = false;

  reportSummary.innerHTML = [
    metricTemplate("Loại báo cáo", "P.DVVH · So sánh kỳ theo tỉnh"),
    metricTemplate("Mã tỉnh", `${provinceComparison.code} · ${provinceComparison.name}`),
    metricTemplate("Kỳ báo cáo", `${toApiDate(from)} - ${toApiDate(to)}`),
    metricTemplate("So sánh với kỳ", `${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`),
    metricTemplate("Số đơn vị xếp hạng", rows.length.toLocaleString("vi-VN")),
    metricTemplate("File Excel", fileName)
  ].join("");

  reportText.textContent = buildProvinceComparisonText({
    rows,
    provinceComparison,
    from,
    to,
    compareFrom,
    compareTo
  });
}

function buildProvinceComparisonText({ rows, provinceComparison, from, to, compareFrom, compareTo }) {
  const totalUnits = rows.length || 36;
  const periodText = `${toApiDate(from)} - ${toApiDate(to)}`;
  const comparePeriodText = `${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`;
  const metricBlocks = provinceComparison.metrics.map((metric) => {
    const isAverage = metric.label.toLowerCase().includes("trung bình");
    const rateLabel = isAverage ? "Tỷ lệ bình quân" : "Tỷ lệ";
    return `🔹 ${metric.label}\n${rateLabel}: ${formatPercentValue(metric.currentRate)} (${rateTrendText(metric)})\nXếp hạng: ${formatRank(metric.currentRank, totalUnits)} (${rankTrendText(metric)})`;
  });

  return [
    `📊 P.DVVH – Kết quả xếp hạng BĐ Hà Nội (So sánh ${totalUnits} đơn vị)`,
    `Kỳ báo cáo: ${periodText}`,
    `So sánh với kỳ: ${comparePeriodText}`,
    "so sánh 2 dữ liệu của 2 kỳ báo cáo và xếp hạng 2 dữ liệu của 2 kỳ báo cáo",
    "",
    metricBlocks.join("\n\n")
  ].join("\n");
}

function metricTemplate(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function downloadExcel({ rows, tuyChonGR, from, to, fileName }) {
  const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8" />
      <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Tinh</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
      <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; }
        th, td { border: 1px solid #000; padding: 5px; vertical-align: middle; }
        th { font-weight: 700; text-align: center; white-space: normal; }
        td { mso-number-format:"\\@"; }
        .number { mso-number-format:"#,##0"; text-align: right; }
        .percent { mso-number-format:"0.00%"; text-align: right; }
        .group { background: #f8cbad; }
        .head { background: #d9e1f2; }
      </style>
    </head>
    <body>
      <table>
        ${renderExcelHeaderRows()}
        ${rows.map((row) => `<tr>${row.map((cell, index) => renderExcelCell(cell, index)).join("")}</tr>`).join("")}
      </table>
    </body>
    </html>`;

  const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
  saveBlobAsFile(blob, fileName);
}

function downloadProvinceTemplateExcel({ currentRows, compareRows, comparisonRows, from, to, compareFrom, compareTo, fileName }) {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Author>DKCL Export Sidebar</Author>
    <Created>${new Date().toISOString()}</Created>
  </DocumentProperties>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10"/></Style>
    <Style ss:ID="Title"><Font ss:FontName="Arial" ss:Size="14" ss:Bold="1"/></Style>
    <Style ss:ID="Header"><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="Group"><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="Cell"><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="Number"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="#,##0"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="Percent"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="Text"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Up"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#008000"/><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Down"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#C00000"/><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Same"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#666666"/><NumberFormat ss:Format="@"/></Style>

    <Style ss:ID="s134"><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="s135"><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center"/></Style>
    <Style ss:ID="s136"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="s137"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#C00000"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="s138"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#008000"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="s139"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#C00000"/></Style>
    <Style ss:ID="s140"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#00B050"/></Style>
    <Style ss:ID="s141"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#A6A6A6"/></Style>
    <Style ss:ID="s142"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#FFC000"/></Style>
    <Style ss:ID="s63"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="14" ss:Bold="1"/></Style>
    <Style ss:ID="s64"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="m1730392435924"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Italic="1"/></Style>
    <Style ss:ID="m1730392435944"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/></Style>
    <Style ss:ID="m1730392434032"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/><Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/></Style>
    <Style ss:ID="m1730392434112"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/></Style>
    <Style ss:ID="m1730392434132"><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/><Interior ss:Color="#FFF2CC" ss:Pattern="Solid"/></Style>
    <Style ss:ID="m1730392431952"><Alignment ss:Vertical="Top" ss:WrapText="1"/><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Size="11"/><Interior ss:Color="#F2F2F2" ss:Pattern="Solid"/></Style>

  </Styles>
  ${renderDataWorksheet("Du lieu dau ky", currentRows, `Dữ liệu đầu kỳ: ${toApiDate(from)} - ${toApiDate(to)}`)}
  ${renderDataWorksheet("Ky so sanh", compareRows, `Dữ liệu kỳ so sánh: ${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`)}
  ${renderComparisonWorksheet(comparisonRows, from, to, compareFrom, compareTo)}
</Workbook>`;

  const blob = new Blob([workbookXml], { type: "application/vnd.ms-excel;charset=utf-8" });
  saveBlobAsFile(blob, fileName);
}

function renderV2ReportPreview({ rows, tuyChonGR, from, to, compareFrom, compareTo, fileName, mode, provinceComparison }) {
  const reportName = mode === "compare" ? "So sánh kỳ V2" : tuyChonGR === "TINH" ? "Theo tỉnh V2" : "Theo bưu cục V2";
  reportPanel.classList.remove("is-empty");
  reportBadge.textContent = reportName;
  emptyState.hidden = true;
  reportSummary.hidden = false;
  reportText.hidden = false;

  const summaryItems = [
    metricTemplate("Loại báo cáo", reportName),
    metricTemplate("Số dòng", rows.length.toLocaleString("vi-VN")),
    metricTemplate("Kỳ báo cáo", `${toApiDate(from)} - ${toApiDate(to)}`)
  ];
  if (mode === "compare" && compareFrom && compareTo) {
    summaryItems.push(metricTemplate("So sánh với kỳ", `${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`));
  }
  summaryItems.push(metricTemplate("File Excel", fileName));
  reportSummary.innerHTML = summaryItems.join("");

  if (mode === "compare" && provinceComparison) {
    reportText.textContent = buildProvinceComparisonText({
      rows,
      provinceComparison,
      from,
      to,
      compareFrom,
      compareTo
    });
    return;
  }

  reportText.textContent = `Đã lấy ${rows.length.toLocaleString("vi-VN")} dòng dữ liệu ${reportName}.\nFile Excel: ${fileName}`;
}

function downloadV2Excel({ rows, tuyChonGR, from, to, fileName }) {
  const workbookXml = buildV2WorkbookXml({ rows, tuyChonGR, from, to });
  const blob = new Blob([workbookXml], { type: "application/vnd.ms-excel;charset=utf-8" });
  saveBlobAsFile(blob, fileName);
}

function downloadV2ComparisonExcel({ currentRows, compareRows, comparisonRows, tuyChonGR, from, to, compareFrom, compareTo, fileName }) {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">
  ${renderWorkbookPropertiesAndStyles()}
  ${renderV2ComparisonDashboardWorksheet(comparisonRows, tuyChonGR, from, to, compareFrom, compareTo)}
  ${renderV2DataWorksheet("Du lieu dau ky V2", currentRows, `Dữ liệu V2 đầu kỳ: ${toApiDate(from)} - ${toApiDate(to)}`)}
  ${renderV2DataWorksheet("Ky so sanh V2", compareRows, `Dữ liệu V2 kỳ so sánh: ${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`)}
  ${renderV2ComparisonWorksheet(comparisonRows, from, to, compareFrom, compareTo)}
</Workbook>`;
  const blob = new Blob([workbookXml], { type: "application/vnd.ms-excel;charset=utf-8" });
  saveBlobAsFile(blob, fileName);
}

function buildV2WorkbookXml({ rows, tuyChonGR, from, to }) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">
  ${renderWorkbookPropertiesAndStyles()}
  ${renderV2DataWorksheet(tuyChonGR === "TINH" ? "Tinh V2" : "Buu cuc V2", rows, `Dữ liệu filebaocaov2.xlsx: ${toApiDate(from)} - ${toApiDate(to)}`)}
</Workbook>`;
}

function renderWorkbookPropertiesAndStyles() {
  return `<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Author>DKCL Export Sidebar</Author>
    <Created>${new Date().toISOString()}</Created>
  </DocumentProperties>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10"/></Style>
    <Style ss:ID="Title"><Font ss:FontName="Arial" ss:Size="14" ss:Bold="1"/></Style>
    <Style ss:ID="Header"><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="Group"><Font ss:FontName="Arial" ss:Bold="1"/><Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="Cell"><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="Number"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="#,##0"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="Percent"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Right"/></Style>
    <Style ss:ID="Text"><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Up"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#008000"/><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Down"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#C00000"/><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="Same"><Borders>${excelXmlBorders()}</Borders><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#666666"/><NumberFormat ss:Format="@"/></Style>
    <Style ss:ID="ProvCard1Title"><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1B365D" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
    <Style ss:ID="ProvCard1Val"><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#1B365D"/><Interior ss:Color="#E8EEF5" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="ProvCard2Title"><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#006666" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
    <Style ss:ID="ProvCard2Val"><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#006666"/><Interior ss:Color="#E0F2F1" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="ProvCard3Title"><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#8B4513" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
    <Style ss:ID="ProvCard3Val"><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#8B4513"/><Interior ss:Color="#F5EBE6" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="ProvCard4Title"><Font ss:FontName="Arial" ss:Size="10" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#4B0082" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
    <Style ss:ID="ProvCard4Val"><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#4B0082"/><Interior ss:Color="#F0E6FA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="RankTop"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#004085"/><Interior ss:Color="#CCE5FF" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="RankUp"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#155724"/><Interior ss:Color="#D4EDDA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="RankDown"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#721C24"/><Interior ss:Color="#F8D7DA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders></Style>
    <Style ss:ID="NoteHeader"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#383D41"/><Interior ss:Color="#E2E3E5" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Left" ss:Vertical="Center"/></Style>
    <Style ss:ID="NoteText"><Font ss:FontName="Arial" ss:Size="10" ss:Color="#1B1E21"/><Interior ss:Color="#F8F9FA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/></Style>
  </Styles>`;
}

function renderV2DataWorksheet(sheetName, rows, title) {
  return `<Worksheet ss:Name="${escapeXml(sheetName)}">
    <Table>
      ${renderColumnWidths([46, 78, 190, 115, 150, 90, 70, 115, 150, 90, 70, 115, 150, 90, 70, 90, 70, 115, 150, 90, 70, 120, 120, 150, 90, 70, 140, 150, 150, 90, 70, 140, 160, 100, 70, 150, 70])}
      <Row><Cell ss:MergeAcross="36" ss:StyleID="Title"><Data ss:Type="String">${escapeXml(title)}</Data></Cell></Row>
      ${renderExcelXmlGroupHeaderRow(V2_GROUP_HEADERS)}
      <Row>${V2_EXCEL_COLUMNS.map((column) => excelXmlCell(column, "String", "Header")).join("")}</Row>
      ${rows.map((row) => `<Row>${row.map((cell, index) => excelXmlCell(cell, getV2ExcelXmlDataType(cell, index), getV2ExcelXmlStyle(index))).join("")}</Row>`).join("")}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;
}

function buildV2ComparisonRows(currentRows, compareRows) {
  const compareMap = new Map(compareRows.map((row) => [String(row[PROVINCE_CODE_INDEX]), row]));
  const rows = currentRows.map((currentRow) => {
    const compareRow = compareMap.get(String(currentRow[PROVINCE_CODE_INDEX]));
    if (!compareRow) return null;
    return {
      code: currentRow[PROVINCE_CODE_INDEX],
      name: currentRow[PROVINCE_NAME_INDEX],
      metrics: [
        buildMetricComparison("F4.1 – PTC TMĐT", currentRow, compareRow, 5, 6),
        buildMetricComparison("F4.1 – PTC Truyền thống", currentRow, compareRow, 9, 10),
        buildMetricComparison("F3.3 – Thu gom", currentRow, compareRow, 13, 14),
        buildMetricComparison("📊 Trung bình", currentRow, compareRow, 15, 16),
        buildMetricComparison("F4.1 – Chất lượng phát thành công tại bưu cục", currentRow, compareRow, 19, 20),
        buildMetricComparison("F1.1 – Nội tỉnh", currentRow, compareRow, 24, 25),
        buildMetricComparison("F1.2 – Thu gom bưu gửi đi liên tỉnh", currentRow, compareRow, 29, 30),
        buildMetricComparison("F1.3 – Chất lượng phát bưu gửi liên tỉnh", currentRow, compareRow, 33, 34),
        buildMetricComparison("Tỷ lệ trung bình chung chất lượng (TGLT,PTC LT)", currentRow, compareRow, 35, 36)
      ]
    };
  }).filter(Boolean).sort((a, b) => compareNumberOrText(a.code, b.code));

  rows.forEach((row) => (row.totalUnits = rows.length));
  return rows;
}


function renderV2ComparisonDashboardWorksheet(rows, tuyChonGR, from, to, compareFrom, compareTo) {
  const entityName = tuyChonGR === 'TINH' ? 'TOÀN QUỐC' : 'BƯU CỤC';
  const unitCode = rows.length > 0 && tuyChonGR === 'BC' ? rows[0].code.substring(0, 2) : '34';
  const dashName = tuyChonGR === 'TINH' ? 'Dashboard_Tinh' : `Dashboard_BC_${unitCode}`;
  
  let totalCurrentVol = 0;
  let totalCompareVol = 0;
  let totalCurrentSuccess = 0;
  let totalCompareSuccess = 0;
  
  let highestVolUnit = null;
  let highestVol = -1;
  let riskUnits = [];
  
  const processedRows = rows.map((r, idx) => {
    const curVol = Number(r.metrics.find(m => m.key === 'phatThanhCongV2')?.currentVolume || 0);
    const compVol = Number(r.metrics.find(m => m.key === 'phatThanhCongV2')?.compareVolume || 0);
    const curRate = Number(r.metrics.find(m => m.key === 'phatThanhCongV2')?.currentRate || 0);
    const compRate = Number(r.metrics.find(m => m.key === 'phatThanhCongV2')?.compareRate || 0);
    const curSucc = Math.round(curVol * curRate);
    const compSucc = Math.round(compVol * compRate);
    
    totalCurrentVol += curVol;
    totalCompareVol += compVol;
    totalCurrentSuccess += curSucc;
    totalCompareSuccess += compSucc;
    
    if (curVol > highestVol) {
      highestVol = curVol;
      highestVolUnit = r.name;
    }
    
    const diffRate = (curRate - compRate) * 100;
    let category = '';
    let categoryStyle = '';
    
    if (curVol > 100 && curRate < 0.8 && diffRate < 5) {
      category = '🔴 NGUY CƠ CAO (SL lớn, Tỷ lệ giảm/thấp)';
      categoryStyle = 's139';
      riskUnits.push({ name: r.name, vol: curVol, rate: curRate });
    } else if (curVol > 100 && curRate >= 0.85) {
      category = '🟢 XUẤT SẮC (SL lớn, Tỷ lệ cao)';
      categoryStyle = 's140';
    } else if (diffRate >= 10) {
      category = '🟡 TIẾN BỘ (Tỷ lệ tăng tốt)';
      categoryStyle = 's142';
    } else {
      category = '⚪ THÚC ĐẨY (SL nhỏ / Tỷ lệ chưa đạt)';
      categoryStyle = 's141';
    }
    
    let chartBar = '';
    let chartStyle = '';
    if (diffRate >= 0) {
      chartBar = '📈 ' + '█'.repeat(Math.min(10, Math.ceil(diffRate/2))) + ' +' + diffRate.toFixed(2) + '%';
      chartStyle = 's138';
    } else {
      chartBar = '📉 ' + '█'.repeat(Math.min(10, Math.ceil(-diffRate/2))) + ' ' + diffRate.toFixed(2) + '%';
      chartStyle = 's137';
    }
    
    return {
      stt: idx + 1, code: r.code, name: r.name,
      curVol, compVol, curRate, compRate, diffRate,
      category, categoryStyle, chartBar, chartStyle
    };
  });
  
  riskUnits.sort((a, b) => b.vol - a.vol);
  const topRisk = riskUnits.slice(0, 3);
  
  const avgCurRate = totalCurrentVol > 0 ? (totalCurrentSuccess / totalCurrentVol) : 0;
  const avgCompRate = totalCompareVol > 0 ? (totalCompareSuccess / totalCompareVol) : 0;
  
  let headerXml = `<Worksheet ss:Name="${escapeXml(dashName)}">
  <Table ss:ExpandedColumnCount="12" ss:ExpandedRowCount="${processedRows.length + 20}" x:FullColumns="1" x:FullRows="1">
   <Column ss:Width="45.75" />
   <Column ss:Width="84.75" />
   <Column ss:AutoFitWidth="0" ss:Width="141.75" />
   <Column ss:AutoFitWidth="0" ss:Width="72.75" ss:Span="4" />
   <Column ss:Index="9" ss:AutoFitWidth="0" ss:Width="74.25" />
   <Column ss:Width="180" />
   <Column ss:Width="150" />
   <Column ss:Width="230.25" />
   <Row ss:Height="18">
    <Cell ss:MergeAcross="11" ss:StyleID="s63"><Data ss:Type="String">📊 BI EXECUTIVE DASHBOARD - TỔNG QUAN NĂNG SUẤT &amp; CHẤT LƯỢNG ${escapeXml(entityName)} (MÃ ${unitCode})</Data></Cell>
   </Row>
   <Row>
    <Cell ss:MergeAcross="11" ss:StyleID="m1730392435924"><Data ss:Type="String">Kỳ báo cáo: ${toApiDate(from)} - ${toApiDate(to)}  |  Kỳ so sánh: ${toApiDate(compareFrom)} - ${toApiDate(compareTo)}</Data></Cell>
   </Row>
   <Row ss:Height="13.5" />
   <Row ss:AutoFitHeight="0" ss:Height="21.9375">
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392435944"><Data ss:Type="String">1. TỔNG SẢN LƯỢNG PHÁT</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392435944"><Data ss:Type="String">2. TỶ LỆ PHÁT THÀNH CÔNG BQ</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392435944"><Data ss:Type="String">3. ĐƠN VỊ GÁNH SẢN LƯỢNG LỚN NHẤT</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392435944"><Data ss:Type="String">4. CẢNH BÁO NGUY CƠ NĂNG SUẤT</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="38.0625">
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392434032"><Data ss:Type="String">${totalCurrentVol.toLocaleString()} bưu gửi&#10;(${totalCurrentVol - totalCompareVol > 0 ? '+' : ''}${(totalCurrentVol - totalCompareVol).toLocaleString()} đơn so với kỳ trước)</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392434032"><Data ss:Type="String">${(avgCurRate * 100).toFixed(2)}%&#10;(${avgCurRate - avgCompRate > 0 ? '+' : ''}${((avgCurRate - avgCompRate) * 100).toFixed(2)} điểm % so với kỳ trước)</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392434032"><Data ss:Type="String">${escapeXml(highestVolUnit || 'N/A')}&#10;(${highestVol.toLocaleString()} đơn - chiếm ${((highestVol / (totalCurrentVol||1)) * 100).toFixed(1)}% toàn mạng)</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392434032"><Data ss:Type="String">${topRisk.length > 0 ? `🔴 ${riskUnits.length} Đơn vị nguy cơ cao&#10;(${topRisk.map(r => r.name).join(', ')})` : '🟢 Không có Đơn vị nguy cơ cao'}</Data></Cell>
   </Row>
   <Row ss:Index="7">
    <Cell ss:MergeAcross="11" ss:StyleID="m1730392434112"><Data ss:Type="String">BẢNG PHÂN TÍCH SO SÁNH 2 KỲ KÉP: SẢN LƯỢNG (ĐƠN VỊ) &amp; TỶ LỆ CHẤT LƯỢNG PHÁT F4.1 (%)</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s134"><Data ss:Type="String">STT</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Mã BC</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Tên Đơn Vị</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Tỷ trọng SL (%)</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">SL Kỳ này</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">SL Kỳ trước</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Biến động SL</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Tỷ lệ Kỳ này</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Tỷ lệ Kỳ trước</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">So sánh Tỷ lệ</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Biểu đồ Biến động</Data></Cell>
    <Cell ss:StyleID="s134"><Data ss:Type="String">Phân loại BI</Data></Cell>
   </Row>`;
   
  const tableXml = processedRows.map(r => {
    const slDiff = r.curVol - r.compVol;
    const diffStr = slDiff > 0 ? `+${slDiff} đơn` : `${slDiff} đơn`;
    const slDiffStyle = slDiff > 0 ? 's138' : (slDiff < 0 ? 's137' : 's135');
    const rateDiffStr = r.diffRate > 0 ? `tăng ${r.diffRate.toFixed(2)} điểm %, từ ${(r.compRate*100).toFixed(2)}% lên ${(r.curRate*100).toFixed(2)}%` : `giảm ${Math.abs(r.diffRate).toFixed(2)} điểm %, từ ${(r.compRate*100).toFixed(2)}% xuống ${(r.curRate*100).toFixed(2)}%`;
    return `   <Row>
    <Cell ss:StyleID="s135"><Data ss:Type="Number">${r.stt}</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">${escapeXml(r.code)}</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">${escapeXml(r.name)}</Data></Cell>
    <Cell ss:StyleID="s136"><Data ss:Type="Number">${r.curVol / (totalCurrentVol || 1)}</Data></Cell>
    <Cell ss:StyleID="s135"><Data ss:Type="Number">${r.curVol}</Data></Cell>
    <Cell ss:StyleID="s135"><Data ss:Type="Number">${r.compVol}</Data></Cell>
    <Cell ss:StyleID="${slDiffStyle}"><Data ss:Type="String">${diffStr}</Data></Cell>
    <Cell ss:StyleID="s136"><Data ss:Type="Number">${r.curRate}</Data></Cell>
    <Cell ss:StyleID="s136"><Data ss:Type="Number">${r.compRate}</Data></Cell>
    <Cell ss:StyleID="${r.chartStyle}"><Data ss:Type="String">${rateDiffStr}</Data></Cell>
    <Cell ss:StyleID="${r.chartStyle}"><Data ss:Type="String">${r.chartBar}</Data></Cell>
    <Cell ss:StyleID="${r.categoryStyle}"><Data ss:Type="String">${escapeXml(r.category)}</Data></Cell>
   </Row>`;
  }).join('\n');
  
  const footerXml = `   <Row ss:Index="${processedRows.length + 10}" ss:Height="15">
    <Cell ss:MergeAcross="11" ss:StyleID="m1730392434132"><Data ss:Type="String">📝 PHÂN TÍCH QUẢN TRỊ BI &amp; CHỈ ĐẠO ĐIỀU HÀNH TỰ ĐỘNG - ${escapeXml(entityName)}</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="110.0625">
    <Cell ss:MergeAcross="11" ss:StyleID="m1730392431952"><Data ss:Type="String">📊 BÁO CÁO PHÂN TÍCH QUẢN TRỊ NĂNG SUẤT &amp; CHẤT LƯỢNG (BI EXECUTIVE REPORT) - ${escapeXml(entityName)}:&#10;1. Tổng sản lượng phát toàn mạng: ${totalCurrentVol.toLocaleString()} bưu gửi (kỳ trước ${totalCompareVol.toLocaleString()} đơn, ${totalCurrentVol - totalCompareVol > 0 ? 'tăng' : 'giảm'} ${(totalCurrentVol - totalCompareVol).toLocaleString()} đơn).&#10;2. Tỷ lệ phát thành công (F4.1) toàn mạng: ${(avgCurRate * 100).toFixed(2)}% (kỳ trước ${(avgCompRate * 100).toFixed(2)}%, ${avgCurRate - avgCompRate > 0 ? 'tăng' : 'giảm'} ${((avgCurRate - avgCompRate) * 100).toFixed(2)} điểm %).&#10;3. Đơn vị gánh sản lượng lớn nhất: ${escapeXml(highestVolUnit || 'N/A')} với ${highestVol.toLocaleString()} đơn (chiếm ${((highestVol / (totalCurrentVol||1)) * 100).toFixed(1)}% tổng sản lượng).&#10;4. ${topRisk.length > 0 ? `🔴 CẢNH BÁO NGUY CƠ CAO (${riskUnits.length} đơn vị): ${topRisk.map(r => `${r.name} (SL ${r.vol} đơn, đạt ${(r.rate*100).toFixed(1)}%)`).join('; ')}. Cần kiểm tra khâu phát gấp!` : '🟢 KHÔNG CÓ CẢNH BÁO NGUY CƠ CAO.'}&#10;5. CHỈ ĐẠO ĐIỀU HÀNH: Yêu cầu Trưởng các Đơn vị thuộc nhóm Cảnh báo nguy cơ tập trung rà soát lực lượng bưu tá và kiểm soát quét TMS hoàn thành đúng chỉ tiêu.</Data></Cell>
   </Row>
  </Table>
 </Worksheet>`;
  return headerXml + '\n' + tableXml + '\n' + footerXml;
}

function renderV2ComparisonWorksheet(rows, from, to, compareFrom, compareTo) {
  return `<Worksheet ss:Name="SoSanhV2">
    <Table>
      ${renderColumnWidths([46, 78, 190, 210, 90, 190, 90, 190])}
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Title"><Data ss:Type="String">📊 P.DVVH – Kết quả xếp hạng BĐ Hà Nội (So sánh ${rows.length} đơn vị)</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Text"><Data ss:Type="String">Kỳ báo cáo: ${escapeXml(toApiDate(from))} - ${escapeXml(toApiDate(to))}</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Text"><Data ss:Type="String">So sánh với kỳ: ${escapeXml(toApiDate(compareFrom))} - ${escapeXml(toApiDate(compareTo))}</Data></Cell></Row>
      <Row>${["STT", "Mã", "Tên", "Chỉ số", "Tỷ lệ", "So sánh tỷ lệ", "Xếp hạng", "So sánh xếp hạng"].map((column) => excelXmlCell(column, "String", "Header")).join("")}</Row>
      ${rows.map(renderComparisonXmlRows).join("")}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;
}

function renderDataWorksheet(sheetName, rows, title) {
  return `<Worksheet ss:Name="${escapeXml(sheetName)}">
    <Table>
      ${renderColumnWidths([46, 78, 190, 115, 150, 90, 70, 115, 150, 90, 70, 115, 150, 90, 70, 90, 70])}
      <Row><Cell ss:MergeAcross="16" ss:StyleID="Title"><Data ss:Type="String">${escapeXml(title)}</Data></Cell></Row>
      ${renderExcelXmlGroupHeaderRow(GROUP_HEADERS)}
      <Row>${EXCEL_COLUMNS.map((column) => excelXmlCell(column, "String", "Header")).join("")}</Row>
      ${rows.map((row) => `<Row>${row.map((cell, index) => excelXmlCell(cell, getExcelXmlDataType(cell, index), getExcelXmlStyle(index))).join("")}</Row>`).join("")}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;
}

function renderComparisonWorksheet(rows, from, to, compareFrom, compareTo) {
  return `<Worksheet ss:Name="SoSanh">
    <Table>
      ${renderColumnWidths([46, 78, 190, 170, 90, 190, 90, 190])}
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Title"><Data ss:Type="String">📊 P.DVVH – Kết quả xếp hạng BĐ Hà Nội (So sánh ${rows.length} đơn vị)</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Text"><Data ss:Type="String">Kỳ báo cáo: ${escapeXml(toApiDate(from))} - ${escapeXml(toApiDate(to))}</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="7" ss:StyleID="Text"><Data ss:Type="String">So sánh với kỳ: ${escapeXml(toApiDate(compareFrom))} - ${escapeXml(toApiDate(compareTo))}</Data></Cell></Row>
      <Row>${["STT", "Mã tỉnh", "Tên tỉnh", "Chỉ số", "Tỷ lệ", "So sánh tỷ lệ", "Xếp hạng", "So sánh xếp hạng"].map((column) => excelXmlCell(column, "String", "Header")).join("")}</Row>
      ${rows.map(renderComparisonXmlRows).join("")}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;
}

function renderComparisonXmlRows(row, index) {
  const totalUnits = row.totalUnits || row.metricsTotal || 36;
  return row.metrics
    .map((metric, metricIndex) => {
      return `<Row>
        ${excelXmlCell(index + 1, "Number", "Number")}
        ${excelXmlCell(row.code, "String", "Text")}
        ${excelXmlCell(row.name, "String", "Text")}
        ${excelXmlCell(metric.label, "String", "Group")}
        ${excelXmlCell(metric.currentRate, "Number", "Percent")}
        ${excelXmlCell(rateTrendText(metric, { plain: true }), "String", trendStyle(metric.rateDiffPoints))}
        ${excelXmlCell(formatRank(metric.currentRank, totalUnits), "String", "Text")}
        ${excelXmlCell(rankTrendText(metric, { plain: true }), "String", trendStyle(metric.rankDiff))}
      </Row>`;
    })
    .join("");
}

function renderColumnWidths(widths) {
  return widths.map((width) => `<Column ss:Width="${width}"/>`).join("");
}

function renderExcelXmlGroupHeaderRow(groups) {
  return `<Row>${groups.map(({ label, colspan }) => `<Cell ss:MergeAcross="${colspan - 1}" ss:StyleID="Group"><Data ss:Type="String">${escapeXml(label)}</Data></Cell>`).join("")}</Row>`;
}

function excelXmlCell(value, type = "String", styleId = "Cell") {
  const safeType = type === "Number" && Number.isFinite(Number(value)) ? "Number" : "String";
  const safeValue = safeType === "Number" ? String(Number(value)) : escapeXml(value);
  return `<Cell ss:StyleID="${styleId}"><Data ss:Type="${safeType}">${safeValue}</Data></Cell>`;
}

function getExcelXmlStyle(index) {
  if (isPercentColumn(index)) return "Percent";
  if (index === PROVINCE_NAME_INDEX) return "Text";
  return isNumberColumn(index) ? "Number" : "Text";
}

function getV2ExcelXmlStyle(index) {
  if (isV2PercentColumn(index)) return "Percent";
  if (index === PROVINCE_NAME_INDEX) return "Text";
  return isV2NumberColumn(index) ? "Number" : "Text";
}

function getExcelXmlDataType(cell, index) {
  if (isPercentColumn(index) || isNumberColumn(index)) return "Number";
  return "String";
}

function getV2ExcelXmlDataType(cell, index) {
  if (isV2PercentColumn(index) || isV2NumberColumn(index)) return "Number";
  return "String";
}

function excelXmlBorders() {
  return '<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>';
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function saveBlobAsFile(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const cleanup = () => setTimeout(() => URL.revokeObjectURL(url), 20_000);

  if (typeof chrome !== "undefined" && chrome.downloads && typeof chrome.downloads.download === "function") {
    chrome.downloads.download({ url, filename: fileName, saveAs: true }, cleanup);
    return;
  }

  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  cleanup();
}

function renderHeaderRows() {
  return `${renderGroupHeaderRow()}<tr>${EXCEL_COLUMNS.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;
}

function renderExcelHeaderRows() {
  return `${renderGroupHeaderRow("group")}<tr>${EXCEL_COLUMNS.map((column) => `<th class="head">${escapeHtml(column)}</th>`).join("")}</tr>`;
}

function renderGroupHeaderRow(className = "") {
  return `<tr>${GROUP_HEADERS.map(({ label, colspan }) => `<th class="${className}" colspan="${colspan}">${escapeHtml(label)}</th>`).join("")}</tr>`;
}

function renderTemplateHeaderRows() {
  return `${renderTemplateGroupHeaderRow()}<tr>${TEMPLATE_EXCEL_COLUMNS.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;
}

function renderTemplateExcelHeaderRows() {
  return `${renderTemplateGroupHeaderRow("group")}<tr>${TEMPLATE_EXCEL_COLUMNS.map((column) => `<th class="head">${escapeHtml(column)}</th>`).join("")}</tr>`;
}

function renderTemplateGroupHeaderRow(className = "") {
  return `<tr>${TEMPLATE_GROUP_HEADERS.map(({ label, colspan }) => `<th class="${className}" colspan="${colspan}">${escapeHtml(label)}</th>`).join("")}</tr>`;
}

function renderExcelCell(cell, index) {
  const className = isPercentColumn(index) ? "percent" : isNumberColumn(index) ? "number" : "";
  return `<td class="${className}">${escapeHtml(displayCell(cell))}</td>`;
}

function isNumberColumn(index) {
  return ![PROVINCE_NAME_INDEX, 5, 9, 13, 15].includes(index);
}

function isPercentColumn(index) {
  return [5, 9, 13, 15].includes(index);
}

function isV2PercentColumn(index) {
  return [5, 9, 13, 15, 19, 24, 29, 33, 35].includes(index);
}

function isV2NumberColumn(index) {
  return index !== PROVINCE_NAME_INDEX && !isV2PercentColumn(index);
}

function normalizeCell(value, type) {
  if (type === "percent") return normalizePercent(value);
  return cleanNumber(value);
}

function cleanNumber(value) {
  return String(value ?? "").replace(/,/g, "").trim();
}

function normalizePercent(value) {
  const raw = cleanNumber(value).replace("%", "");
  if (!raw) return "";
  const number = Number(raw);
  if (!Number.isFinite(number)) return raw;
  return number > 1 ? String(number / 100) : String(number);
}

function readWeights() {
  return {
    tmdt: Number(weightTmdtInput.value || 0),
    traditional: Number(weightTraditionalInput.value || 0),
    pickup: Number(weightPickupInput.value || 0)
  };
}

function zeroIfBlank(value) {
  const raw = cleanNumber(value);
  if (!raw) return "0";
  const number = Number(raw);
  return Number.isFinite(number) ? String(number) : "0";
}

function percentFrom(numerator, denominator) {
  const top = Number(zeroIfBlank(numerator));
  const bottom = Number(zeroIfBlank(denominator));
  if (!Number.isFinite(top) || !Number.isFinite(bottom) || bottom === 0) return "0";
  return String(top / bottom);
}

function weightedAveragePercent(values, weights) {
  const weightValues = [weights.tmdt, weights.traditional, weights.pickup];
  let totalWeight = 0;
  let totalScore = 0;

  values.forEach((value, index) => {
    const percent = Number(normalizePercent(value));
    const weight = Number(weightValues[index]);
    if (Number.isFinite(percent) && Number.isFinite(weight) && weight > 0) {
      totalScore += percent * weight;
      totalWeight += weight;
    }
  });

  return totalWeight ? String(totalScore / totalWeight) : "";
}

function buildProvinceComparisonRows(currentRows, compareRows) {
  const compareMap = new Map(compareRows.map((row) => [String(row[PROVINCE_CODE_INDEX]), row]));

  return currentRows
    .map((currentRow) => {
      const compareRow = compareMap.get(String(currentRow[PROVINCE_CODE_INDEX]));
      if (!compareRow) return null;
      return {
        code: currentRow[PROVINCE_CODE_INDEX],
        name: currentRow[PROVINCE_NAME_INDEX],
        metrics: [
          buildMetricComparison("F4.1 – PTC TMĐT", currentRow, compareRow, 5, 6),
          buildMetricComparison("F4.1 – PTC Truyền thống", currentRow, compareRow, 9, 10),
          buildMetricComparison("F3.3 – Thu gom", currentRow, compareRow, 13, 14),
          buildMetricComparison("📊 Trung bình", currentRow, compareRow, 15, 16)
        ]
      };
    })
    .filter(Boolean)
    .sort((a, b) => compareNumberOrText(a.code, b.code));
}

function buildMetricComparison(label, currentRow, compareRow, rateIndex, rankIndex) {
  const currentRate = Number(normalizePercent(currentRow[rateIndex]));
  const compareRate = Number(normalizePercent(compareRow[rateIndex]));
  const currentRank = Number(currentRow[rankIndex]);
  const compareRank = Number(compareRow[rankIndex]);
  return {
    label,
    currentRate,
    compareRate,
    rateDiffPoints: (currentRate - compareRate) * 100,
    currentRank,
    compareRank,
    rankDiff: compareRank - currentRank
  };
}

function renderComparisonHeaderRows() {
  return `
    <tr>
      <th>STT</th><th>Mã tỉnh</th><th>Tên tỉnh</th><th>Chỉ số</th>
      <th>Tỷ lệ kỳ báo cáo</th><th>Tỷ lệ kỳ so sánh</th><th>Biến động tỷ lệ</th>
      <th>Hạng kỳ báo cáo</th><th>Hạng kỳ so sánh</th><th>Biến động hạng</th>
    </tr>`;
}

function renderComparisonPreviewRow(row, index) {
  return row.metrics
    .map((metric, metricIndex) => `
      <tr>
        ${metricIndex === 0 ? `<td rowspan="4">${index + 1}</td><td rowspan="4">${escapeHtml(row.code)}</td><td rowspan="4">${escapeHtml(row.name)}</td>` : ""}
        <td>${escapeHtml(metric.label)}</td>
        <td>${formatPercentValue(metric.currentRate)}</td>
        <td>${formatPercentValue(metric.compareRate)}</td>
        <td>${renderRateTrend(metric)}</td>
        <td>${formatRank(metric.currentRank, row.metrics.length ? row.metricsTotal : null)}</td>
        <td>${formatRank(metric.compareRank, row.metrics.length ? row.metricsTotal : null)}</td>
        <td>${renderRankTrend(metric)}</td>
      </tr>`)
    .join("");
}

function renderComparisonExcelHeaderRows() {
  return `
    <tr>
      <th>STT</th><th>Mã tỉnh</th><th>Tên tỉnh</th><th>Chỉ số</th>
      <th>Tỷ lệ</th><th>So sánh tỷ lệ</th><th>Xếp hạng</th><th>So sánh xếp hạng</th>
    </tr>`;
}

function renderComparisonExcelRow(row, index) {
  const totalUnits = "36";
  return row.metrics
    .map((metric, metricIndex) => `
      <tr>
        ${metricIndex === 0 ? `<td rowspan="4" class="number">${index + 1}</td><td rowspan="4" class="number">${escapeHtml(row.code)}</td><td rowspan="4">${escapeHtml(row.name)}</td>` : ""}
        <td class="metric">🔹 ${escapeHtml(metric.label)}</td>
        <td class="percent">${formatPercentValue(metric.currentRate)}</td>
        <td>${renderRateTrend(metric)}</td>
        <td>${formatRank(metric.currentRank, totalUnits)}</td>
        <td>${renderRankTrend(metric)}</td>
      </tr>`)
    .join("");
}

function renderRateTrend(metric) {
  return `<span class="${trendCssClass(metric.rateDiffPoints)}">${escapeHtml(rateTrendText(metric))}</span>`;
}

function renderRankTrend(metric) {
  return `<span class="${trendCssClass(metric.rankDiff)}">${escapeHtml(rankTrendText(metric))}</span>`;
}

function rateTrendText(metric, options = {}) {
  const diff = metric.rateDiffPoints;
  const direction = getTrendDirection(diff);
  const verb = direction > 0 ? "tăng" : direction < 0 ? "giảm" : "không đổi";
  const icon = options.plain ? "" : direction > 0 ? "🟩▲ " : direction < 0 ? "🟥▼ " : "➖ ";
  const connector = direction > 0 ? "lên" : direction < 0 ? "xuống" : "đến";
  return `${icon}${verb} ${formatPointDiff(Math.abs(diff))}, từ ${formatPercentValue(metric.compareRate)} ${connector} ${formatPercentValue(metric.currentRate)}`;
}

function rankTrendText(metric, options = {}) {
  const diff = metric.rankDiff;
  const direction = getTrendDirection(diff);
  const verb = direction > 0 ? "tăng" : direction < 0 ? "giảm" : "không đổi";
  const icon = options.plain ? "" : direction > 0 ? "🟩▲ " : direction < 0 ? "🟥▼ " : "➖ ";
  const rankText = direction === 0 ? "0 bậc" : `${Math.abs(diff)} bậc`;
  return `${icon}${verb} ${rankText}, hạng ${formatRankNumber(metric.compareRank)} -> ${formatRankNumber(metric.currentRank)}`;
}

function trendCssClass(diff) {
  const direction = getTrendDirection(diff);
  return direction > 0 ? "up" : direction < 0 ? "down" : "same";
}

function trendStyle(diff) {
  const direction = getTrendDirection(diff);
  return direction > 0 ? "Up" : direction < 0 ? "Down" : "Same";
}

function getTrendDirection(diff) {
  if (!Number.isFinite(diff) || Math.abs(diff) < 0.005) return 0;
  return diff > 0 ? 1 : -1;
}

function formatPercentValue(value) {
  return Number.isFinite(value) ? `${(value * 100).toLocaleString("vi-VN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%` : "";
}

function formatPointDiff(value) {
  return `${value.toLocaleString("vi-VN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} điểm %`;
}

function formatRank(rank, totalUnits = "36") {
  return Number.isFinite(rank) ? `${rank} / ${totalUnits || "36"}` : "";
}

function formatRankNumber(rank) {
  return Number.isFinite(rank) ? rank : "";
}

function assignRankByColumn(rows, valueIndex, rankIndex) {
  const sorted = [...rows]
    .filter((row) => Number.isFinite(Number(row[valueIndex])) && Number(row[valueIndex]) > 0)
    .sort((a, b) => Number(b[valueIndex]) - Number(a[valueIndex]) || compareNumberOrText(a[PROVINCE_CODE_INDEX], b[PROVINCE_CODE_INDEX]));

  rows.forEach((row) => {
    row[rankIndex] = "";
  });

  sorted.forEach((row, index) => {
    row[rankIndex] = String(index + 1);
  });
}

function isPercentLike(value) {
  const raw = cleanNumber(value).replace("%", "");
  const number = Number(raw);
  return Number.isFinite(number) && number >= 0 && number <= 1;
}

function compareNumberOrText(a, b) {
  const numberA = Number(a);
  const numberB = Number(b);
  if (Number.isFinite(numberA) && Number.isFinite(numberB)) return numberA - numberB;
  return String(a).localeCompare(String(b), "vi");
}

function displayCell(value) {
  return String(value ?? "");
}

function toInputDate(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function toFileDate(date) {
  return toInputDate(date);
}

function parseControlDate(value) {
  const match = String(value ?? "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return date;
}

function toDisplayDate(date) {
  return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function syncAllDateDisplays() {
  [fromDateInput, toDateInput, nextFromDateInput, nextToDateInput].forEach(syncDateDisplay);
}

function syncDateDisplay(input) {
  const displayInput = dateDisplayMap.get(input);
  if (!displayInput) return;

  const date = parseControlDate(input.value);
  displayInput.value = date ? toDisplayDate(date) : "";
}

function syncNativeDateFromDisplay(displayInput, options = {}) {
  const nativeInput = displayDateMap.get(displayInput);
  if (!nativeInput) return false;

  const raw = cleanText(displayInput.value);
  if (!raw) {
    nativeInput.value = "";
    return false;
  }

  const date = parseDisplayDate(raw);
  if (!date) {
    if (options.showError) setStatus("Ngày nhập chưa đúng định dạng dd/MM/yyyy.", "error");
    return false;
  }

  nativeInput.value = toInputDate(date);
  if (options.normalize) displayInput.value = toDisplayDate(date);
  return true;
}

function parseDisplayDate(value) {
  const match = String(value ?? "").trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return date;
}

function toApiDate(date) {
  return `${pad(date.getMonth() + 1)}/${pad(date.getDate())}/${date.getFullYear()}`;
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function validateRequiredDates(...dates) {
  if (dates.every(Boolean)) return true;
  setStatus("Vui lòng chọn ngày hợp lệ.", "error");
  return false;
}

function validatePeriodDates(from, to, nextFrom, nextTo) {
  if (from > to) {
    setStatus("Ngày bắt đầu của Kỳ báo cáo không được lớn hơn ngày kết thúc.", "error");
    return false;
  }

  if (nextFrom > nextTo) {
    setStatus("Ngày bắt đầu của Kỳ so sánh không được lớn hơn ngày kết thúc.", "error");
    return false;
  }

  return true;
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function pickReportCodeText(cells, config, tuyChonGR) {
  const codeIndex = getReportCodeIndex(cells, config, tuyChonGR);
  if (Number.isInteger(toIntegerCode(cells[codeIndex]))) return cells[codeIndex];
  return cells.find((cell, index) => index > codeIndex && Number.isInteger(toIntegerCode(cell))) || "";
}

function pickReportName(cells, config, tuyChonGR) {
  const codeIndex = getReportCodeIndex(cells, config, tuyChonGR);
  const nameIndex = Number.isInteger(config.nameIndex) ? config.nameIndex : codeIndex + 1;
  if (cells[nameIndex] && !Number.isInteger(toIntegerCode(cells[nameIndex]))) return cells[nameIndex];
  return cells.find((cell, index) => index > codeIndex && /[A-Za-zÀ-ỹ]/.test(cell) && !isPercentLike(cell)) || "";
}

function getReportCodeIndex(cells, config, tuyChonGR) {
  if (Number.isInteger(config.codeIndex)) return config.codeIndex;
  if (tuyChonGR === "BC") return (config.key === "thuGomLienTinh" || config.key === "thuGomLienTinhF12") ? 3 : 5;
  return getProvinceCodeIndex(cells);
}

function pickPostOfficeCodeText(cells, config) {
  const codeIndex = getPostOfficeCodeIndex(cells, config);
  if (Number.isInteger(toIntegerCode(cells[codeIndex]))) return cells[codeIndex];
  return cells.find((cell, index) => index > codeIndex && Number.isInteger(toIntegerCode(cell))) || "";
}

function pickPostOfficeName(cells, config) {
  const codeIndex = getPostOfficeCodeIndex(cells, config);
  const nameIndex = codeIndex + 1;
  if (cells[nameIndex] && !Number.isInteger(toIntegerCode(cells[nameIndex]))) return cells[nameIndex];
  return cells.find((cell, index) => index > codeIndex && /[A-Za-zÀ-ỹ]/.test(cell) && !isPercentLike(cell)) || "";
}

function getPostOfficeCodeIndex(cells, config) {
  return (config.key === "thuGomLienTinh" || config.key === "thuGomLienTinhF12") ? 3 : 5;
}

function pickProvinceCodeText(cells) {
  if (Number.isInteger(toIntegerCode(cells[1]))) return cells[1];
  if (Number.isInteger(toIntegerCode(cells[2]))) return cells[2];
  if (Number.isInteger(toIntegerCode(cells[3]))) return cells[3];
  return cells.find((cell, index) => index > 0 && Number.isInteger(toIntegerCode(cell))) || "";
}

function pickProvinceName(cells) {
  const codeIndex = getProvinceCodeIndex(cells);
  if (codeIndex >= 0 && cells[codeIndex + 1] && !Number.isInteger(toIntegerCode(cells[codeIndex + 1]))) {
    return cells[codeIndex + 1];
  }
  return cells.find((cell, index) => index > 0 && /[A-Za-zÀ-ỹ]/.test(cell) && !isPercentLike(cell)) || "";
}

function pickDataValuesAfterProvince(cells) {
  const codeIndex = getProvinceCodeIndex(cells);
  const startIndex = codeIndex >= 0 ? codeIndex + 2 : 3;
  return cells.slice(startIndex).filter((cell) => isNumericLike(cell) && !isPercentLike(cell));
}

function getProvinceCodeIndex(cells) {
  if (Number.isInteger(toIntegerCode(cells[1]))) return 1;
  if (Number.isInteger(toIntegerCode(cells[2]))) return 2;
  if (Number.isInteger(toIntegerCode(cells[3]))) return 3;
  return cells.findIndex((cell, index) => index > 0 && Number.isInteger(toIntegerCode(cell)));
}

function toIntegerCode(value) {
  const raw = String(value ?? "").trim();
  if (!/^\d+$/.test(raw)) return NaN;
  return Number.parseInt(raw, 10);
}

function isNumericLike(value) {
  const raw = cleanNumber(value);
  return raw !== "" && Number.isFinite(Number(raw));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setStatus(message, type = "") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`.trim();
}

function setBusy(isBusy) {
  allActionButtons.forEach((button) => (button.disabled = isBusy));
}

function getSelectedProvinceCode() {
  const select = document.getElementById("province-select");
  return select ? select.value : "53";
}

function getSelectedProvinceName() {
  const select = document.getElementById("province-select");
  if (!select || !select.selectedOptions || !select.selectedOptions.length) return "Thừa Thiên Huế";
  const text = select.selectedOptions[0].textContent.trim();
  return text.replace(/^\d+\s*-\s*/, "");
}

function getSelectedMetricLabels() {
  const chks = document.querySelectorAll("#metrics-checkbox-list input[type='checkbox']:checked");
  if (!chks || !chks.length) {
    return [
      "F4.1 – Chất lượng phát thành công tại bưu cục",
      "F1.3 – Chất lượng phát bưu gửi liên tỉnh",
      "F1.1 – Nội tỉnh",
      "F1.2 – Thu gom bưu gửi đi liên tỉnh"
    ];
  }
  return Array.from(chks).map(c => c.value);
}

function generateMonthRanges(fromMonthStr, toMonthStr) {
  if (!fromMonthStr || !toMonthStr) {
    throw new Error("Vui lòng chọn đầy đủ Từ tháng và Đến tháng.");
  }
  const [y1, m1] = fromMonthStr.split("-").map(Number);
  const [y2, m2] = toMonthStr.split("-").map(Number);

  const startYearMonth = y1 * 12 + (m1 - 1);
  const endYearMonth = y2 * 12 + (m2 - 1);

  if (startYearMonth > endYearMonth) {
    throw new Error("Từ tháng không được lớn hơn Đến tháng.");
  }
  if (endYearMonth - startYearMonth >= 12) {
    throw new Error("Khoảng thời gian so sánh không được vượt quá 12 tháng.");
  }

  const list = [];
  for (let ym = startYearMonth; ym <= endYearMonth; ym++) {
    const year = Math.floor(ym / 12);
    const month = (ym % 12) + 1;
    const monthStr = pad(month);

    const fromDate = new Date(year, month - 1, 1);
    const toDate = new Date(year, month, 0);

    list.push({
      year,
      month,
      monthKey: `${monthStr}/${year}`,
      label: `Tháng ${monthStr}/${year}`,
      sheetName: `Thang_${monthStr}_${year}`,
      from: fromDate,
      to: toDate
    });
  }
  return list;
}

async function fetchMultiMonthV2Report(monthsList, tuyChonGR, selectedBcProvCode, selectedBcProvName) {
  const weights = readWeights();
  const results = [];

  for (let i = 0; i < monthsList.length; i++) {
    const monthObj = monthsList[i];
    setStatus(`Đang lấy dữ liệu V2 [Tháng ${i + 1}/${monthsList.length}]: ${monthObj.label}...`);
    const rows = await fetchV2ReportRows(
      tuyChonGR,
      monthObj.from,
      monthObj.to,
      weights,
      monthObj.label,
      selectedBcProvCode,
      selectedBcProvName
    );
    results.push({
      monthObj,
      rows
    });
  }

  return results;
}

function buildVisualTrendBar(diffPts) {
  if (diffPts == null || isNaN(diffPts)) return "⚪ 0.00%";
  const sign = diffPts >= 0 ? "+" : "";
  const valStr = `${sign}${diffPts.toFixed(2)}%`;
  if (diffPts >= 3.0) return `🟢▲▲ ${valStr}`;
  if (diffPts > 0) return `🟢▲ ${valStr}`;
  if (diffPts <= -3.0) return `🔴▼▼ ${valStr}`;
  if (diffPts < 0) return `🔴▼ ${valStr}`;
  return `⚪ ${valStr}`;
}

const MULTI_MONTH_METRICS_MAP = {
  "F1.1 – Nội tỉnh": { shortName: "F1.1", totalCol: 21, onTimeCol: 23 },
  "F1.2 – Thu gom bưu gửi đi liên tỉnh": { shortName: "F1.2", totalCol: 26, onTimeCol: 28 },
  "F1.3 – Chất lượng phát bưu gửi liên tỉnh": { shortName: "F1.3", totalCol: 31, onTimeCol: 32 },
  "F4.1 – Chất lượng phát thành công tại bưu cục": { shortName: "F4.1", totalCol: 17, onTimeCol: 18 }
};

function renderV2MultiMonthDashboardWorksheet(monthlyResults, monthsList, tuyChonGR) {
  const provCode = getSelectedProvinceCode();
  const provName = getSelectedProvinceName();
  const fromLabel = monthsList[0]?.label || "";
  const toLabel = monthsList[monthsList.length - 1]?.label || "";
  const totalMonths = monthsList.length;

  const selectedLabels = getSelectedMetricLabels();
  const activeMetrics = selectedLabels
    .map(label => MULTI_MONTH_METRICS_MAP[label])
    .filter(Boolean);

  if (activeMetrics.length === 0) {
    activeMetrics.push(MULTI_MONTH_METRICS_MAP["F4.1 – Chất lượng phát thành công tại bưu cục"]);
    activeMetrics.push(MULTI_MONTH_METRICS_MAP["F1.3 – Chất lượng phát bưu gửi liên tỉnh"]);
  }

  const unitMap = new Map();

  monthlyResults.forEach(({ monthObj, rows }, monthIdx) => {
    (rows || []).forEach(row => {
      const code = String(row[PROVINCE_CODE_INDEX]);
      const name = row[PROVINCE_NAME_INDEX];
      if (!unitMap.has(code)) {
        unitMap.set(code, {
          code,
          name,
          monthRows: Array(totalMonths).fill(null)
        });
      }
      unitMap.get(code).monthRows[monthIdx] = row;
    });
  });

  const units = Array.from(unitMap.values()).sort((a, b) => compareNumberOrText(a.code, b.code));

  const unitSummaries = units.map((u) => {
    const isTargetUnit = String(u.code) === String(provCode);

    const metricStats = activeMetrics.map((m) => {
      let sumVol = 0;
      let sumOnTime = 0;
      const rates = [];

      u.monthRows.forEach(r => {
        if (r) {
          const v = toNumberValue(r[m.totalCol]);
          const o = toNumberValue(r[m.onTimeCol]);
          const rate = v > 0 ? o / v : 0;
          sumVol += v;
          sumOnTime += o;
          rates.push(rate);
        } else {
          rates.push(0);
        }
      });

      const cumRate = sumVol > 0 ? sumOnTime / sumVol : 0;
      const firstRate = rates[0] || 0;
      const lastRate = rates[totalMonths - 1] || 0;
      const diffRate = (lastRate - firstRate) * 100;

      return { rates, sumVol, sumOnTime, cumRate, diffRate, rank: 0 };
    });

    const minDiff = metricStats.length ? Math.min(...metricStats.map(s => s.diffRate)) : 0;
    const minCum = metricStats.length ? Math.min(...metricStats.map(s => s.cumRate)) : 0;

    let segment = "⚪ ỔN ĐỊNH DUY TRÌ";
    let segStyle = isTargetUnit ? "RankTop" : "Same";
    if (minDiff >= 2.0) {
      segment = "🟢 TĂNG TRƯỞNG MẠNH (MoM >= +2%)";
      segStyle = "RankUp";
    } else if (minDiff < -2.0) {
      segment = "🔴 CẢNH BÁO SUY GIẢM (MoM < -2%)";
      segStyle = "RankDown";
    } else if (minCum >= 0.95) {
      segment = "🥇 TOP CHẤT LƯỢNG CAO (>=95%)";
      segStyle = "RankTop";
    }

    return {
      code: u.code,
      name: u.name,
      isTargetUnit,
      metricStats,
      segment,
      segStyle
    };
  });

  activeMetrics.forEach((m, mIdx) => {
    const sorted = [...unitSummaries]
      .filter(u => u.metricStats[mIdx].sumVol > 0)
      .sort((a, b) => b.metricStats[mIdx].cumRate - a.metricStats[mIdx].cumRate);
    sorted.forEach((u, rankIdx) => {
      u.metricStats[mIdx].rank = rankIdx + 1;
    });
  });

  const focusSummary = unitSummaries.find(u => u.isTargetUnit) || unitSummaries[0];

  // -------------------------------------------------------------
  // TAB 1: EXECUTIVE DASHBOARD FOR FOCUS TARGET UNIT (HUẾ)
  // Clean 13-column grid (fits 100% on laptop/desktop screen)
  // -------------------------------------------------------------
  const dashColWidths = [40, 220];
  monthsList.forEach(() => dashColWidths.push(80));
  dashColWidths.push(95, 120, 110);
  const dashTotalCols = dashColWidths.length;

  const widgets = activeMetrics.map((m, mIdx) => {
    const ms = focusSummary ? focusSummary.metricStats[mIdx] : { sumVol: 0, cumRate: 0, diffRate: 0, rank: 0 };
    const rankText = ms.rank > 0 ? `Hạng ${ms.rank}/${units.length} toàn quốc` : "";
    const diffSign = ms.diffRate >= 0 ? `+${ms.diffRate.toFixed(2)}%` : `${ms.diffRate.toFixed(2)}%`;

    return {
      title: `${mIdx + 1}. CHẤT LƯỢNG BQ (${m.shortName})`,
      value: `${(ms.cumRate * 100).toFixed(2)}%\nMoM: ${diffSign} | ${rankText}\n(SL: ${ms.sumVol.toLocaleString()} bg)`,
      styleTitle: `ProvCard${mIdx % 4 + 1}Title`,
      styleVal: `ProvCard${mIdx % 4 + 1}Val`
    };
  });

  const widgetSpanList = [
    { start: 0, span: 3 },
    { start: 3, span: 3 },
    { start: 6, span: 3 },
    { start: 9, span: dashTotalCols - 9 }
  ];

  const widgetTitleRow = `<Row ss:Height="22">${widgets.map((w, idx) => {
    const sp = widgetSpanList[idx] || { span: 3 };
    return `<Cell ss:MergeAcross="${sp.span - 1}" ss:StyleID="${w.styleTitle}"><Data ss:Type="String">${escapeXml(w.title)}</Data></Cell>`;
  }).join("")}</Row>`;

  const widgetValRow = `<Row ss:Height="56">${widgets.map((w, idx) => {
    const sp = widgetSpanList[idx] || { span: 3 };
    return `<Cell ss:MergeAcross="${sp.span - 1}" ss:StyleID="${w.styleVal}"><Data ss:Type="String">${escapeXml(w.value).replace(/\n/g, '&#10;')}</Data></Cell>`;
  }).join("")}</Row>`;

  const focusTrendRowsXml = activeMetrics.map((m, mIdx) => {
    const ms = focusSummary ? focusSummary.metricStats[mIdx] : { rates: [], cumRate: 0, diffRate: 0, rank: 0 };
    const fullName = m.shortName === 'F4.1' ? 'Chất lượng phát thành công tại bưu cục'
                   : m.shortName === 'F1.3' ? 'Chất lượng phát bưu gửi liên tỉnh'
                   : m.shortName === 'F1.1' ? 'Chất lượng phát bưu gửi nội tỉnh'
                   : 'Chất lượng thu gom bưu gửi đi liên tỉnh';
    return `
      <Row ss:Height="22">
        ${excelXmlCell(mIdx + 1, "Number", "Number")}
        ${excelXmlCell(`${m.shortName} – ${fullName}`, "String", "Text")}
        ${ms.rates.map(r => excelXmlCell(r, "Number", "Percent")).join("")}
        ${excelXmlCell(ms.cumRate, "Number", "Percent")}
        ${excelXmlCell(buildVisualTrendBar(ms.diffRate), "String", trendStyle(ms.diffRate))}
        ${excelXmlCell(ms.rank > 0 ? `${ms.rank} / ${units.length}` : "", "String", "Text")}
      </Row>`;
  }).join("");

  const notesTextLines = [
    `🏆 BÁO CÁO LŨY KẾ & MA TRẬN SO SÁNH CHUỖI XU HƯỚNG ĐA THÁNG (EXECUTIVE BI DASHBOARD):`,
    `1. ĐƠN VỊ TRỌNG ĐIỂM QUAN SÁT: ${provName} (Mã ${provCode}) | KHOẢNG THỜI GIAN: ${fromLabel} ➔ ${toLabel} (${totalMonths} tháng liên tục).`
  ];
  activeMetrics.forEach((m, mIdx) => {
    const ms = focusSummary ? focusSummary.metricStats[mIdx] : { sumVol: 0, cumRate: 0, rank: 0 };
    notesTextLines.push(`${mIdx + 2}. TIÊU ĐIỂM ${m.shortName} (${provName}): Sản lượng lũy kế ${ms.sumVol.toLocaleString()} bưu gửi, Tỷ lệ bình quân ${(ms.cumRate * 100).toFixed(2)}% (Xếp hạng ${ms.rank}/${units.length} toàn quốc).`);
  });
  notesTextLines.push(`${activeMetrics.length + 2}. 🎯 CHỈ ĐẠO ĐIỀU HÀNH BAN GIÁM ĐỐC: Đề nghị các đơn vị có biến động suy giảm (🔴 MoM < -2%) rà soát lại các công đoạn để cải thiện chất lượng.`);
  const notesText = notesTextLines.join("\n");

  const executiveDashboardSheet = `<Worksheet ss:Name="Dashboard_${escapeXml(provName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ''))}">
    <Table>
      ${renderColumnWidths(dashColWidths)}
      <Row ss:Height="26"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Title"><Data ss:Type="String">🏆 BAN GIÁM ĐỐC BI DASHBOARD - TỔNG QUAN CHẤT LƯỢNG KÉP &amp; CHUỖI XU HƯỚNG BƯU ĐIỆN TỈNH ${escapeXml(provName.toUpperCase())} (MÃ ${provCode})</Data></Cell></Row>
      <Row ss:Height="18"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Text"><Data ss:Type="String">Đơn vị tiêu điểm: ${escapeXml(provName)} (Mã ${provCode})  |  Theo dõi chuỗi xu hướng ${totalMonths} tháng liên tục (${escapeXml(fromLabel)} ➔ ${escapeXml(toLabel)})  |  Đối sánh ${units.length} BĐT/TP toàn quốc</Data></Cell></Row>
      <Row></Row>

      <!-- KHỐI 1: 04 KPI CARDS TIÊU ĐIỂM QUẢN TRỊ RIÊNG CHO HUẾ -->
      ${widgetTitleRow}
      ${widgetValRow}
      <Row></Row>

      <!-- KHỐI 2: BẢNG CHUỖI XU HƯỚNG CHẤT LƯỢNG HÀNG THÁNG RIÊNG CHO HUẾ -->
      <Row ss:Height="22"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Group"><Data ss:Type="String">📈 BẢNG CHUỖI XU HƯỚNG CHẤT LƯỢNG TỪNG THÁNG (MONTH-BY-MONTH TREND) RIÊNG CHO ${escapeXml(provName.toUpperCase())}</Data></Cell></Row>
      <Row ss:Height="24">
        ${excelXmlCell("STT", "String", "Header")}
        ${excelXmlCell("Chỉ Tiêu Chất Lượng", "String", "Header")}
        ${monthsList.map(m => excelXmlCell(m.monthKey, "String", "Header")).join("")}
        ${excelXmlCell("BQ Lũy Kế", "String", "Header")}
        ${excelXmlCell(`MoM (${monthsList[totalMonths - 1]?.monthKey} vs ${monthsList[0]?.monthKey})`, "String", "Header")}
        ${excelXmlCell("Hạng Toàn Quốc", "String", "Header")}
      </Row>
      ${focusTrendRowsXml}
      <Row></Row>

      <!-- KHỐI 3: PHÂN TÍCH QUẢN TRỊ & CHỈ ĐẠO ĐIỀU HÀNH -->
      <Row ss:Height="22"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="NoteHeader"><Data ss:Type="String">📝 BẢN PHÂN TÍCH TIÊU ĐIỂM BI &amp; CHỈ ĐẠO ĐIỀU HÀNH - ${escapeXml(provName.toUpperCase())}</Data></Cell></Row>
      <Row ss:Height="80"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="NoteText"><Data ss:Type="String">${escapeXml(notesText).replace(/\n/g, '&#10;')}</Data></Cell></Row>
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>5</SplitHorizontal><TopRowBottomPane>5</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;

  // -------------------------------------------------------------
  // TAB 2: FULL BENCHMARK MATRIX FOR 34 PROVINCES TOÀN QUỐC
  // Full detailed matrix table
  // -------------------------------------------------------------
  const headerCols = [
    "STT", "Mã", "Tên Đơn Vị"
  ];
  activeMetrics.forEach(m => {
    monthsList.forEach(month => {
      headerCols.push(`${m.shortName} ${month.monthKey}`);
    });
    headerCols.push(`Lũy Kế ${m.shortName} BQ`, `MoM ${m.shortName} (Tn vs T1)`);
  });
  headerCols.push("Phân loại Vị thế BI");

  const colWidths = [40, 60, 190];
  activeMetrics.forEach(() => {
    monthsList.forEach(() => colWidths.push(85));
    colWidths.push(100, 130);
  });
  colWidths.push(180);
  const totalCols = headerCols.length;

  const metricShortNames = activeMetrics.map(m => m.shortName).join(" &amp; ");

  const unitMatrixRowsXml = unitSummaries.map((u, idx) => {
    const rowStyle = u.isTargetUnit ? "RankTop" : "";
    const displayName = u.isTargetUnit ? `${u.name} ★` : u.name;

    let rowXml = `
      <Row ss:Height="20">
        ${excelXmlCell(idx + 1, "Number", rowStyle || "Number")}
        ${excelXmlCell(u.code, "String", rowStyle || "Text")}
        ${excelXmlCell(displayName, "String", rowStyle || "Text")}`;

    u.metricStats.forEach(ms => {
      rowXml += ms.rates.map(r => excelXmlCell(r, "Number", rowStyle || "Percent")).join("");
      rowXml += excelXmlCell(ms.cumRate, "Number", rowStyle || "Percent");
      rowXml += excelXmlCell(buildVisualTrendBar(ms.diffRate), "String", trendStyle(ms.diffRate));
    });

    rowXml += `
        ${excelXmlCell(u.segment, "String", u.segStyle)}
      </Row>`;
    return rowXml;
  }).join("");

  const benchmarkMatrixSheet = `<Worksheet ss:Name="MaTran_34_Tinh_ToanQuoc">
    <Table>
      ${renderColumnWidths(colWidths)}
      <Row ss:Height="24"><Cell ss:MergeAcross="${totalCols - 1}" ss:StyleID="Group"><Data ss:Type="String">📊 BẢNG MA TRẬN ĐỐI SÁNH &amp; XẾP HẠNG 34 BƯU ĐIỆN TỈNH/TP TOÀN QUỐC (${metricShortNames}) (${totalMonths} THÁNG - HIGHLIGHT DÒNG HUẾ ★)</Data></Cell></Row>
      <Row ss:Height="24">${headerCols.map(c => excelXmlCell(c, "String", "Header")).join("")}</Row>
      ${unitMatrixRowsXml}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>2</SplitHorizontal><TopRowBottomPane>2</TopRowBottomPane><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane><ActivePane>0</ActivePane></WorksheetOptions>
  </Worksheet>`;

  return `${executiveDashboardSheet}\n${benchmarkMatrixSheet}`;
}

async function exportReportV2MultiMonth(tuyChonGR) {
  const fromMonthInput = document.getElementById("from-month");
  const toMonthInput = document.getElementById("to-month");
  if (!fromMonthInput || !toMonthInput || !fromMonthInput.value || !toMonthInput.value) {
    setStatus("Vui lòng chọn Từ tháng và Đến tháng.", "error");
    return;
  }

  let monthsList;
  try {
    monthsList = generateMonthRanges(fromMonthInput.value, toMonthInput.value);
  } catch (err) {
    setStatus(err.message, "error");
    return;
  }

  setBusy(true);
  const provCode = getSelectedProvinceCode();
  const provName = getSelectedProvinceName();
  setStatus(`Bắt đầu tải dữ liệu lũy kế ${monthsList.length} tháng cho V2 (${tuyChonGR === "TINH" ? "34 Tỉnh/TP" : "Bưu cục thuộc " + provName})...`);

  try {
    const selectedBcProvCode = tuyChonGR === "BC" ? provCode : null;
    const selectedBcProvName = tuyChonGR === "BC" ? provName : null;

    const monthlyResults = await fetchMultiMonthV2Report(monthsList, tuyChonGR, selectedBcProvCode, selectedBcProvName);

    if (!monthlyResults || !monthlyResults.length) {
      throw new Error("Không lấy được dữ liệu cho các tháng đã chọn.");
    }

    setStatus("Đang khởi tạo Dashboard Ma trận Lũy kế Đa tháng XML...");
    const dashboardSheetXml = renderV2MultiMonthDashboardWorksheet(monthlyResults, monthsList, tuyChonGR);

    const dataSheetsXml = monthlyResults.map(({ monthObj, rows }) => {
      return renderV2DataWorksheet(monthObj.sheetName, rows, `Dữ liệu V2 ${monthObj.label} (${tuyChonGR})`);
    }).join("\n");

    const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  ${renderWorkbookPropertiesAndStyles()}
  ${dashboardSheetXml}
  ${dataSheetsXml}
</Workbook>`;

    const fileName = `filebaocaov2_luyke_thang_${tuyChonGR === "TINH" ? "tinh" : "bc"}_${fromMonthInput.value}_vs_${toMonthInput.value}.xls`;
    const blob = new Blob([xmlContent], { type: "application/vnd.ms-excel;charset=utf-8" });
    saveBlobAsFile(blob, fileName);

    renderV2MultiMonthPreview(monthlyResults, monthsList, tuyChonGR);
    setStatus(`Hoàn tất xuất Báo cáo Lũy kế Ma trận ${monthsList.length} tháng (${tuyChonGR}): Đã tải ${fileName}.`, "ok");
  } catch (error) {
    console.error("[DKCL][MultiMonth] Error:", error);
    setStatus(error.message || "Có lỗi xảy ra khi xuất báo cáo lũy kế tháng.", "error");
  } finally {
    setBusy(false);
  }
}

function renderV2MultiMonthPreview(monthlyResults, monthsList, tuyChonGR) {
  if (reportPanel) reportPanel.classList.remove("is-empty");
  if (emptyState) emptyState.hidden = true;
  if (reportBadge) reportBadge.textContent = "Multi-Month BI Report";

  const totalUnits = monthlyResults[0]?.rows?.length || 0;
  if (reportSummary) {
    reportSummary.hidden = false;
    reportSummary.innerHTML = `
      <div class="metric"><span>Chế độ báo cáo</span><strong>Lũy kế đa tháng (V2)</strong></div>
      <div class="metric"><span>Số tháng so sánh</span><strong>${monthsList.length} tháng</strong></div>
      <div class="metric"><span>Đơn vị tổng hợp</span><strong>${tuyChonGR === "BC" ? "Bưu cục" : "34 Tỉnh/TP"} (${totalUnits} đơn vị)</strong></div>
      <div class="metric"><span>Khoảng tháng</span><strong>${monthsList[0].monthKey} ➔ ${monthsList[monthsList.length - 1].monthKey}</strong></div>
    `;
  }
}
