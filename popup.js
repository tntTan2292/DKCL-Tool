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
  { key: "phatThanhCongBuuCuc", name: "Chất lượng phát thành công tại bưu cục (F4.1)", status: "Đang gọi curl V2 F4.1...", buildParams: buildPhatThanhCongV2Params, codeIndex: 1, nameIndex: 2, bcCodeIndex: 5, bcNameIndex: 6, values: { total: 10, onTime: 27 }, bcValues: { total: 10, onTime: 27 }, targets: { total: 17, onTime: 18, rate: 19, rank: 20 }, denominatorField: "total" },
  { key: "noiTinhF11", name: "Nội tỉnh F1.1", status: "Đang gọi curl V2 Nội tỉnh F1.1...", buildParams: buildNoiTinhF11Params, codeIndex: 1, nameIndex: 2, bcCodeIndex: 3, bcNameIndex: 4, values: { total: 11, subTotal: 18, onTime: 19 }, bcValues: { total: 13, subTotal: 20, onTime: 21 }, targets: { total: 21, subTotal: 22, onTime: 23, rate: 24, rank: 25 }, denominatorField: "subTotal" },
  { key: "thuGomLienTinhF12", name: "Thu gom bưu gửi đi liên tỉnh (F1.2)", status: "Đang gọi curl V2 Thu gom liên tỉnh F1.2...", buildParams: buildThuGomLienTinhF12Params, codeIndex: 1, nameIndex: 2, bcCodeIndex: 3, bcNameIndex: 4, values: { total: 7, subTotal: 8, onTime: 9 }, bcValues: { total: 9, subTotal: 10, onTime: 11 }, targets: { total: 26, subTotal: 27, onTime: 28, rate: 29, rank: 30 }, denominatorField: "total" },
  { key: "phatLienTinhF13", name: "Chất lượng phát bưu gửi liên tỉnh (F1.3)", status: "Đang gọi curl V2 Phát liên tỉnh F1.3...", buildParams: buildPhatLienTinhF13Params, codeIndex: 1, nameIndex: 2, bcCodeIndex: 3, bcNameIndex: 4, values: { total: 10, onTime: 18 }, bcValues: { total: 10, onTime: 18 }, targets: { total: 31, onTime: 32, rate: 33, rank: 34 }, denominatorField: "total" }
];

const PROVINCE_CODE_INDEX = 1;
const PROVINCE_NAME_INDEX = 2;
const PROVINCE_NAMES = Object.freeze({
  10: "TP. Hà Nội", 16: "Hưng Yên", 18: "TP. Hải Phòng", 20: "Quảng Ninh", 22: "Bắc Ninh",
  24: "Lạng Sơn", 25: "Thái Nguyên", 27: "Cao Bằng", 29: "Phú Thọ", 30: "Tuyên Quang",
  33: "Lào Cai", 36: "Sơn La", 38: "Điện Biên", 39: "Lai Châu", 43: "Ninh Bình",
  44: "Thanh Hóa", 46: "Nghệ An", 48: "Hà Tĩnh", 52: "Quảng Trị", 53: "Thừa Thiên Huế",
  55: "TP. Đà Nẵng", 57: "Quảng Ngãi", 60: "Gia Lai", 63: "Đắc Lắk", 65: "Khánh Hòa",
  67: "Lâm Đồng", 70: "TP. Hồ Chí Minh", 81: "Đồng Nai", 84: "Tây Ninh", 87: "Đồng Tháp",
  88: "An Giang", 89: "Vĩnh Long", 90: "TP. Cần Thơ", 97: "Cà Mau"
});
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
const NATIVE_XLSX_FEATURE_FLAG_KEY = "dkclNativeXlsxExportEnabled";

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

function isNativeV2XlsxExportEnabled() {
  return localStorage.getItem(NATIVE_XLSX_FEATURE_FLAG_KEY) !== "false";
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
      const comparisonRows = buildV2ComparisonRows(currentRows, compareRows, tuyChonGR);
      const provinceComparison = tuyChonGR === "TINH"
        ? comparisonRows.find((row) => String(row.code) === String(getSelectedProvinceCode()))
        : null;
      const suffix = tuyChonGR === "TINH" ? "tinh" : "bc";
      const fileName = `filebaocaov2_so_sanh_ky_${suffix}_${toFileDate(from)}_${toFileDate(to)}_vs_${toFileDate(compareFrom)}_${toFileDate(compareTo)}.xls`;
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

  // BC only has the "phát thành công tại bưu cục" legacy sources; the others are
  // province-level and would join nothing.
  const configs = tuyChonGR === "BC"
    ? SOURCE_CONFIGS.filter((config) => config.key === "phatThanhCongBuuCuc")
    : SOURCE_CONFIGS;

  for (const config of configs) {
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

// BC reports only ever carried F4.1 and F1.3 in the original extension: F1.1 (nội tỉnh) and
// F1.2 (thu gom liên tỉnh) are province-level indicators with no post-office breakdown, so
// fetching them for BC would only manufacture rows the legacy report never had.
const V2_BC_SOURCE_KEYS = new Set(["phatThanhCongBuuCuc", "phatLienTinhF13"]);
const V2_BC_SHORT_NAMES = ["F4.1", "F1.3"];

// Both the source list and the dashboard metric list must agree on what BC supports.
// Selecting only F1.1/F1.2 for BC is refused outright rather than answered with zero rows.
function assertBcKpiSelection(activeCount) {
  if (activeCount > 0) return;
  throw new Error(
    `Cấp Bưu cục chỉ hỗ trợ ${V2_BC_SHORT_NAMES.join(" và ")} theo Extension gốc. ` +
    "F1.1 và F1.2 là chỉ tiêu cấp tỉnh, không có chia nhỏ theo bưu cục nên không hỗ trợ. " +
    "Vui lòng chọn ít nhất 1 chỉ tiêu chất lượng cấp Bưu cục."
  );
}

function getActiveV2SourceConfigs(tuyChonGR) {
  const selected = getSelectedMetricLabels();
  const scoped = tuyChonGR === "BC"
    ? V2_SOURCE_CONFIGS.filter((config) => V2_BC_SOURCE_KEYS.has(config.key))
    : V2_SOURCE_CONFIGS;
  const active = scoped.filter((config) => {
    if (config.key === "phatThanhCongBuuCuc") return selected.includes("F4.1 – Chất lượng phát thành công tại bưu cục");
    if (config.key === "noiTinhF11") return selected.includes("F1.1 – Nội tỉnh");
    if (config.key === "thuGomLienTinhF12") return selected.includes("F1.2 – Thu gom bưu gửi đi liên tỉnh");
    if (config.key === "phatLienTinhF13") return selected.includes("F1.3 – Chất lượng phát bưu gửi liên tỉnh");
    return true;
  });
  if (tuyChonGR === "BC") assertBcKpiSelection(active.length);
  return active;
}

async function fetchV2ReportRows(tuyChonGR, from, to, weights, periodLabel = "", selectedBcProvCode = null, selectedBcProvName = null, options = {}) {
  let baseRows = [];
  let selectedProvinceCodeInt = NaN;
  let selectedProvinceName = "";

  if (tuyChonGR === "TINH") {
    selectedProvinceCodeInt = toIntegerCode(getSelectedProvinceCode());
    selectedProvinceName = getSelectedProvinceName() || PROVINCE_NAMES[selectedProvinceCodeInt] || "Không xác định";
    if (!Object.prototype.hasOwnProperty.call(PROVINCE_NAMES, selectedProvinceCodeInt)) {
      throw new Error(`Tỉnh tiêu điểm mã ${getSelectedProvinceCode() || "trống"} (${selectedProvinceName}) không thuộc danh mục 34 tỉnh/TP.`);
    }

    baseRows = Object.entries(PROVINCE_NAMES).map(([code, name], idx) => {
      const v2Row = Array(V2_EXCEL_COLUMNS.length).fill("");
      v2Row[0] = String(idx + 1);
      v2Row[PROVINCE_CODE_INDEX] = code;
      v2Row[PROVINCE_NAME_INDEX] = name;
      return v2Row;
    });
  } else {
    // BC keeps the legacy seeding: the post-office catalogue (code, name and the V1
    // PTC TMĐT / Truyền thống columns) comes from the combined report, and the V2 sources
    // are joined onto it. Without this seed the BC report has no unit list of its own.
    const combinedRows = await fetchCombinedReport(tuyChonGR, from, to, periodLabel);
    const finalizedBaseRows = finalizeRows(combinedRows, weights, tuyChonGR);
    if (!finalizedBaseRows.length) {
      throw new Error(buildMissingSourceDataMessage(SOURCE_CONFIGS.map((config) => config.name).join(" / "), periodLabel, from, to));
    }
    baseRows = finalizedBaseRows.map((row) => {
      const v2Row = Array(V2_EXCEL_COLUMNS.length).fill("");
      TEMPLATE_EXCEL_COLUMNS.forEach((_, index) => (v2Row[index] = row[index] ?? ""));
      return v2Row;
    });
  }

  const activeConfigs = getActiveV2SourceConfigs(tuyChonGR);
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
    const integrityContext = { periodLabel, requestUrl, headerCells: records.headerCells };
    records.forEach((record) => assertV2RecordIntegrity(record, config, integrityContext));
    const aggregatedRecords = aggregateV2RecordsByCode(records, config);
    console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: parsed records`, records.length, records.slice(0, 3));
    console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: aggregated records before join`, aggregatedRecords.length, aggregatedRecords.slice(0, 5));

    if (tuyChonGR === "TINH") {
      const hasSelectedProvinceRecord = records.some((record) => record.provinceCodeInt === selectedProvinceCodeInt);
      if (!hasSelectedProvinceRecord && !options.allowMissingSelectedProvince) {
        throw new Error(`Biểu "${config.name}" không join được tỉnh tiêu điểm mã ${selectedProvinceCodeInt} (${selectedProvinceName}) cho ${periodLabel || "kỳ báo cáo"} (${toApiDate(from)} - ${toApiDate(to)}). Dừng xuất để không thay số liệu bằng 0.\nLink curl: ${requestUrl}`);
      }
    } else if (!aggregatedRecords.length) {
      throw new Error(buildMissingSourceDataMessage(config.name, periodLabel, from, to, requestUrl));
    }

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
          if (tuyChonGR === "TINH") continue;
          row = Array(V2_EXCEL_COLUMNS.length).fill("");
          row[0] = String(baseRows.length + 1);
          row[PROVINCE_CODE_INDEX] = String(record.provinceCodeInt);
          row[PROVINCE_NAME_INDEX] = record.provinceName || `Đơn vị ${record.provinceCodeInt}`;
          baseRows.push(row);
          rowMap.set(String(record.provinceCodeInt), row);
        }
        assertV2RecordIntegrity(record, config, integrityContext);
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

  // BC keeps only operating post offices (BCVH) that actually carry data, exactly as the
  // legacy report did. TINH never drops a row for a zero volume — there a 0 is real data
  // and must stay visible rather than vanish from the report.
  const finalRows = tuyChonGR === "BC"
    ? baseRows.filter((row) => isBcOperatingRowWithData(row, activeConfigs))
    : baseRows;

  finalizeV2Rows(finalRows, tuyChonGR);
  console.log(`[DKCL][V2][${tuyChonGR}] final rows after V2 join`, finalRows.length, finalRows.slice(0, 10));
  return finalRows;
}

// BC-only row filter, restored from the original extension. Kept out of fetchV2ReportRows
// so the province path provably never drops a row for having a zero volume.
function isBcOperatingRowWithData(row, activeConfigs) {
  const isOperatingPostOffice = String(row[PROVINCE_NAME_INDEX] || "").toUpperCase().includes("BCVH");
  const hasData = activeConfigs.some((config) => toNumberValue(row[config.targets.total]) > 0);
  return isOperatingPostOffice && hasData;
}

function parseV2ApiRows(html, config, tuyChonGR, joinCodeSet = null) {
  const documentHtml = new DOMParser().parseFromString(`<table><tbody>${html}</tbody></table>`, "text/html");
  const rows = [...documentHtml.querySelectorAll("tr")]
    .filter((tr) => tr.children.length > 2 && !tr.classList.contains("tr_tong"));
  const cellRows = rows.map((tr) => [...tr.children].map((td) => cleanText(td.textContent)));
  console.log(`[DKCL][V2][${tuyChonGR}] ${config.name}: raw rows`, rows.length, "sample cells", cellRows.slice(0, 10));

  const records = cellRows
    .map((cells) => {
      const provinceCodeText = pickV2JoinCodeText(cells, config, tuyChonGR, joinCodeSet);
      const provinceCodeInt = toIntegerCode(provinceCodeText);
      const provinceName = pickReportName(cells, config, tuyChonGR);
      const valuesConfig = tuyChonGR === "BC" && config.bcValues ? config.bcValues : config.values;
      return {
        cells,
        provinceCodeText,
        provinceCodeInt,
        provinceName,
        total: cells[valuesConfig.total] || "",
        subTotal: Number.isInteger(valuesConfig.subTotal) ? cells[valuesConfig.subTotal] || "" : "",
        onTime: cells[valuesConfig.onTime] || ""
      };
    })
    .filter((record) => Number.isInteger(record.provinceCodeInt));

  // The golden extension defines positional offsets for this upstream table layout.
  // Preserve a non-data row as context when an invariant fails, but never use it to
  // infer or shift an offset at runtime.
  records.headerCells = cellRows.find((cells) => !Number.isInteger(toIntegerCode(pickV2JoinCodeText(cells, config, tuyChonGR, joinCodeSet)))) || null;
  return records;
}

// V2_SOURCE_CONFIGS is the golden mapping for the upstream fixed table layout. A changed
// layout must fail loudly; validation is a safety net, never evidence for selecting cells.
function assertV2RecordIntegrity(record, config, context) {
  const total = readV2IntegrityNumber(record.total, "total", record, config, context);
  const onTime = readV2IntegrityNumber(record.onTime, "onTime", record, config, context);
  const hasSubTotal = Number.isInteger(config.values.subTotal);
  const subTotal = hasSubTotal ? readV2IntegrityNumber(record.subTotal, "subTotal", record, config, context) : null;
  const denominatorField = config.denominatorField || "total";
  const denominator = denominatorField === "subTotal" ? subTotal : total;

  let violation = "";
  if (onTime > 0 && denominator <= 0) {
    violation = `mẫu số ${denominatorField} = ${denominator} trong khi onTime = ${onTime}`;
  } else if (onTime > denominator) {
    violation = `onTime = ${onTime} lớn hơn mẫu số ${denominatorField} = ${denominator}`;
  }
  if (!violation) return;

  throw new Error(buildV2MappingErrorMessage(violation, record, config, context));
}

function readV2IntegrityNumber(value, field, record, config, context) {
  const cleaned = cleanNumber(value);
  let violation = "";
  if (cleaned === "") violation = `cột ${field} trống`;
  else if (!Number.isFinite(Number(cleaned))) violation = `cột ${field} không phải số hợp lệ (${JSON.stringify(value)})`;
  else if (Number(cleaned) < 0) violation = `cột ${field} âm (${cleaned})`;
  if (violation) throw new Error(buildV2MappingErrorMessage(violation, record, config, context));
  return Number(cleaned);
}

// Last line of defence, expressed against the sheet columns the Dashboard and DuLieu_Chart
// read: nothing may reach the workbook with a numerator larger than its denominator.
function assertV2SheetRowIntegrity(row, config) {
  const onTime = toNumberValue(row[config.targets.onTime]);
  const denominatorField = config.denominatorField || "total";
  const denominatorTarget = config.targets[denominatorField];
  const denominator = toNumberValue(row[denominatorTarget]);

  let violation = "";
  if (onTime > 0 && denominator <= 0) {
    violation = `mẫu số ${denominatorField} cột ${denominatorTarget} = ${denominator} trong khi đúng hạn cột ${config.targets.onTime} = ${onTime}`;
  } else if (onTime > denominator) {
    violation = `đúng hạn cột ${config.targets.onTime} = ${onTime} lớn hơn mẫu số ${denominatorField} cột ${denominatorTarget} = ${denominator}`;
  }
  if (!violation) return;

  throw new Error([
    `Dòng báo cáo V2 vi phạm bất biến ở biểu "${config.name}": ${violation}.`,
    `Đơn vị: mã ${row[PROVINCE_CODE_INDEX]} (${row[PROVINCE_NAME_INDEX] || "?"}).`,
    "Dừng xuất: không ghi số 0 giả vào sheet tháng / DuLieu_Chart / Dashboard."
  ].join("\n"));
}

function buildV2MappingErrorMessage(violation, record, config, context = {}) {
  const period = context.periodLabel ? ` cho ${context.periodLabel}` : "";
  const lines = [
    `Dữ liệu nguồn không thỏa invariant ở biểu "${config.name}"${period}: ${violation}.`,
    `Đơn vị: mã ${record.provinceCodeInt} (${record.provinceName || "?"}).`,
    `Offset đang dùng: ${JSON.stringify(config.values)}.`,
    `Ô đọc được: ${JSON.stringify(record.cells)}.`
  ];
  if (context.headerCells) lines.push(`Header nguồn: ${JSON.stringify(context.headerCells)}.`);
  if (context.requestUrl) lines.push(`Link curl: ${context.requestUrl}`);
  lines.push("Dừng xuất để rà soát dữ liệu nguồn hoặc mapping đã được phê duyệt; không thay bằng 0.");
  return lines.join("\n");
}

function pickV2JoinCodeText(cells, config, tuyChonGR, joinCodeSet) {
  if (tuyChonGR === "BC" && joinCodeSet?.size) {
    // Skip the province column so the post-office code wins the join. The province is read
    // from the dropdown, never hardcoded.
    const provinceCode = String(toIntegerCode(getSelectedProvinceCode()));
    const matchedCode = cells
      .map((cell) => String(toIntegerCode(cell)))
      .find((code) => code !== provinceCode && joinCodeSet.has(code));
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

function finalizeV2Rows(rows, tuyChonGR) {
  const activeConfigs = getActiveV2SourceConfigs(tuyChonGR);
  activeConfigs.forEach((config) => {
    rows.forEach((row) => {
      const denominatorTarget = config.targets[config.denominatorField || "total"];
      const totalRaw = row[config.targets.total];
      const denominatorRaw = row[denominatorTarget];
      const numeratorRaw = row[config.targets.onTime];
      // A source record with 0/0 is valid data. A blank cell means this province/KPI
      // has no valid record for this month and must stay N/A, rather than becoming 0.
      if (![totalRaw, denominatorRaw, numeratorRaw].every(hasV2SourceValue)) {
        row[config.targets.total] = "";
        row[config.targets.onTime] = "";
        if (Number.isInteger(config.targets.subTotal)) row[config.targets.subTotal] = "";
        row[config.targets.rate] = "";
        row[config.targets.rank] = "";
        return;
      }
      const total = zeroIfBlank(totalRaw);
      const denominator = zeroIfBlank(denominatorRaw);
      const numerator = zeroIfBlank(numeratorRaw);
      row[config.targets.total] = total;
      row[config.targets.onTime] = numerator;
      if (Number.isInteger(config.targets.subTotal)) row[config.targets.subTotal] = zeroIfBlank(row[config.targets.subTotal]);
      assertV2SheetRowIntegrity(row, config);
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
    
    // Pad empty spaces to avoid sparse arrays that skip columns during XML Generation
    rows.forEach((row) => {
      for (let i = 0; i < V2_EXCEL_COLUMNS.length; i++) {
        if (row[i] === undefined) row[i] = "";
      }
    });
  }

function getV2SelectedProvinceFilter(tuyChonGR) {
  if (tuyChonGR !== "BC") return "ALL";
  const selectedCode = String(getSelectedProvinceCode() || "").trim();
  return /^\d+$/.test(selectedCode) ? selectedCode : "ALL";
}

function hasV2SourceValue(value) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function buildPhatThanhCongV2Params(tuyChonGR, from, to) {
  return {
    TuyChonGR: tuyChonGR,
    stMaTinhPhat: getV2SelectedProvinceFilter(tuyChonGR),
    stMaLoaiBCPhat: "NULL",
    stMaBuuCucPhat: tuyChonGR === "BC" ? "ALL" : "NULL",
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
    stMaTinhChapNhan: getV2SelectedProvinceFilter(tuyChonGR),
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
    stMaTinhNhan: getV2SelectedProvinceFilter(tuyChonGR),
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
    stMaTinhPhat: getV2SelectedProvinceFilter(tuyChonGR),
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
    stMaTinhPhat: getV2SelectedProvinceFilter(tuyChonGR),
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
    stMaTinhThuGom: getV2SelectedProvinceFilter(tuyChonGR),
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

  const finalRows = tuyChonGR === "BC"
    ? rows.filter((row) => String(row[PROVINCE_NAME_INDEX]).toUpperCase().includes("BCVH"))
    : rows.filter(hasAllPositiveRateColumns);

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

  if (mode === "compare") {
    if (tuyChonGR === "TINH" && provinceComparison) {
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
    if (tuyChonGR === "BC") {
      reportText.textContent = buildBcComparisonText({ rows, from, to, compareFrom, compareTo });
      return;
    }
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
  ${tuyChonGR === "TINH" ? renderV2ComparisonDashboardWorksheet(currentRows, compareRows, tuyChonGR, from, to, compareFrom, compareTo) : ""}
  ${renderV2DataWorksheet("Du lieu dau ky V2", currentRows, `Dữ liệu V2 đầu kỳ: ${toApiDate(from)} - ${toApiDate(to)}`)}
  ${renderV2DataWorksheet("Ky so sanh V2", compareRows, `Dữ liệu V2 kỳ so sánh: ${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`)}
  ${tuyChonGR === "TINH"
    ? renderV2ComparisonWorksheet(comparisonRows, from, to, compareFrom, compareTo)
    : renderV2BcComparisonWorksheet(comparisonRows, from, to, compareFrom, compareTo)}
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
    <Style ss:ID="KpiGreen"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#155724"/><Interior ss:Color="#D4EDDA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="KpiYellow"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#856404"/><Interior ss:Color="#FFF3CD" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="KpiRed"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#721C24"/><Interior ss:Color="#F8D7DA" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="KpiGray"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#5A6268"/><Interior ss:Color="#E2E3E5" ss:Pattern="Solid"/><Borders>${excelXmlBorders()}</Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
    <Style ss:ID="KpiTargetLine"><Font ss:FontName="Arial" ss:Bold="1" ss:Color="#1F4E78"/><Interior ss:Color="#D9EAF7" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2" ss:Color="#1F4E78"/></Borders><NumberFormat ss:Format="0.00%"/><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/></Style>
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

// Restored verbatim from the original extension: the BC comparison sheet and its preview
// text. Dropping them made BC fall through to the province layout.
function renderV2BcComparisonWorksheet(rows, from, to, compareFrom, compareTo) {
  return `<Worksheet ss:Name="SoSanhV2_BC">
    <Table>
      ${renderColumnWidths([46, 78, 190, 210, 90, 190])}
      <Row><Cell ss:MergeAcross="5" ss:StyleID="Title"><Data ss:Type="String">📊 TTVH – Kết quả so sánh kỳ ${escapeXml(getSelectedProvinceName() || "")} (${rows.length} đơn vị trực thuộc)</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="5" ss:StyleID="Text"><Data ss:Type="String">Kỳ báo cáo: ${escapeXml(toApiDate(from))} - ${escapeXml(toApiDate(to))}</Data></Cell></Row>
      <Row><Cell ss:MergeAcross="5" ss:StyleID="Text"><Data ss:Type="String">So sánh với kỳ: ${escapeXml(toApiDate(compareFrom))} - ${escapeXml(toApiDate(compareTo))}</Data></Cell></Row>
      <Row>${["STT", "Mã", "Tên", "Chỉ số", "Tỷ lệ", "So sánh tỷ lệ"].map((column) => excelXmlCell(column, "String", "Header")).join("")}</Row>
      ${rows.map(renderBcComparisonXmlRows).join("")}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
  </Worksheet>`;
}

function renderBcComparisonXmlRows(row, index) {
  return row.metrics
    .map((metric) => {
      return `<Row>
        ${excelXmlCell(index + 1, "Number", "Number")}
        ${excelXmlCell(row.code, "String", "Text")}
        ${excelXmlCell(row.name, "String", "Text")}
        ${excelXmlCell(metric.label, "String", "Group")}
        ${excelXmlCell(metric.currentRate, "Number", "Percent")}
        ${excelXmlCell(rateTrendText(metric, { plain: true }), "String", trendStyle(metric.rateDiffPoints))}
      </Row>`;
    })
    .join("");
}

function buildBcComparisonText({ rows, from, to, compareFrom, compareTo }) {
  const periodText = `${toApiDate(from)} - ${toApiDate(to)}`;
  const comparePeriodText = `${toApiDate(compareFrom)} - ${toApiDate(compareTo)}`;

  const bcBlocks = rows.map((row) => {
    const metricBlocks = row.metrics.map((metric) => {
      return `  🔹 ${metric.label}\n  Tỷ lệ: ${formatPercentValue(metric.currentRate)} (${rateTrendText(metric)})`;
    });
    return `📍 ${row.code} - ${row.name}\n${metricBlocks.join("\n")}`;
  });

  return [
    `📊 Kết quả so sánh Bưu cục Vận hành (Tổng số: ${rows.length} bưu cục)`,
    `Kỳ báo cáo: ${periodText}`,
    `So sánh với kỳ: ${comparePeriodText}`,
    "",
    bcBlocks.join("\n\n")
  ].join("\n");
}

function buildV2ComparisonRows(currentRows, compareRows, tuyChonGR) {
  // Join is by unit code only. Names are display data and may differ between periods
  // (spacing, Unicode form), so they must never take part in the join.
  const compareMap = new Map(compareRows.map((row) => [String(row[PROVINCE_CODE_INDEX]), row]));
  const rows = currentRows.map((currentRow) => {
    const compareRow = compareMap.get(String(currentRow[PROVINCE_CODE_INDEX]));
    if (!compareRow) return null;

    // BC carries only the indicators the post-office report actually has.
    const metrics = tuyChonGR === "BC"
      ? [
        buildMetricComparison("F4.1 – PTC TMĐT", currentRow, compareRow, 5, 6),
        buildMetricComparison("F4.1 – PTC Truyền thống", currentRow, compareRow, 9, 10),
        buildMetricComparison("F4.1 – Chất lượng phát thành công tại bưu cục", currentRow, compareRow, 19, 20),
        buildMetricComparison("F1.3 – Chất lượng phát bưu gửi liên tỉnh", currentRow, compareRow, 33, 34)
      ]
      : [
        buildMetricComparison("F4.1 – PTC TMĐT", currentRow, compareRow, 5, 6),
        buildMetricComparison("F4.1 – PTC Truyền thống", currentRow, compareRow, 9, 10),
        buildMetricComparison("F3.3 – Thu gom", currentRow, compareRow, 13, 14),
        buildMetricComparison("📊 Trung bình", currentRow, compareRow, 15, 16),
        buildMetricComparison("F4.1 – Chất lượng phát thành công tại bưu cục", currentRow, compareRow, 19, 20),
        buildMetricComparison("F1.1 – Nội tỉnh", currentRow, compareRow, 24, 25),
        buildMetricComparison("F1.2 – Thu gom bưu gửi đi liên tỉnh", currentRow, compareRow, 29, 30),
        buildMetricComparison("F1.3 – Chất lượng phát bưu gửi liên tỉnh", currentRow, compareRow, 33, 34),
        buildMetricComparison("Tỷ lệ trung bình chung chất lượng (TGLT,PTC LT)", currentRow, compareRow, 35, 36)
      ];

    return {
      code: currentRow[PROVINCE_CODE_INDEX],
      name: currentRow[PROVINCE_NAME_INDEX],
      metrics
    };
  }).filter(Boolean).sort((a, b) => compareNumberOrText(a.code, b.code));

  rows.forEach((row) => (row.totalUnits = tuyChonGR === "TINH" ? 34 : rows.length));
  return rows;
}


function renderV2ComparisonDashboardWorksheet(currentRows, compareRows, tuyChonGR, from, to, compareFrom, compareTo) {
  const entityName = tuyChonGR === 'TINH' ? 'TOÀN QUỐC' : 'BƯU CỤC';
  const unitCode = currentRows.length > 0 && tuyChonGR === 'BC' ? currentRows[0][PROVINCE_CODE_INDEX].substring(0, 2) : '34';
  const dashName = tuyChonGR === 'TINH' ? 'Dashboard_Tinh' : `Dashboard_BC_${unitCode}`;
  
  let totalCurrentVol = 0;
  let totalCompareVol = 0;
  let totalCurrentSuccess = 0;
  let totalCompareSuccess = 0;
  
  let highestVolUnit = null;
  let highestVol = -1;
  let riskUnits = [];
  
  const compareMap = new Map(compareRows.map((row) => [String(row[PROVINCE_CODE_INDEX]), row]));

  const parseFormattedNumber = (val) => {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    const clean = String(val).replace(/,/g, '');
    const parsed = Number(clean);
    return isNaN(parsed) ? 0 : parsed;
  };

  const processedRows = currentRows.map((currentRow, idx) => {
    const compareRow = compareMap.get(String(currentRow[PROVINCE_CODE_INDEX]));
    if (!compareRow) return null;

    const code = currentRow[PROVINCE_CODE_INDEX];
    const name = currentRow[PROVINCE_NAME_INDEX];
    
    // index 17 = total volume for F4.1, 19 = rate
    const curVol = parseFormattedNumber(currentRow[17]);
    const compVol = parseFormattedNumber(compareRow[17]);
    const curRateStr = currentRow[19] || "0";
    const compRateStr = compareRow[19] || "0";
    
    // rates are often strings like "50.5%", need to parse them
    const curRate = Number(String(curRateStr).replace(/[^0-9.-]/g, '')) / 100 || 0;
    const compRate = Number(String(compRateStr).replace(/[^0-9.-]/g, '')) / 100 || 0;
    
    const curSucc = Math.round(curVol * curRate);
    const compSucc = Math.round(compVol * compRate);
    
    totalCurrentVol += curVol;
    totalCompareVol += compVol;
    totalCurrentSuccess += curSucc;
    totalCompareSuccess += compSucc;
    
    if (curVol > highestVol) {
      highestVol = curVol;
      highestVolUnit = name;
    }
    
    const diffRate = (curRate - compRate) * 100;
    let category = '';
    let categoryStyle = '';
    
    if (curVol > 100 && curRate < 0.8 && diffRate < 5) {
      category = '🔴 NGUY CƠ CAO (SL lớn, Tỷ lệ giảm/thấp)';
      categoryStyle = 's139';
      riskUnits.push({ name: name, vol: curVol, rate: curRate });
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
      code, name, curVol, compVol, curRate, compRate, diffRate, category, categoryStyle, chartBar, chartStyle
    };
  }).filter(Boolean);
  
  processedRows.forEach((r, idx) => r.stt = idx + 1);
  riskUnits.sort((a, b) => b.vol - a.vol);
  const topRisk = riskUnits.slice(0, 3);
  
  const avgCurRate = totalCurrentVol > 0 ? (totalCurrentSuccess / totalCurrentVol) : 0;
  const avgCompRate = totalCompareVol > 0 ? (totalCompareSuccess / totalCompareVol) : 0;
  
  let headerXml = `<Worksheet ss:Name="${escapeXml(dashName)}">
  <Table x:FullColumns="1" x:FullRows="1">
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
    <Cell ss:MergeAcross="2" ss:StyleID="m1730392434032"><Data ss:Type="String">${topRisk.length > 0 ? `🔴 ${riskUnits.length} Đơn vị nguy cơ cao&#10;(${escapeXml(topRisk.map(r => r.name).join(', '))})` : '🟢 Không có Đơn vị nguy cơ cao'}</Data></Cell>
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
    const diffStr = slDiff > 0 ? `+${slDiff.toLocaleString()} đơn` : `${slDiff.toLocaleString()} đơn`;
    const slDiffStyle = slDiff > 0 ? 's138' : (slDiff < 0 ? 's137' : 's135');
    const rateDiffStr = r.diffRate > 0 ? `tăng ${r.diffRate.toFixed(2)} điểm %, từ ${(r.compRate*100).toFixed(2)}% lên ${(r.curRate*100).toFixed(2)}%` : `giảm ${Math.abs(r.diffRate).toFixed(2)} điểm %, từ ${(r.compRate*100).toFixed(2)}% xuống ${(r.curRate*100).toFixed(2)}%`;
    return `   <Row>
    <Cell ss:StyleID="s135"><Data ss:Type="Number">${r.stt}</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">${escapeXml(r.code)}</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">${escapeXml(r.name)}</Data></Cell>
    <Cell ss:StyleID="s136"><Data ss:Type="Number">${totalCurrentVol > 0 ? (r.curVol / totalCurrentVol).toFixed(4) : 0}</Data></Cell>
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
    <Cell ss:MergeAcross="11" ss:StyleID="m1730392431952"><Data ss:Type="String">📊 BÁO CÁO PHÂN TÍCH QUẢN TRỊ NĂNG SUẤT &amp; CHẤT LƯỢNG (BI EXECUTIVE REPORT) - ${escapeXml(entityName)}:&#10;1. Tổng sản lượng phát toàn mạng: ${totalCurrentVol.toLocaleString()} bưu gửi (kỳ trước ${totalCompareVol.toLocaleString()} đơn, ${totalCurrentVol - totalCompareVol > 0 ? 'tăng' : 'giảm'} ${Math.abs(totalCurrentVol - totalCompareVol).toLocaleString()} đơn).&#10;2. Tỷ lệ phát thành công (F4.1) toàn mạng: ${(avgCurRate * 100).toFixed(2)}% (kỳ trước ${(avgCompRate * 100).toFixed(2)}%, ${avgCurRate - avgCompRate > 0 ? 'tăng' : 'giảm'} ${Math.abs((avgCurRate - avgCompRate) * 100).toFixed(2)} điểm %).&#10;3. Đơn vị gánh sản lượng lớn nhất: ${escapeXml(highestVolUnit || 'N/A')} với ${highestVol.toLocaleString()} đơn (chiếm ${((highestVol / (totalCurrentVol||1)) * 100).toFixed(1)}% tổng sản lượng).&#10;4. ${topRisk.length > 0 ? `🔴 CẢNH BÁO NGUY CƠ CAO (${riskUnits.length} đơn vị): ${escapeXml(topRisk.map(r => `${r.name} (SL ${r.vol.toLocaleString()} đơn, đạt ${(r.rate*100).toFixed(1)}%)`).join('; '))}. Cần kiểm tra khâu phát gấp!` : '🟢 KHÔNG CÓ CẢNH BÁO NGUY CƠ CAO.'}&#10;5. CHỈ ĐẠO ĐIỀU HÀNH: Yêu cầu Trưởng các Đơn vị thuộc nhóm Cảnh báo nguy cơ tập trung rà soát lực lượng bưu tá và kiểm soát quét TMS hoàn thành đúng chỉ tiêu.</Data></Cell>
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
  const isNumber = type === "Number" && value !== "" && Number.isFinite(Number(value));
  const safeType = isNumber ? "Number" : "String";
  const safeValue = isNumber ? String(Number(value)) : escapeXml(value);
  const safeStyleId = !isNumber && (styleId === "Percent" || styleId === "Number") ? "Text" : styleId;
  return `<Cell ss:StyleID="${safeStyleId}"><Data ss:Type="${safeType}">${safeValue}</Data></Cell>`;
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
    .filter((row) => hasV2SourceValue(row[valueIndex]) && Number.isFinite(Number(row[valueIndex])) && Number(row[valueIndex]) >= 0)
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
  const configuredNameIndex = tuyChonGR === "BC" && Number.isInteger(config.bcNameIndex) ? config.bcNameIndex : config.nameIndex;
  const nameIndex = Number.isInteger(configuredNameIndex) ? configuredNameIndex : codeIndex + 1;
  if (cells[nameIndex] && !Number.isInteger(toIntegerCode(cells[nameIndex]))) return cells[nameIndex];
  return cells.find((cell, index) => index > codeIndex && /[A-Za-zÀ-ỹ]/.test(cell) && !isPercentLike(cell)) || "";
}

function getReportCodeIndex(cells, config, tuyChonGR) {
  // BC rows are keyed by the post-office code, which sits in its own column. Without this
  // branch every BC row falls back to config.codeIndex (the province column) and the whole
  // report collapses into a single province row.
  if (tuyChonGR === "BC" && Number.isInteger(config.bcCodeIndex)) return config.bcCodeIndex;
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
  return select ? cleanText(select.value) : "";
}

function getSelectedProvinceName() {
  const select = document.getElementById("province-select");
  if (!select || !select.selectedOptions || !select.selectedOptions.length) return "";
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
      selectedBcProvName,
      { allowMissingSelectedProvince: true }
    );
    results.push({
      monthObj,
      rows
    });
  }

  return results;
}

function buildVisualTrendBar(diffPts) {
  if (diffPts == null || isNaN(diffPts)) return "⚪ N/A";
  const sign = diffPts >= 0 ? "+" : "";
  const valStr = `${sign}${diffPts.toFixed(2)}%`;
  if (diffPts >= 3.0) return `🟢▲▲ ${valStr}`;
  if (diffPts > 0) return `🟢▲ ${valStr}`;
  if (diffPts <= -3.0) return `🔴▼▼ ${valStr}`;
  if (diffPts < 0) return `🔴▼ ${valStr}`;
  return `⚪ ${valStr}`;
}

const MULTI_MONTH_METRICS_MAP = {
  "F1.1 – Nội tỉnh": { shortName: "F1.1", target: 0.95, totalCol: 22, onTimeCol: 23 },
  "F1.2 – Thu gom bưu gửi đi liên tỉnh": { shortName: "F1.2", target: 0.95, totalCol: 26, onTimeCol: 28 },
  "F1.3 – Chất lượng phát bưu gửi liên tỉnh": { shortName: "F1.3", target: 0.90, totalCol: 31, onTimeCol: 32 },
  "F4.1 – Chất lượng phát thành công tại bưu cục": { shortName: "F4.1", target: 0.90, totalCol: 17, onTimeCol: 18 }
};

// SSOT for which KPIs a multi-month export renders. The legacy .xls dashboard and the
// native .xlsx exporter must both honour the PO's tick boxes: an unticked KPI is never
// fetched, so rendering it would only produce fake 0 volumes and 0% rates.
function getActiveMultiMonthMetrics(tuyChonGR) {
  const activeMetrics = getSelectedMetricLabels()
    .map(label => MULTI_MONTH_METRICS_MAP[label])
    .filter(Boolean);

  if (activeMetrics.length === 0) {
    activeMetrics.push(MULTI_MONTH_METRICS_MAP["F4.1 – Chất lượng phát thành công tại bưu cục"]);
    activeMetrics.push(MULTI_MONTH_METRICS_MAP["F1.3 – Chất lượng phát bưu gửi liên tỉnh"]);
  }

  // BC never fetches F1.1/F1.2 (see V2_BC_SOURCE_KEYS), so rendering a card, a series or a
  // matrix column for them would be a fabricated 0.
  if (tuyChonGR === "BC") {
    const bcMetrics = activeMetrics.filter((metric) => V2_BC_SHORT_NAMES.includes(metric.shortName));
    assertBcKpiSelection(bcMetrics.length);
    return bcMetrics;
  }
  return activeMetrics;
}

function getKpiStatus(rate, target) {
  if (rate == null || !Number.isFinite(Number(rate))) {
    return { key: "gray", label: "⚪ N/A", styleId: "KpiGray", gapPoints: null };
  }

  const gapPoints = (Number(rate) - Number(target)) * 100;
  if (gapPoints >= 0) {
    return { key: "green", label: "🟢 ĐẠT MỤC TIÊU", styleId: "KpiGreen", gapPoints };
  }
  if (gapPoints >= -5 - 1e-9) {
    return { key: "yellow", label: "🟡 THẤP HƠN ≤ 5 ĐIỂM %", styleId: "KpiYellow", gapPoints };
  }
  return { key: "red", label: "🔴 THẤP HƠN > 5 ĐIỂM %", styleId: "KpiRed", gapPoints };
}

function formatKpiRate(rate) {
  return rate == null || !Number.isFinite(Number(rate)) ? "N/A" : `${(Number(rate) * 100).toFixed(2)}%`;
}

function formatKpiGap(gapPoints) {
  if (gapPoints == null || !Number.isFinite(Number(gapPoints))) return "N/A";
  const sign = gapPoints >= 0 ? "+" : "";
  return `${sign}${Number(gapPoints).toFixed(2)} điểm %`;
}

function getMonthlyRateTrendDirection(delta) {
  if (!Number.isFinite(delta)) return null;
  if (Math.abs(delta) < 1e-12) return 0;
  return delta > 0 ? 1 : -1;
}

// BC has no province row of its own, so the province headline is the sum of its post
// offices. Only the metric columns are synthesised — that is all the summary reads.
function aggregateUnitRows(units, code, name, totalMonths, activeMetrics) {
  const monthRows = Array.from({ length: totalMonths }, (_, monthIdx) => {
    const row = [];
    row[PROVINCE_CODE_INDEX] = String(code);
    row[PROVINCE_NAME_INDEX] = name;
    let monthHasData = false;
    activeMetrics.forEach((metric) => {
      let total = 0;
      let onTime = 0;
      let hasData = false;
      units.forEach((unit) => {
        const unitRow = unit.monthRows[monthIdx];
        if (!unitRow || !hasV2SourceValue(unitRow[metric.totalCol]) || !hasV2SourceValue(unitRow[metric.onTimeCol])) return;
        hasData = true;
        total += toNumberValue(unitRow[metric.totalCol]);
        onTime += toNumberValue(unitRow[metric.onTimeCol]);
      });
      if (!hasData) return;
      monthHasData = true;
      row[metric.totalCol] = String(total);
      row[metric.onTimeCol] = String(onTime);
    });
    return monthHasData ? row : null;
  });
  return { code: String(code), name, monthRows };
}

function renderV2MultiMonthDashboardWorksheet(monthlyResults, monthsList, tuyChonGR) {
  const provCode = getSelectedProvinceCode();
  const provName = getSelectedProvinceName();
  const fromLabel = monthsList[0]?.label || "";
  const toLabel = monthsList[monthsList.length - 1]?.label || "";
  const totalMonths = monthsList.length;

  const activeMetrics = getActiveMultiMonthMetrics(tuyChonGR);

  // BC lists post offices of one province; TINH lists the 34 provinces. Every wording that
  // names the peer set has to follow, so a BC rank is never read as a national one.
  const isBcMatrix = tuyChonGR === "BC";
  const peerScopeLabel = isBcMatrix ? `Bưu cục thuộc ${provName}` : "BĐT/TP toàn quốc";
  const rankScopeLabel = isBcMatrix
    ? "xếp hạng trong các Bưu cục có dữ liệu hợp lệ của tỉnh được chọn"
    : "xếp hạng toàn quốc";

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

  const summariseUnit = (u) => {
    const isTargetUnit = String(u.code) === String(provCode);

    const metricStats = activeMetrics.map((m) => {
      const points = u.monthRows.map((row) => {
        const hasData = Boolean(row)
          && hasV2SourceValue(row[m.totalCol])
          && hasV2SourceValue(row[m.onTimeCol]);
        if (!hasData) return { volume: null, onTime: null, rate: null, rank: null, rankCount: 0, hasData: false };
        const volume = toNumberValue(row[m.totalCol]);
        const onTime = toNumberValue(row[m.onTimeCol]);
        return { volume, onTime, rate: volume > 0 ? onTime / volume : 0, rank: null, rankCount: 0, hasData: true };
      });
      const sumVol = points.reduce((sum, point) => sum + (point.hasData ? point.volume : 0), 0);
      const sumOnTime = points.reduce((sum, point) => sum + (point.hasData ? point.onTime : 0), 0);
      const hasData = points.some(point => point.hasData);
      const cumRate = hasData ? (sumVol > 0 ? sumOnTime / sumVol : 0) : null;
      const lastPoint = points[totalMonths - 1];
      const previousPoint = points[totalMonths - 2];
      const diffRate = lastPoint?.hasData && previousPoint?.hasData ? (lastPoint.rate - previousPoint.rate) * 100 : null;
      const status = getKpiStatus(cumRate, m.target);

      return { points, rates: points.map(point => point.rate), sumVol, sumOnTime, cumRate, diffRate, rank: 0, hasData, status };
    });

    const statusKeys = metricStats.map(s => s.status.key);
    let segment = "⚪ N/A";
    let segStyle = "KpiGray";
    if (statusKeys.includes("red")) {
      segment = "🔴 CẢNH BÁO: CÓ KPI THẤP HƠN MỤC TIÊU > 5 ĐIỂM %";
      segStyle = "KpiRed";
    } else if (statusKeys.includes("yellow")) {
      segment = "🟡 CẦN CẢI THIỆN: CÓ KPI THẤP HƠN MỤC TIÊU ≤ 5 ĐIỂM %";
      segStyle = "KpiYellow";
    } else if (statusKeys.includes("green")) {
      segment = "🟢 ĐẠT MỤC TIÊU KPI";
      segStyle = "KpiGreen";
    }

    return {
      code: u.code,
      name: u.name,
      isTargetUnit,
      metricStats,
      segment,
      segStyle
    };
  };

  const unitSummaries = units.map(summariseUnit);

  // Ranking sample, per KPI and per month, over the units that have valid data. For TINH
  // that is the 34 provinces; for BC it is the post offices of the SELECTED province —
  // `unitSummaries` holds exactly those, so no post office of another province can enter
  // the sample and the 34-province scale is never reused. N/A months are filtered out and
  // stay N/A. Ties follow the project's existing rule (assignRankByColumn): ordinal
  // ranking with equal rates ordered by ascending unit code.
  activeMetrics.forEach((m, mIdx) => {
    monthsList.forEach((_, monthIdx) => {
      const sorted = [...unitSummaries]
        .filter(u => u.metricStats[mIdx].points[monthIdx].hasData)
        .sort((a, b) => b.metricStats[mIdx].points[monthIdx].rate - a.metricStats[mIdx].points[monthIdx].rate
          || compareNumberOrText(a.code, b.code));
      sorted.forEach((u, rankIdx) => {
        const point = u.metricStats[mIdx].points[monthIdx];
        point.rank = rankIdx + 1;
        point.rankCount = sorted.length;
      });
    });
  });

  // TINH focuses on one province among its units. BC has no province row — its units are
  // that province's post offices — so the focus is their aggregate.
  const focusSummary = tuyChonGR === "BC"
    ? summariseUnit(aggregateUnitRows(units, provCode, provName, totalMonths, activeMetrics))
    : unitSummaries.find(u => u.isTargetUnit);
  if (!focusSummary) {
    throw new Error(`Không tìm thấy dữ liệu dashboard cho tỉnh tiêu điểm mã ${provCode} (${provName}).`);
  }

  // -------------------------------------------------------------
  // TAB 1: EXECUTIVE DASHBOARD FOR FOCUS TARGET UNIT
  // Clean 13-column grid (fits 100% on laptop/desktop screen)
  // -------------------------------------------------------------
  const dashColWidths = [40, 220];
  monthsList.forEach(() => dashColWidths.push(90, 80, 72, 82, 70));
  const dashTotalCols = dashColWidths.length;

  const focusTrendRowsXml = activeMetrics.map((m, mIdx) => {
    const ms = focusSummary.metricStats[mIdx];
    const fullName = m.shortName === 'F4.1' ? 'Chất lượng phát thành công tại bưu cục'
                   : m.shortName === 'F1.3' ? 'Chất lượng phát bưu gửi liên tỉnh'
                   : m.shortName === 'F1.1' ? 'Chất lượng phát bưu gửi nội tỉnh'
                   : 'Chất lượng thu gom bưu gửi đi liên tỉnh';
    const monthCells = ms.points.map((point, monthIdx) => {
      const previous = monthIdx > 0 ? ms.points[monthIdx - 1] : null;
      const delta = previous?.hasData && point.hasData ? point.rate - previous.rate : null;
      const direction = getMonthlyRateTrendDirection(delta);
      const deltaCell = monthIdx === 0
        ? excelXmlCell("—", "String", "Same")
        : delta == null
          ? excelXmlCell("N/A", "String", "KpiGray")
          : excelXmlCell(delta, "Number", direction > 0 ? "Up" : direction < 0 ? "Down" : "Same");
      const trendCell = monthIdx === 0
        ? excelXmlCell("—", "String", "Same")
        : direction == null
          ? excelXmlCell("N/A", "String", "KpiGray")
          : excelXmlCell(direction > 0 ? "▲" : direction < 0 ? "▼" : "→", "String", direction > 0 ? "Up" : direction < 0 ? "Down" : "Same");
      if (!point.hasData) {
        return [
          excelXmlCell("N/A", "String", "KpiGray"),
          excelXmlCell("N/A", "String", "KpiGray"),
          excelXmlCell("N/A", "String", "KpiGray"),
          deltaCell,
          trendCell
        ].join("");
      }
      return [
        excelXmlCell(point.volume, "Number", "Number"),
        excelXmlCell(point.rate, "Number", "Number"),
        excelXmlCell(`${point.rank} / ${point.rankCount}`, "String", "Text"),
        deltaCell,
        trendCell
      ].join("");
    }).join("");
    return `
      <Row ss:Height="22">
        ${excelXmlCell(mIdx + 1, "Number", "Number")}
        ${excelXmlCell(`${m.shortName} – ${fullName}`, "String", "Text")}
        ${monthCells}
      </Row>`;
  }).join("");

  const comboChartRowsXml = activeMetrics.map((m, mIdx) => {
    const ms = focusSummary.metricStats[mIdx];
    const actualCells = ms.rates.map(rate => excelXmlCell(rate, "Number", getKpiStatus(rate, m.target).styleId)).join("");
    const targetCells = monthsList.map(() => excelXmlCell(m.target, "Number", "KpiTargetLine")).join("");
    return `
      <Row ss:Height="22">
        ${excelXmlCell(mIdx + 1, "Number", "Number")}
        ${excelXmlCell(`${m.shortName} – Cột tỷ lệ thực tế`, "String", "Text")}
        ${actualCells}
      </Row>
      <Row ss:Height="20">
        ${excelXmlCell("", "String", "Text")}
        ${excelXmlCell(`${m.shortName} – Đường mục tiêu riêng`, "String", "KpiTargetLine")}
        ${targetCells}
      </Row>`;
  }).join("");

  const notesTextLines = [
    `🏆 BÁO CÁO LŨY KẾ & MA TRẬN SO SÁNH CHUỖI XU HƯỚNG ĐA THÁNG (EXECUTIVE BI DASHBOARD):`,
    `1. ĐƠN VỊ TRỌNG ĐIỂM QUAN SÁT: ${provName} (Mã ${provCode}) | KHOẢNG THỜI GIAN: ${fromLabel} ➔ ${toLabel} (${totalMonths} tháng liên tục).`
  ];
  activeMetrics.forEach((m, mIdx) => {
    const ms = focusSummary ? focusSummary.metricStats[mIdx] : { sumVol: 0, cumRate: null, rank: 0, status: getKpiStatus(null, m.target) };
    notesTextLines.push(`${mIdx + 2}. TIÊU ĐIỂM ${m.shortName} (${provName}): Mục tiêu SSOT ${(m.target * 100).toFixed(0)}%, thực tế ${formatKpiRate(ms.cumRate)}, chênh ${formatKpiGap(ms.status.gapPoints)}, trạng thái ${ms.status.label}; sản lượng lũy kế ${ms.sumVol.toLocaleString()} bưu gửi${ms.rank > 0 ? `, xếp hạng ${ms.rank}/${units.length} ${peerScopeLabel}` : ""}.`);
  });
  const redKpis = activeMetrics.filter((m, idx) => focusSummary.metricStats[idx].status.key === "red").map(m => m.shortName);
  const yellowKpis = activeMetrics.filter((m, idx) => focusSummary.metricStats[idx].status.key === "yellow").map(m => m.shortName);
  const naKpis = activeMetrics.filter((m, idx) => focusSummary.metricStats[idx].status.key === "gray").map(m => m.shortName);
  const executiveActions = [];
  if (redKpis.length) executiveActions.push(`ưu tiên xử lý KPI đỏ ${redKpis.join(", ")} đang thấp hơn mục tiêu riêng trên 5 điểm %`);
  if (yellowKpis.length) executiveActions.push(`lập kế hoạch cải thiện KPI vàng ${yellowKpis.join(", ")} đang thấp hơn mục tiêu riêng không quá 5 điểm %`);
  if (naKpis.length) executiveActions.push(`rà soát dữ liệu N/A của ${naKpis.join(", ")}`);
  if (!executiveActions.length) executiveActions.push(`duy trì ${activeMetrics.length} KPI đã chọn đạt ngưỡng mục tiêu riêng`);
  // The narrative must name only the KPIs actually exported: quoting a threshold for an
  // unticked KPI implies data the workbook does not contain.
  const thresholdText = activeMetrics.map((m) => `${m.shortName} = ${(m.target * 100).toFixed(0)}%`).join(", ");
  notesTextLines.push(`${activeMetrics.length + 2}. 🎯 CHỈ ĐẠO ĐIỀU HÀNH BAN GIÁM ĐỐC: Ngưỡng SSOT ${thresholdText}; ${executiveActions.join("; ")}.`);
  const notesText = notesTextLines.join("\n");

  const executiveDashboardSheet = `<Worksheet ss:Name="Dashboard_${escapeXml(provName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ''))}">
    <Table>
      ${renderColumnWidths(dashColWidths)}
      <Row ss:Height="26"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Title"><Data ss:Type="String">🏆 BAN GIÁM ĐỐC BI DASHBOARD - TỔNG QUAN CHẤT LƯỢNG KÉP &amp; CHUỖI XU HƯỚNG BƯU ĐIỆN TỈNH ${escapeXml(provName.toUpperCase())} (MÃ ${provCode})</Data></Cell></Row>
      <Row ss:Height="18"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Text"><Data ss:Type="String">Đơn vị tiêu điểm: ${escapeXml(provName)} (Mã ${provCode})  |  Theo dõi chuỗi xu hướng ${totalMonths} tháng liên tục (${escapeXml(fromLabel)} ➔ ${escapeXml(toLabel)})  |  Đối sánh ${units.length} ${peerScopeLabel}</Data></Cell></Row>
      <Row></Row>

      <!-- KHỐI 1: BẢNG DỮ LIỆU KPI RIÊNG TỪNG THÁNG -->
      <Row ss:Height="22"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Group"><Data ss:Type="String">📈 BẢNG DỮ LIỆU KPI TỪNG THÁNG RIÊNG CHO ${escapeXml(provName.toUpperCase())} — mỗi tháng lấy trực tiếp từ monthlyResults và ${rankScopeLabel} của tháng đó</Data></Cell></Row>
      <Row ss:Height="24">
        ${excelXmlCell("STT", "String", "Header")}
        ${excelXmlCell("Chỉ Tiêu Chất Lượng", "String", "Header")}
        ${monthsList.map(m => ["Sản lượng", "Tỷ lệ", "Hạng", "Chênh TL", "Xu hướng"].map(label => excelXmlCell(`${m.monthKey} ${label}`, "String", "Header")).join("")).join("")}
      </Row>
      ${focusTrendRowsXml}
      <Row></Row>

      <!-- KHỐI 2: COMBO CHART - CỘT THỰC TẾ & ĐƯỜNG MỤC TIÊU RIÊNG -->
      <Row ss:Height="22"><Cell ss:MergeAcross="${dashTotalCols - 1}" ss:StyleID="Group"><Data ss:Type="String">📊 COMBO CHART KPI - CỘT TỶ LỆ THỰC TẾ &amp; ĐƯỜNG MỤC TIÊU RIÊNG THEO SSOT</Data></Cell></Row>
      <Row ss:Height="24">
        ${excelXmlCell("STT", "String", "Header")}
        ${excelXmlCell("KPI / Series", "String", "Header")}
        ${monthsList.map(m => excelXmlCell(m.monthKey, "String", "Header")).join("")}
      </Row>
      ${comboChartRowsXml}
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
  // BC lists each post office with volume, rate, month-over-month delta and direction per
  // month, so a reader can follow one post office across the period. TINH keeps its
  // rate-only layout untouched.

  const headerCols = [
    "STT", "Mã", "Tên Đơn Vị"
  ];
  activeMetrics.forEach(m => {
    monthsList.forEach(month => {
      if (isBcMatrix) {
        headerCols.push(
          `${m.shortName} ${month.monthKey} – Sản lượng`,
          `${m.shortName} ${month.monthKey} – Tỷ lệ`,
          `${m.shortName} ${month.monthKey} – Hạng`,
          `${m.shortName} ${month.monthKey} – Chênh lệch`,
          `${m.shortName} ${month.monthKey} – Xu hướng`
        );
      } else {
        headerCols.push(`${m.shortName} ${month.monthKey}`);
      }
    });
    headerCols.push(`Lũy Kế ${m.shortName} BQ`, `MoM ${m.shortName} (Tn vs T1)`);
  });
  headerCols.push("Phân loại Vị thế BI");

  const colWidths = [40, 60, 190];
  activeMetrics.forEach(() => {
    monthsList.forEach(() => {
      if (isBcMatrix) colWidths.push(95, 80, 70, 90, 80);
      else colWidths.push(85);
    });
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

    u.metricStats.forEach((ms, mIdx) => {
      const metric = activeMetrics[mIdx];
      if (isBcMatrix) {
        rowXml += ms.points.map((point, monthIdx) => {
          if (!point.hasData) {
            return excelXmlCell("N/A", "String", "KpiGray").repeat(5);
          }
          const previous = monthIdx > 0 ? ms.points[monthIdx - 1] : null;
          const delta = previous?.hasData ? (point.rate - previous.rate) * 100 : null;
          const direction = getMonthlyRateTrendDirection(delta);
          // The first month has no predecessor, so its delta/trend are "—" rather than
          // N/A — the month itself still carries a real rank.
          const deltaCell = monthIdx === 0
            ? excelXmlCell("—", "String", "Same")
            : delta == null
              ? excelXmlCell("N/A", "String", "KpiGray")
              : excelXmlCell(delta, "Number", trendStyle(delta));
          const trendCell = monthIdx === 0
            ? excelXmlCell("—", "String", "Same")
            : direction == null
              ? excelXmlCell("N/A", "String", "KpiGray")
              : excelXmlCell(direction > 0 ? "▲" : direction < 0 ? "▼" : "→", "String", direction > 0 ? "Up" : direction < 0 ? "Down" : "Same");
          return [
            excelXmlCell(point.volume, "Number", "Number"),
            excelXmlCell(point.rate, "Number", getKpiStatus(point.rate, metric.target).styleId),
            excelXmlCell(point.rank ? `${point.rank}/${point.rankCount}` : "N/A", "String", point.rank ? "Text" : "KpiGray"),
            deltaCell,
            trendCell
          ].join("");
        }).join("");
      } else {
        rowXml += ms.rates.map(r => excelXmlCell(r, "Number", getKpiStatus(r, metric.target).styleId)).join("");
      }
      rowXml += excelXmlCell(ms.cumRate, "Number", ms.status.styleId);
      rowXml += excelXmlCell(buildVisualTrendBar(ms.diffRate), "String", trendStyle(ms.diffRate));
    });

    rowXml += `
        ${excelXmlCell(u.segment, "String", u.segStyle)}
      </Row>`;
    return rowXml;
  }).join("");

  const matrixSheetName = isBcMatrix ? "MaTran_BuuCuc" : "MaTran_34_Tinh_ToanQuoc";
  const matrixTitle = isBcMatrix
    ? `📊 BẢNG MA TRẬN BƯU CỤC TRỰC THUỘC ${escapeXml(provName.toUpperCase())} (MÃ ${escapeXml(String(provCode))}) – ${units.length} BƯU CỤC (${metricShortNames}) (${totalMonths} THÁNG)`
    : `📊 BẢNG MA TRẬN ĐỐI SÁNH &amp; XẾP HẠNG 34 BƯU ĐIỆN TỈNH/TP TOÀN QUỐC (${metricShortNames}) (${totalMonths} THÁNG - HIGHLIGHT DÒNG ${escapeXml(provName.toUpperCase())} ★)`;

  // The ranking sample must be stated on the sheet so no reader mistakes a BC rank for a
  // national one.
  const rankingSampleNote = isBcMatrix
    ? `Hạng trong các Bưu cục có dữ liệu hợp lệ thuộc tỉnh được chọn (mã ${escapeXml(String(provCode))}) trong từng tháng, tính riêng theo từng KPI. Đồng tỷ lệ: xếp theo mã Bưu cục tăng dần. Tháng không có dữ liệu: N/A, không tham gia xếp hạng.`
    : "Hạng trong 34 Bưu điện Tỉnh/TP toàn quốc trong từng tháng, tính riêng theo từng KPI.";

  const benchmarkMatrixSheet = `<Worksheet ss:Name="${matrixSheetName}">
    <Table>
      ${renderColumnWidths(colWidths)}
      <Row ss:Height="24"><Cell ss:MergeAcross="${totalCols - 1}" ss:StyleID="Group"><Data ss:Type="String">${matrixTitle}</Data></Cell></Row>
      <Row ss:Height="20"><Cell ss:MergeAcross="${totalCols - 1}" ss:StyleID="Text"><Data ss:Type="String">${escapeXml(rankingSampleNote)}</Data></Cell></Row>
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

    let fileName;
    let blob;
    if (isNativeV2XlsxExportEnabled()) {
      setStatus("Đang tạo Dashboard .xlsx native với Combo Chart OOXML...");
      if (!globalThis.DkclXlsxExport || !globalThis.DkclXlsxExport.isAvailable()) {
        throw new Error(`Không nạp được fflate offline để tạo .xlsx. Có thể rollback tạm thời bằng localStorage.${NATIVE_XLSX_FEATURE_FLAG_KEY} = "false".`);
      }
      const packageResult = globalThis.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
        monthlyResults,
        monthsList,
        tuyChonGR,
        selectedProvinceCode: provCode,
        selectedProvinceName: provName,
        metrics: getActiveMultiMonthMetrics(tuyChonGR),
        columns: V2_EXCEL_COLUMNS,
        classifyKpi: getKpiStatus,
        // SSOT: the native .xlsx must read row cells with the same numeric semantics the
        // legacy .xls dashboard uses, otherwise the two exports disagree on total/rate.
        toNumberValue
      });
      if (packageResult.bytes[0] !== 0x50 || packageResult.bytes[1] !== 0x4B) {
        throw new Error("Package .xlsx không có magic bytes PK.");
      }
      fileName = `filebaocaov2_luyke_thang_${tuyChonGR === "TINH" ? "tinh" : "bc"}_${fromMonthInput.value}_vs_${toMonthInput.value}.xlsx`;
      blob = new Blob([packageResult.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    } else {
      setStatus("Feature flag rollback đang bật: xuất SpreadsheetML .xls tương thích...");
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
      fileName = `filebaocaov2_luyke_thang_${tuyChonGR === "TINH" ? "tinh" : "bc"}_${fromMonthInput.value}_vs_${toMonthInput.value}.xls`;
      blob = new Blob([xmlContent], { type: "application/vnd.ms-excel;charset=utf-8" });
    }
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
