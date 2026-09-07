const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const popupSource = fs.readFileSync(path.join(__dirname, "..", "popup.js"), "utf8");
const fflateSource = fs.readFileSync(path.join(__dirname, "..", "vendor", "fflate.min.js"), "utf8");
const xlsxExporterSource = fs.readFileSync(path.join(__dirname, "..", "xlsx-export.js"), "utf8");

function decodeHtml(value) {
  return value
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'");
}

class TestDomParser {
  parseFromString(html) {
    const rows = [...html.matchAll(/<tr([^>]*)>([\s\S]*?)<\/tr>/gi)].map((match) => {
      const cells = [...match[2].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)].map((cell) => ({
        textContent: decodeHtml(cell[1].replace(/<[^>]+>/g, "").trim())
      }));
      return {
        children: cells,
        classList: { contains: (name) => new RegExp(`\\b${name}\\b`).test(match[1]) }
      };
    });
    return { querySelectorAll: (selector) => (selector === "tr" ? rows : []) };
  }
}

function createElement(overrides = {}) {
  return {
    value: "",
    checked: false,
    hidden: false,
    dataset: {},
    classList: { add() {}, remove() {}, toggle() {} },
    selectedOptions: [],
    addEventListener() {},
    setAttribute() {},
    getAttribute() { return "false"; },
    querySelector() { return null; },
    ...overrides
  };
}

function createSandbox(fetchImpl, selectedProvince = { code: "", name: "" }, storageValues = {}) {
  const elements = new Map();
  const storage = new Map(Object.entries(storageValues));
  const document = {
    getElementById(id) {
      if (!elements.has(id)) {
        const overrides = id === "province-select" ? {
          value: String(selectedProvince.code),
          selectedOptions: [{ textContent: `${selectedProvince.code} - ${selectedProvince.name}` }]
        } : {};
        elements.set(id, createElement(overrides));
      }
      return elements.get(id);
    },
    querySelectorAll() { return []; }
  };
  const sandbox = {
    document,
    window: {},
    localStorage: {
      getItem(key) { return storage.has(key) ? storage.get(key) : null; },
      setItem(key, value) { storage.set(key, String(value)); }
    },
    DOMParser: TestDomParser,
    console: { log() {}, error() {} },
    fetch: fetchImpl,
    URLSearchParams,
    URL,
    Map,
    Set,
    Date,
    Math,
    Number,
    String,
    Array,
    Object,
    RegExp,
    JSON,
    isNaN
  };
  sandbox.window.parent = sandbox.window;
  vm.createContext(sandbox);
  vm.runInContext(fflateSource, sandbox);
  vm.runInContext(xlsxExporterSource, sandbox);
  vm.runInContext(popupSource, sandbox);
  return sandbox;
}

function makeSourceHtml(codes, provinceNames = {}, { zeroValues = false } = {}) {
  return codes.map((code) => {
    const cells = Array(28).fill("");
    cells[1] = String(code);
    cells[2] = provinceNames[code] || "Ngoài danh mục";
    // One synthetic aggregate row is served for every endpoint. Its values satisfy the
    // golden offsets from the old extension: F1.2 (7,8,9), F4.1 (10,27),
    // F1.1 (11,18,19), and F1.3 (10,18).
    const value = (amount) => (zeroValues ? "0" : String(amount));
    cells[7] = value(200);
    cells[8] = value(190);
    cells[9] = value(180);
    cells[10] = value(200);
    cells[11] = value(170);
    cells[18] = value(160);
    cells[19] = value(150);
    cells[27] = value(115);
    return `<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`;
  }).join("");
}

function assertWellFormedXml(xml) {
  const stack = [];
  const tags = xml.match(/<[^>]+>/g) || [];
  for (const tag of tags) {
    if (/^<\?|^<!--/.test(tag)) continue;
    if (/^<\//.test(tag)) {
      const name = tag.slice(2, -1).trim();
      assert.equal(stack.pop(), name, `XML closing tag mismatch for ${name}`);
    } else if (!/\/>$/.test(tag)) {
      const name = tag.slice(1, -1).trim().split(/\s+/)[0];
      stack.push(name);
    }
  }
  assert.equal(stack.length, 0, "XML has unclosed elements");
}

function assertNoDuplicateXmlAttributes(xml, partName = "XML") {
  for (const match of xml.matchAll(/<[^!?/][^>]*>/g)) {
    const attributes = new Set();
    const tag = match[0];
    for (const attribute of tag.matchAll(/\s([A-Za-z_][\w:.-]*)\s*=/g)) {
      assert.ok(!attributes.has(attribute[1]), `${partName} duplicates attribute ${attribute[1]} in ${tag}`);
      attributes.add(attribute[1]);
    }
  }
}

function assertDashboardTableIsValid(xml) {
  assert.doesNotMatch(xml, /Expanded(?:Row|Column)Count/, "dashboard must not declare stale Expanded*Count values");
  const table = xml.match(/<Table[^>]*>([\s\S]*?)<\/Table>/)?.[1];
  assert.ok(table, "dashboard contains a Table");
  const declaredColumns = [...table.matchAll(/<Column\b([^>]*)\/?>(?:<\/Column>)?/g)].reduce((count, match) => {
    const span = Number(match[1].match(/ss:Span="(\d+)"/)?.[1] || 0);
    return count + span + 1;
  }, 0);
  assert.equal(declaredColumns, 12, "dashboard has exactly 12 declared columns");
  for (const row of table.matchAll(/<Row\b[^>]*>([\s\S]*?)<\/Row>/g)) {
    const width = [...row[1].matchAll(/<Cell\b([^>]*)>/g)].reduce((count, match) => {
      const mergeAcross = Number(match[1].match(/ss:MergeAcross="(\d+)"/)?.[1] || 0);
      return count + mergeAcross + 1;
    }, 0);
    assert.ok(width <= 12, `row uses ${width} columns, exceeds dashboard width`);
  }
}

function assertAllStyleIdsDefined(workbook) {
  const declaredStyleIds = new Set([...workbook.matchAll(/<Style ss:ID="([^"]+)"/g)].map((match) => match[1]));
  for (const match of workbook.matchAll(/<Cell\b[^>]*ss:StyleID="([^"]+)"/g)) {
    assert.ok(declaredStyleIds.has(match[1]), `undefined SpreadsheetML style: ${match[1]}`);
  }
}

function assertWorksheetTablesFitDeclaredColumns(xml) {
  assert.doesNotMatch(xml, /ss:MergeAcross="-\d+"/, "SpreadsheetML must not contain a negative merge span");
  const tables = [...xml.matchAll(/<Table[^>]*>([\s\S]*?)<\/Table>/g)];
  assert.ok(tables.length > 0, "workbook contains at least one table");
  tables.forEach((tableMatch) => {
    const table = tableMatch[1];
    const declaredColumns = [...table.matchAll(/<Column\b([^>]*)\/?>(?:<\/Column>)?/g)].reduce((count, match) => {
      const span = Number(match[1].match(/ss:Span="(\d+)"/)?.[1] || 0);
      return count + span + 1;
    }, 0);
    assert.ok(declaredColumns > 0, "table declares its columns");
    for (const row of table.matchAll(/<Row\b[^>]*>([\s\S]*?)<\/Row>/g)) {
      const width = [...row[1].matchAll(/<Cell\b([^>]*)>/g)].reduce((count, match) => {
        const mergeAcross = Number(match[1].match(/ss:MergeAcross="(\d+)"/)?.[1] || 0);
        return count + mergeAcross + 1;
      }, 0);
      assert.ok(width <= declaredColumns, `row uses ${width} columns, exceeds declared width ${declaredColumns}`);
    }
  });
}

function makeComparisonRows(prefix) {
  return [
    { code: 53, name: "Thừa Thiên Huế" },
    { code: 10, name: "TP. Hà Nội" }
  ].map(({ code, name }, index) => {
    const row = Array(37).fill("");
    row[0] = String(index + 1);
    row[1] = String(code);
    row[2] = `${name} ${prefix}`;
    row[17] = String(100 + index * 20);
    row[18] = String(90 + index * 20);
    row[19] = "0.9";
    return row;
  });
}

const V2_SOURCE_CASES = [
  { path: "/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc", name: "Chất lượng phát thành công tại bưu cục (F4.1)" },
  { path: "/kpi/chat-luong-toan-trinh-buu-gui-noi-tinh", name: "Nội tỉnh F1.1" },
  { path: "/kpi/chat-luong-thu-gom-buu-lien-tinh", name: "Thu gom bưu gửi đi liên tỉnh (F1.2)" },
  { path: "/kpi/chat-luong-phat-buu-gui-lien-tinh", name: "Chất lượng phát bưu gửi liên tỉnh (F1.3)" }
];

async function testProvinceJoin(selectedProvince) {
  const requestedSources = new Set();
  const sandbox = createSandbox(async (requestUrl) => {
    requestedSources.add(new URL(requestUrl).pathname);
    return {
      ok: true,
      text: async () => JSON.stringify({
        data: makeSourceHtml([selectedProvince.code, 999], { [selectedProvince.code]: selectedProvince.name })
      })
    };
  }, selectedProvince);
  const rows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "test");
  assert.deepEqual(
    [...requestedSources].sort(),
    V2_SOURCE_CASES.map((source) => source.path).sort(),
    "F4.1, F1.1, F1.2, and F1.3 must all be joined"
  );
  assert.equal(rows.length, 34, "TINH must only contain the 34-code province catalog");
  const selectedRow = rows.find((row) => String(row[1]) === String(selectedProvince.code));
  assert.ok(selectedRow, `catalog includes selected province code ${selectedProvince.code}`);
  [17, 18, 21, 23, 26, 28, 31, 32].forEach((index) => {
    assert.ok(Number(selectedRow[index]) > 0, `code ${selectedProvince.code} must receive real source data at column ${index}`);
  });

  for (const missingSource of V2_SOURCE_CASES) {
    const missingSelectedSandbox = createSandbox(async (requestUrl) => {
      const sourcePath = new URL(requestUrl).pathname;
      const codes = sourcePath === missingSource.path ? [999] : [selectedProvince.code, 999];
      return {
        ok: true,
        text: async () => JSON.stringify({
          data: makeSourceHtml(codes, { [selectedProvince.code]: selectedProvince.name })
        })
      };
    }, selectedProvince);
    await assert.rejects(
      () => missingSelectedSandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "test"),
      (error) => error.message.includes(`Biểu "${missingSource.name}"`)
        && error.message.includes(`mã ${selectedProvince.code} (${selectedProvince.name})`)
        && error.message.includes(`Link curl: https://dkcl.vnpost.vn${missingSource.path}`),
      `${missingSource.name} must reject a missing selected province instead of producing zeros`
    );
  }
}

async function testZeroValuedSelectedProvince(selectedProvince) {
  const sandbox = createSandbox(async () => ({
    ok: true,
    text: async () => JSON.stringify({
      data: makeSourceHtml(
        [selectedProvince.code, 999],
        { [selectedProvince.code]: selectedProvince.name },
        { zeroValues: true }
      )
    })
  }), selectedProvince);
  const rows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "zero-value test");
  const selectedRow = rows.find((row) => String(row[1]) === String(selectedProvince.code));
  assert.ok(selectedRow, "zero-valued selected province record must remain joined");
  [17, 18, 21, 22, 23, 26, 27, 28, 31, 32].forEach((index) => {
    assert.equal(Number(selectedRow[index]), 0, `zero source value must be preserved at column ${index}`);
  });

  const months = sandbox.generateMonthRanges("2026-01", "2026-02");
  const monthlyResults = months.map((monthObj) => ({ monthObj, rows }));
  const dashboardXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  assert.ok(dashboardXml.includes(`MÃ ${selectedProvince.code})`), "zero-valued dashboard keeps the selected province focus");
  assert.match(getRowContaining(dashboardXml, "F4.1 – Chất lượng phát thành công tại bưu cục"), /ss:Type="Number">0<\/Data>/, "zero-valued dashboard displays a valid zero monthly volume");
  assert.match(getRowContaining(dashboardXml, "F4.1 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiRed"/, "valid zero rate is retained and classified against its KPI target");
}

function makeMonthlyRow(code, name, seed) {
  const row = Array(37).fill("");
  row[1] = String(code);
  row[2] = name;
  [17, 21, 26, 31].forEach((index) => { row[index] = String(100 + seed); });
  row[22] = String(100 + seed);
  [18, 23, 28, 32].forEach((index) => { row[index] = String(90 + seed); });
  return row;
}

function makeKpiMonthlyRow(code, name, rates) {
  const row = Array(37).fill("");
  row[1] = String(code);
  row[2] = name;
  const columns = {
    "F4.1": [17, 18],
    "F1.1": [21, 23],
    "F1.2": [26, 28],
    "F1.3": [31, 32]
  };
  Object.entries(columns).forEach(([kpi, [totalCol, onTimeCol]]) => {
    row[totalCol] = "10000";
    row[onTimeCol] = String(Math.round(rates[kpi] * 10000));
    if (kpi === "F1.1") row[22] = row[totalCol];
  });
  return row;
}

function getRowContaining(xml, text) {
  return [...xml.matchAll(/<Row\b[^>]*>[\s\S]*?<\/Row>/g)]
    .map((match) => match[0])
    .find((row) => row.includes(text));
}

function testKpiSsotColorChartAndCommentary() {
  const selectedProvince = { code: 10, name: "TP. Hà Nội" };
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
  const targets = JSON.parse(vm.runInContext(
    "JSON.stringify(Object.fromEntries(Object.values(MULTI_MONTH_METRICS_MAP).map(metric => [metric.shortName, metric.target])))",
    sandbox
  ));
  assert.deepEqual(targets, { "F1.1": 0.95, "F1.2": 0.95, "F1.3": 0.90, "F4.1": 0.90 });

  const statusCases = [
    [0.95, 0.95, "green"],
    [0.9499, 0.95, "yellow"],
    [0.90, 0.95, "yellow"],
    [0.8999, 0.95, "red"],
    [0.90, 0.90, "green"],
    [0.85, 0.90, "yellow"],
    [0.8499, 0.90, "red"],
    [0, 0.90, "red"],
    [null, 0.90, "gray"]
  ];
  statusCases.forEach(([rate, target, expected]) => {
    assert.equal(sandbox.getKpiStatus(rate, target).key, expected, `${rate} against ${target} must be ${expected}`);
  });
  assert.equal(sandbox.getKpiStatus(null, 0.90).styleId, "KpiGray", "N/A uses the gray KPI style");

  const months = sandbox.generateMonthRanges("2026-01", "2026-02");
  const rates = { "F1.1": 0.95, "F1.2": 0.925, "F1.3": 0.84, "F4.1": 0.90 };
  const monthlyResults = months.map((monthObj) => ({
    monthObj,
    rows: [makeKpiMonthlyRow(selectedProvince.code, selectedProvince.name, rates)]
  }));
  const dashboard = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");

  assert.ok(dashboard.includes("COMBO CHART KPI - CỘT TỶ LỆ THỰC TẾ &amp; ĐƯỜNG MỤC TIÊU RIÊNG THEO SSOT"));
  ["F1.1", "F1.2", "F1.3", "F4.1"].forEach((kpi) => {
    assert.ok(getRowContaining(dashboard, `${kpi} – Cột tỷ lệ thực tế`), `${kpi} has an actual column series`);
    assert.ok(getRowContaining(dashboard, `${kpi} – Đường mục tiêu riêng`), `${kpi} has its own target line series`);
  });

  const f11TargetRow = getRowContaining(dashboard, "F1.1 – Đường mục tiêu riêng");
  const f13TargetRow = getRowContaining(dashboard, "F1.3 – Đường mục tiêu riêng");
  assert.equal((f11TargetRow.match(/ss:Type="Number">0\.95<\/Data>/g) || []).length, months.length, "F1.1 target line stays at 95% for every month");
  assert.equal((f13TargetRow.match(/ss:Type="Number">0\.9<\/Data>/g) || []).length, months.length, "F1.3 target line stays at 90% for every month");
  assert.match(getRowContaining(dashboard, "F1.1 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiGreen"/);
  assert.match(getRowContaining(dashboard, "F1.2 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiYellow"/);
  assert.match(getRowContaining(dashboard, "F1.3 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiRed"/);
  // The threshold sentence must name exactly the exported KPIs. Quoting a threshold for a
  // KPI the workbook does not contain tells the reader data is there when it is not.
  const thresholdNote = (dashboard.match(/Ngưỡng SSOT[^;]*/) || [""])[0];
  ["F1.1 = 95%", "F1.2 = 95%", "F1.3 = 90%", "F4.1 = 90%"].forEach((entry) => {
    assert.ok(thresholdNote.includes(entry), `with all four ticked the note states ${entry}`);
  });
  assert.ok(dashboard.includes("ưu tiên xử lý KPI đỏ F1.3"));
  assert.ok(dashboard.includes("lập kế hoạch cải thiện KPI vàng F1.2"));

  const partialDashboard = sandbox.renderV2MultiMonthDashboardWorksheet([
    { monthObj: months[0], rows: [makeKpiMonthlyRow(selectedProvince.code, selectedProvince.name, { "F1.1": 0.95, "F1.2": 0.95, "F1.3": 0.90, "F4.1": 0.90 })] },
    { monthObj: months[1], rows: [] }
  ], months, "TINH");
  assert.match(getRowContaining(partialDashboard, "F1.1 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiGray"/, "missing monthly KPI data renders N/A in gray");
  assert.ok(partialDashboard.includes("⚪ N/A"), "missing monthly KPI data is labeled N/A");

  const dashboardSource = sandbox.renderV2MultiMonthDashboardWorksheet.toString();
  assert.doesNotMatch(dashboardSource, /minCum\s*>=\s*0\.95|MỤC TIÊU CHUNG|đường mục tiêu 95% chung/i);

  const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel">${sandbox.renderWorkbookPropertiesAndStyles()}${dashboard}</Workbook>`;
  assertWellFormedXml(workbook);
  assertAllStyleIdsDefined(workbook);
  assertWorksheetTablesFitDeclaredColumns(dashboard);
}

function relationshipSourcePath(relsPath) {
  if (relsPath === "_rels/.rels") return "";
  const relDir = path.posix.dirname(relsPath);
  const sourceDir = path.posix.dirname(relDir);
  const sourceName = path.posix.basename(relsPath).replace(/\.rels$/, "");
  return path.posix.join(sourceDir, sourceName);
}

function validatePackageRelationships(files) {
  const relationshipParts = Object.keys(files).filter(name => name.endsWith(".rels"));
  relationshipParts.forEach((relsPath) => {
    const relsXml = Buffer.from(files[relsPath]).toString("utf8");
    const sourcePath = relationshipSourcePath(relsPath);
    for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
      const attrs = match[1];
      if (/TargetMode="External"/.test(attrs)) continue;
      const target = attrs.match(/Target="([^"]+)"/)?.[1];
      assert.ok(target, `${relsPath} relationship has a target`);
      const resolved = target.startsWith("/")
        ? target.slice(1)
        : path.posix.normalize(path.posix.join(path.posix.dirname(sourcePath), target));
      assert.ok(files[resolved], `${relsPath} resolves ${target} to existing part ${resolved}`);
    }
  });
}

// The pipeline stores row cells as the API formatted them -- Vietnamese KPI tables use
// "," as the THOUSANDS separator ("1,234,567"), and popup.js reads them with
// cleanNumber/toNumberValue. The legacy SpreadsheetML .xls dashboard is the golden
// oracle: given one monthlyResults, the native .xlsx must report the same total/onTime/
// rate for every KPI, every month, and whichever province is selected.
function makeFormattedKpiMonthlyRow(code, name, volumesByKpi, ratesByKpi) {
  const row = Array(37).fill("");
  row[1] = String(code);
  row[2] = name;
  const columns = { "F4.1": [17, 18], "F1.1": [21, 23], "F1.2": [26, 28], "F1.3": [31, 32] };
  Object.entries(columns).forEach(([kpi, [totalCol, onTimeCol]]) => {
    const volume = volumesByKpi[kpi];
    row[totalCol] = volume.toLocaleString("en-US");
    row[onTimeCol] = Math.round(volume * ratesByKpi[kpi]).toLocaleString("en-US");
    if (kpi === "F1.1") row[22] = row[totalCol];
  });
  return row;
}

function testNativeXlsxMatchesLegacyXlsNumbers(selectedProvince) {
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
  const provinceNames = JSON.parse(vm.runInContext("JSON.stringify(PROVINCE_NAMES)", sandbox));
  const metrics = JSON.parse(vm.runInContext("JSON.stringify(Object.values(MULTI_MONTH_METRICS_MAP))", sandbox));
  const months = sandbox.generateMonthRanges("2026-01", "2026-03");

  // Volumes deliberately straddle the separator boundaries that broke the native
  // exporter: 4 digits (one comma) and 7 digits (two commas -> previously NaN -> 0).
  const focusVolumes = { "F1.1": 1234567, "F1.2": 2500000, "F1.3": 9876, "F4.1": 1234 };
  const otherVolumes = { "F1.1": 1000000, "F1.2": 1500000, "F1.3": 5432, "F4.1": 4321 };
  const focusRates = { "F1.1": 0.96, "F1.2": 0.93, "F1.3": 0.89, "F4.1": 0.91 };
  const otherRates = { "F1.1": 0.95, "F1.2": 0.95, "F1.3": 0.90, "F4.1": 0.90 };

  const monthlyResults = months.map((monthObj, monthIndex) => ({
    monthObj,
    rows: Object.entries(provinceNames).map(([code, name], unitIndex) => {
      const isFocus = String(code) === String(selectedProvince.code);
      const scale = 1 + monthIndex * 0.1;
      const volumes = Object.fromEntries(
        Object.entries(isFocus ? focusVolumes : otherVolumes)
          .map(([kpi, base]) => [kpi, Math.round(base * scale)])
      );
      const row = makeFormattedKpiMonthlyRow(code, name, volumes, isFocus ? focusRates : otherRates);
      row[0] = String(unitIndex + 1);
      return row;
    })
  }));

  const native = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList: months,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus,
    toNumberValue: sandbox.toNumberValue
  });

  const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  const focusRow = monthlyResults.map(({ rows }) => rows.find((row) => String(row[1]) === String(selectedProvince.code)));
  focusRow.forEach((row, monthIndex) => assert.ok(row, `focus row exists for month ${monthIndex + 1}`));

  metrics.forEach((metric) => {
    // --- Oracle 1: the legacy .xls dashboard's own monthly values ---
    // Its focus KPI row emits: index, name, then five cells per month
    // (volume, rate, rank, rate delta, trend icon).
    const legacyLine = legacyXml
      .split(/<\/Row>/)
      .find((chunk) => chunk.includes(`${metric.shortName} – `));
    assert.ok(legacyLine, `legacy .xls renders a trend row for ${metric.shortName}`);
    const legacyNumbers = [...legacyLine.matchAll(/ss:Type="Number">([^<]*)<\/Data>/g)].map((match) => Number(match[1]));
    const legacyRates = months.map((_, monthIndex) => legacyNumbers[monthIndex === 0 ? 2 : 4 + (monthIndex - 1) * 3]);

    const nativeMetric = native.model.metrics.find((item) => item.shortName === metric.shortName);
    assert.ok(nativeMetric, `native .xlsx model contains ${metric.shortName}`);

    assert.deepEqual(
      nativeMetric.points.map((point) => point.rate),
      legacyRates,
      `${metric.shortName}: native .xlsx monthly rates must equal the legacy .xls rates (province ${selectedProvince.code})`
    );

    // --- Oracle 2: volumes, read with popup.js's own numeric SSOT (toNumberValue) ---
    const expectedTotal = focusRow.reduce((sum, row) => sum + sandbox.toNumberValue(row[metric.totalCol]), 0);
    const expectedOnTime = focusRow.reduce((sum, row) => sum + sandbox.toNumberValue(row[metric.onTimeCol]), 0);
    assert.ok(expectedTotal > 0, `${metric.shortName}: fixture volume is non-zero`);
    assert.equal(nativeMetric.total, expectedTotal, `${metric.shortName}: native .xlsx total must match the legacy numeric reading, not 0`);
    assert.equal(nativeMetric.onTime, expectedOnTime, `${metric.shortName}: native .xlsx onTime must match the legacy numeric reading`);
    nativeMetric.points.forEach((point, monthIndex) => {
      assert.equal(point.volume, sandbox.toNumberValue(focusRow[monthIndex][metric.totalCol]), `${metric.shortName} month ${monthIndex + 1}: volume parity`);
      assert.equal(point.onTime, sandbox.toNumberValue(focusRow[monthIndex][metric.onTimeCol]), `${metric.shortName} month ${monthIndex + 1}: onTime parity`);
    });
  });

  // The exporter's built-in fallback (used when popup.js does not inject its reader) must
  // follow the same numeric convention, otherwise the SSOT is only half-closed.
  const nativeFallback = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList: months,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus
  });
  metrics.forEach((metric) => {
    const injected = native.model.metrics.find((item) => item.shortName === metric.shortName);
    const fallback = nativeFallback.model.metrics.find((item) => item.shortName === metric.shortName);
    assert.equal(fallback.total, injected.total, `${metric.shortName}: exporter fallback reader must match the injected popup.js reader`);
    assert.equal(fallback.cumulativeRate, injected.cumulativeRate, `${metric.shortName}: exporter fallback rate must match the injected reader`);
  });

  // The matrix sheet must agree with the same oracle for a non-focus province too.
  const otherCode = String(selectedProvince.code) === "10" ? "53" : "10";
  const otherEntry = native.model.matrix.find((entry) => String(entry.code) === otherCode);
  assert.ok(otherEntry, `matrix contains comparison province ${otherCode}`);
  metrics.forEach((metric, metricIndex) => {
    const expected = monthlyResults.reduce((sum, { rows }) => {
      const row = rows.find((item) => String(item[1]) === otherCode);
      return sum + sandbox.toNumberValue(row[metric.totalCol]);
    }, 0);
    assert.equal(otherEntry.metrics[metricIndex].total, expected, `${metric.shortName}: matrix total parity for province ${otherCode}`);
  });
}

// PO decision: the native .xlsx exports ONLY the ticked KPIs, exactly like the legacy
// .xls. An unticked KPI is never fetched, so a card/series/matrix column for it would be
// a fake 0 -- it must be absent from the package entirely, not rendered as zero.
function testNativeXlsxExportsOnlySelectedKpis() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const allLabels = [
    "F1.1 – Nội tỉnh",
    "F1.2 – Thu gom bưu gửi đi liên tỉnh",
    "F1.3 – Chất lượng phát bưu gửi liên tỉnh",
    "F4.1 – Chất lượng phát thành công tại bưu cục"
  ];

  for (let selectedCount = 1; selectedCount <= 4; selectedCount++) {
    const chosenLabels = allLabels.slice(0, selectedCount);
    const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
    // Drive the real selection path (getSelectedMetricLabels -> getActiveMultiMonthMetrics).
    vm.runInContext(
      `document.querySelectorAll = (selector) => selector.includes(":checked")
         ? ${JSON.stringify(chosenLabels)}.map((value) => ({ value, checked: true }))
         : [];`,
      sandbox
    );
    const activeMetrics = JSON.parse(vm.runInContext("JSON.stringify(getActiveMultiMonthMetrics())", sandbox));
    assert.equal(activeMetrics.length, selectedCount, `${selectedCount} KPI ticked -> ${selectedCount} active metrics`);

    const chosenShortNames = activeMetrics.map((metric) => metric.shortName);
    const omittedShortNames = ["F1.1", "F1.2", "F1.3", "F4.1"].filter((name) => !chosenShortNames.includes(name));

    const provinceNames = JSON.parse(vm.runInContext("JSON.stringify(PROVINCE_NAMES)", sandbox));
    const months = sandbox.generateMonthRanges("2026-01", "2026-02");
    const volumes = { "F1.1": 1234567, "F1.2": 2500000, "F1.3": 9876, "F4.1": 1234 };
    const rates = { "F1.1": 0.96, "F1.2": 0.93, "F1.3": 0.89, "F4.1": 0.91 };
    const monthlyResults = months.map((monthObj) => ({
      monthObj,
      rows: Object.entries(provinceNames).map(([code, name], unitIndex) => {
        const row = makeFormattedKpiMonthlyRow(code, name, volumes, rates);
        row[0] = String(unitIndex + 1);
        return row;
      })
    }));

    const result = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
      monthlyResults,
      monthsList: months,
      tuyChonGR: "TINH",
      selectedProvinceCode: selectedProvince.code,
      selectedProvinceName: selectedProvince.name,
      metrics: activeMetrics,
      columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
      classifyKpi: sandbox.getKpiStatus,
      toNumberValue: sandbox.toNumberValue
    });

    assert.equal(result.model.metrics.length, selectedCount, `model carries exactly the ${selectedCount} ticked KPIs`);
    const files = sandbox.fflate.unzipSync(result.bytes);

    // One chart part + one drawing relationship per ticked KPI, and no more.
    const chartParts = Object.keys(files).filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
    assert.equal(chartParts.length, selectedCount, `${selectedCount} ticked KPI -> ${selectedCount} chart parts`);
    const drawingRels = Buffer.from(files["xl/drawings/_rels/drawing1.xml.rels"]).toString("utf8");
    assert.equal(
      (drawingRels.match(/Type="[^"]*\/chart"/g) || []).length,
      selectedCount,
      `drawing declares exactly ${selectedCount} chart relationships`
    );
    const contentTypes = Buffer.from(files["[Content_Types].xml"]).toString("utf8");
    assert.equal(
      (contentTypes.match(/\/xl\/charts\/chart\d+\.xml/g) || []).length,
      selectedCount,
      `[Content_Types].xml overrides exactly ${selectedCount} charts`
    );
    validatePackageRelationships(files);

    // Sheet geometry must shrink with the selection: DuLieu_Chart uses 4 columns per KPI
    // after the month column, the matrix 3 per KPI after code+name. Leaving the old fixed
    // 16/12-column layout would declare empty columns for KPIs that were never exported.
    const chartDataSheet = Buffer.from(files["xl/worksheets/sheet2.xml"]).toString("utf8");
    assert.equal(
      (chartDataSheet.match(/<col /g) || []).length,
      1 + 4 * selectedCount,
      `DuLieu_Chart declares 1 + 4x${selectedCount} columns`
    );
    const matrixSheet = Buffer.from(files["xl/worksheets/sheet3.xml"]).toString("utf8");
    assert.equal(
      (matrixSheet.match(/<col /g) || []).length,
      2 + 3 * selectedCount,
      `matrix declares 2 + 3x${selectedCount} columns`
    );

    // No trace of an unticked KPI anywhere in the package -- no card row, no
    // DuLieu_Chart block, no matrix column, no chart series title.
    const packageText = Object.entries(files)
      .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
      .map(([, bytes]) => Buffer.from(bytes).toString("utf8"))
      .join("\n");
    omittedShortNames.forEach((shortName) => {
      assert.ok(
        !packageText.includes(shortName),
        `unticked ${shortName} must not appear anywhere in the package (no fake 0 card/series/column)`
      );
    });
    chosenShortNames.forEach((shortName) => {
      assert.ok(packageText.includes(shortName), `ticked ${shortName} is present`);
    });

    // Every KPI that IS exported must carry real numbers, never a fake 0.
    result.model.metrics.forEach((metric) => {
      assert.ok(metric.total > 0, `${metric.shortName}: exported KPI has a real total, not 0`);
      assert.ok(metric.cumulativeRate > 0, `${metric.shortName}: exported KPI has a real rate, not 0`);
    });

    // The legacy .xls must agree: it renders the same ticked KPIs and no others.
    const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
    chosenShortNames.forEach((shortName) => {
      assert.ok(legacyXml.includes(`${shortName} – `), `legacy .xls also renders ticked ${shortName}`);
    });
    omittedShortNames.forEach((shortName) => {
      assert.ok(!legacyXml.includes(`${shortName} – `), `legacy .xls also omits unticked ${shortName}`);
    });
  }
}

function testNativeXlsxPackageAndCharts() {
  const selectedProvince = { code: 10, name: "TP. Hà Nội" };
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
  const provinceNames = JSON.parse(vm.runInContext("JSON.stringify(PROVINCE_NAMES)", sandbox));
  const metrics = JSON.parse(vm.runInContext("JSON.stringify(Object.values(MULTI_MONTH_METRICS_MAP))", sandbox));
  const months = sandbox.generateMonthRanges("2026-01", "2026-03");
  const focusRates = { "F1.1": 0.96, "F1.2": 0.93, "F1.3": 0.89, "F4.1": 0.91 };
  const monthlyResults = months.map((monthObj, monthIndex) => ({
    monthObj,
    rows: Object.entries(provinceNames).map(([code, name], unitIndex) => {
      const rates = String(code) === String(selectedProvince.code)
        ? focusRates
        : { "F1.1": 0.95, "F1.2": 0.95, "F1.3": 0.90, "F4.1": 0.90 };
      const row = makeKpiMonthlyRow(code, name, rates);
      row[0] = String(unitIndex + 1);
      [17, 21, 26, 31].forEach((col) => { row[col] = String(10000 + monthIndex * 100 + unitIndex); });
      [18, 23, 28, 32].forEach((col) => {
        const kpi = col === 18 ? "F4.1" : col === 23 ? "F1.1" : col === 28 ? "F1.2" : "F1.3";
        row[col] = String(Math.round(Number(row[col === 18 ? 17 : col === 23 ? 21 : col === 28 ? 26 : 31]) * rates[kpi]));
      });
      return row;
    })
  }));

  const result = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList: months,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus
  });

  assert.equal(result.bytes[0], 0x50, ".xlsx starts with P");
  assert.equal(result.bytes[1], 0x4B, ".xlsx starts with PK ZIP magic bytes");
  assert.equal(result.model.units.length, 34, "native TINH package retains exactly 34 provinces");
  assert.equal(result.model.selectedCode, "10", "native dashboard focus is dynamic");

  // Adding the BC ranking must not move the province ranking: TINH still ranks across the
  // 34 provinces, per KPI and per month, and says so.
  assert.match(result.model.rankingSample, /34 Bưu điện Tỉnh\/TP toàn quốc/, "TINH keeps the 34-province ranking sample");
  assert.equal(result.model.matrix[0].monthlyByMetric, undefined, "the per-post-office monthly breakdown is BC-only");
  result.model.metrics.forEach((metric) => {
    metric.points.forEach((point, monthIndex) => {
      if (!point.hasData) return;
      assert.equal(point.rankCount, 34, `${metric.shortName} month ${monthIndex + 1}: TINH ranks against all 34 provinces`);
      assert.ok(point.rank >= 1 && point.rank <= 34, `${metric.shortName} month ${monthIndex + 1}: rank stays inside the 34-province scale`);
    });
  });

  const files = sandbox.fflate.unzipSync(result.bytes);
  Object.entries(files)
    .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
    .forEach(([name, bytes]) => {
      const xml = Buffer.from(bytes).toString("utf8");
      assertWellFormedXml(xml);
      assertNoDuplicateXmlAttributes(xml, name);
    });
  const requiredParts = [
    "[Content_Types].xml",
    "xl/workbook.xml",
    "xl/worksheets/sheet1.xml",
    "xl/drawings/drawing1.xml",
    "xl/drawings/_rels/drawing1.xml.rels"
  ];
  requiredParts.forEach(part => assert.ok(files[part], `native .xlsx includes ${part}`));
  const chartParts = Object.keys(files).filter(name => /^xl\/charts\/chart\d+\.xml$/.test(name)).sort();
  assert.equal(chartParts.length, 4, "native .xlsx contains one chart part for each KPI");
  validatePackageRelationships(files);

  const targetByKpi = { "F1.1": 0.95, "F1.2": 0.95, "F1.3": 0.90, "F4.1": 0.90 };
  chartParts.forEach((chartPart) => {
    const chartXml = Buffer.from(files[chartPart]).toString("utf8");
    const kpi = Object.keys(targetByKpi).find(code => chartXml.includes(`${code}: Sản lượng`));
    assert.ok(kpi, `${chartPart} identifies its KPI`);
    assert.equal((chartXml.match(/<c:barChart>/g) || []).length, 1, `${kpi} has a native column chart group`);
    assert.equal((chartXml.match(/<c:lineChart>/g) || []).length, 1, `${kpi} has a native line chart group`);
    const barBlock = chartXml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/)?.[0] || "";
    const lineBlock = chartXml.match(/<c:lineChart>[\s\S]*?<\/c:lineChart>/)?.[0] || "";
    assert.equal((barBlock.match(/<c:ser>/g) || []).length, 1, `${kpi} column group contains volume series`);
    assert.equal((lineBlock.match(/<c:ser>/g) || []).length, 2, `${kpi} line group contains rate and target series`);
    assert.ok(lineBlock.includes(`${kpi} Tỷ lệ`), `${kpi} has a rate line`);
    assert.ok(lineBlock.includes(`${kpi} Mục tiêu ${(targetByKpi[kpi] * 100).toFixed(0)}%`), `${kpi} has its own target line`);
    const secondaryAxis = [...chartXml.matchAll(/<c:valAx>[\s\S]*?<\/c:valAx>/g)]
      .map(match => match[0])
      .find(axis => axis.includes('<c:axPos val="r"/>'));
    assert.ok(secondaryAxis, `${kpi} has a secondary value axis`);
    assert.ok(secondaryAxis.includes('<c:min val="0"/>') && secondaryAxis.includes('<c:max val="1"/>'), `${kpi} secondary axis spans 0-100%`);
    assert.ok(secondaryAxis.includes('formatCode="0%"'), `${kpi} secondary axis uses percentage formatting`);
    const targetPoint = `<c:v>${targetByKpi[kpi]}</c:v>`;
    assert.ok((chartXml.match(new RegExp(targetPoint.replace(".", "\\."), "g")) || []).length >= months.length, `${kpi} target cache uses its SSOT threshold for every month`);
  });

  const drawingXml = Buffer.from(files["xl/drawings/drawing1.xml"]).toString("utf8");
  assert.equal((drawingXml.match(/<c:chart\b/g) || []).length, 4, "dashboard drawing embeds four native chart references");
  const contentTypes = Buffer.from(files["[Content_Types].xml"]).toString("utf8");
  chartParts.forEach((part) => assert.ok(contentTypes.includes(`PartName="/${part}"`), `${part} is declared in [Content_Types].xml`));

  const popupHtml = fs.readFileSync(path.join(__dirname, "..", "popup.html"), "utf8");
  const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "manifest.json"), "utf8"));
  assert.equal(manifest.manifest_version, 3, "extension remains MV3");
  assert.equal(manifest.content_security_policy.extension_pages, "script-src 'self'; object-src 'self'", "MV3 extension CSP only permits local scripts");
  assert.ok(popupHtml.indexOf('src="vendor/fflate.min.js"') < popupHtml.indexOf('src="xlsx-export.js"'), "vendored fflate loads before the native exporter");
  assert.ok(popupHtml.indexOf('src="xlsx-export.js"') < popupHtml.indexOf('src="popup.js"'), "native exporter loads before popup logic");
  assert.doesNotMatch(popupHtml, /<script[^>]+src="https?:\/\//i, "MV3 page loads no remote script");
  assert.doesNotMatch(fflateSource, /eval\(|new Function|importScripts|new Worker|createObjectURL|https?:\/\//, "tree-shaken fflate bundle contains no dynamic or remote code");
  assert.ok(manifest.web_accessible_resources[0].resources.includes("vendor/fflate.min.js"), "fflate is packaged locally for offline use");
  assert.ok(manifest.web_accessible_resources[0].resources.includes("xlsx-export.js"), "native exporter is packaged locally");
  assert.equal(sandbox.isNativeV2XlsxExportEnabled(), true, "native .xlsx is enabled by default");
  const rollbackSandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince, { dkclNativeXlsxExportEnabled: "false" });
  assert.equal(rollbackSandbox.isNativeV2XlsxExportEnabled(), false, "feature flag can rollback to legacy SpreadsheetML .xls");
  assert.match(popupSource, /\.xlsx`;/, "default native export filename uses .xlsx");
  assert.match(popupSource, /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/, "native export uses the .xlsx MIME type");

  if (process.env.DKCL_XLSX_SAMPLE_PATH) {
    const samplePath = path.resolve(process.env.DKCL_XLSX_SAMPLE_PATH);
    fs.mkdirSync(path.dirname(samplePath), { recursive: true });
    fs.writeFileSync(samplePath, Buffer.from(result.bytes));
    console.log(`Sample xlsx written: ${samplePath}`);
  }
  console.log(`Native XLSX package validated: ${Object.keys(files).length} parts, ${chartParts.length} combo charts, all relationships resolved`);
}

function testDynamicDashboardFocus(selectedProvince) {
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
  const months = sandbox.generateMonthRanges("2026-01", "2026-02");
  const otherProvince = selectedProvince.code === 10
    ? { code: 53, name: "Thừa Thiên Huế" }
    : { code: 10, name: "TP. Hà Nội" };
  const monthlyResults = months.map((monthObj, monthIndex) => ({
    monthObj,
    rows: [
      makeMonthlyRow(selectedProvince.code, selectedProvince.name, monthIndex),
      makeMonthlyRow(otherProvince.code, otherProvince.name, monthIndex + 5)
    ]
  }));
  const xml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  assert.match(xml, new RegExp(`ĐƠN VỊ TRỌNG ĐIỂM QUAN SÁT: ${selectedProvince.name} \\(Mã ${selectedProvince.code}\\)`));
  assert.ok(xml.includes(`MÃ ${selectedProvince.code})`), "dashboard title displays selected province code");
  assert.ok(xml.includes(`HIGHLIGHT DÒNG ${selectedProvince.name.toUpperCase()} ★`), "matrix title displays selected province name");
  assert.ok(xml.includes(`${selectedProvince.name} ★`), "selected province row is highlighted");
}

function testValidationHasNoHardcodedFocusProvince() {
  const sandbox = createSandbox(async () => { throw new Error("not used"); });
  const validationSource = sandbox.fetchV2ReportRows.toString();
  assert.doesNotMatch(validationSource, /provinceCodeInt\s*===\s*53|mã tỉnh 53|Thừa Thiên Huế/);
  assert.doesNotMatch(validationSource, /(total|onTime)[^\n]*>\s*0|>\s*0[^\n]*(total|onTime)/);
  assert.match(validationSource, /getSelectedProvinceCode\(\)/);
  assert.match(validationSource, /getSelectedProvinceName\(\)/);

  // The legacy BC report does drop post offices without data, but that filter must stay
  // BC-only: a province row with a zero volume is real data and must never disappear.
  assert.match(validationSource, /tuyChonGR === "BC"\s*\n?\s*\? baseRows\.filter\(\(row\) => isBcOperatingRowWithData/,
    "the row filter is gated on BC and delegated to the BC-only helper");
  assert.match(sandbox.isBcOperatingRowWithData.toString(), /BCVH/, "the BC helper keeps the legacy BCVH rule");

  // Both exports must select KPIs through the one SSOT helper. Passing the whole
  // MULTI_MONTH_METRICS_MAP to the native exporter would resurrect fake 0 cards for
  // KPIs the user never ticked (and never fetched).
  const exportSource = sandbox.exportReportV2MultiMonth.toString();
  assert.match(exportSource, /metrics:\s*getActiveMultiMonthMetrics\(tuyChonGR\)/, "native exporter receives only the ticked KPIs, scoped to the report level");
  assert.doesNotMatch(exportSource, /Object\.values\(MULTI_MONTH_METRICS_MAP\)/, "native exporter is not handed every KPI unconditionally");
  assert.match(
    sandbox.renderV2MultiMonthDashboardWorksheet.toString(),
    /getActiveMultiMonthMetrics\(tuyChonGR\)/,
    "legacy .xls dashboard selects KPIs through the same SSOT helper"
  );
}

function testComparisonDashboardXml() {
  const sandbox = createSandbox(async () => { throw new Error("not used"); });
  const currentRows = makeComparisonRows("kỳ này");
  const compareRows = makeComparisonRows("kỳ trước");
  ["TINH", "BC"].forEach((group) => {
    const comparisonRows = sandbox.buildV2ComparisonRows(currentRows, compareRows);
    const dashboard = sandbox.renderV2ComparisonDashboardWorksheet(
      currentRows, compareRows, group,
      new Date(2026, 7, 25), new Date(2026, 7, 25),
      new Date(2026, 7, 24), new Date(2026, 7, 24)
    );
    const workbook = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel">${sandbox.renderWorkbookPropertiesAndStyles()}${dashboard}${sandbox.renderV2DataWorksheet("Du lieu dau ky V2", currentRows, "Dữ liệu V2 đầu kỳ")}${sandbox.renderV2DataWorksheet("Ky so sanh V2", compareRows, "Dữ liệu V2 kỳ so sánh")}${sandbox.renderV2ComparisonWorksheet(comparisonRows, new Date(2026, 7, 25), new Date(2026, 7, 25), new Date(2026, 7, 24), new Date(2026, 7, 24))}</Workbook>`;
    assertWellFormedXml(workbook);
    assertAllStyleIdsDefined(workbook);
    assertDashboardTableIsValid(dashboard);
  });
}

function makeGoldenF13SourceHtml(code, name) {
  const cells = Array(20).fill("");
  cells[1] = String(code);
  cells[2] = name;
  // Golden F1.3: total and KPI-2026 numerator are at 10 and 18. The old wrong
  // locations deliberately contain the known bad output so a mapping mutation is visible.
  cells[9] = "126228";
  cells[10] = "102596";
  cells[11] = "99514";
  cells[18] = "62848";
  return `<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`;
}

async function testGoldenF13MutationFails() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const sandbox = createSandbox(async () => ({
    ok: true,
    text: async () => JSON.stringify({ data: makeGoldenF13SourceHtml(selectedProvince.code, selectedProvince.name) })
  }), selectedProvince);
  vm.runInContext('getSelectedMetricLabels = () => ["F1.3 – Chất lượng phát bưu gửi liên tỉnh"];', sandbox);
  const assertGoldenRow = (row) => assert.deepEqual(
    [row[31], row[32], row[33]],
    ["102596", "62848", String(62848 / 102596)],
    "F1.3 golden mapping uses KPI-2026 cells"
  );

  const correctRows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "Tháng 01/2026");
  const correct = correctRows.find((row) => String(row[1]) === "53");
  assertGoldenRow(correct);

  vm.runInContext('V2_SOURCE_CONFIGS.find(config => config.key === "phatLienTinhF13").values = { total: 9, onTime: 11 };', sandbox);
  const mutatedRows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "Tháng 01/2026");
  const mutated = mutatedRows.find((row) => String(row[1]) === "53");
  assert.throws(() => assertGoldenRow(mutated), "mutation back to the current bad offsets must fail the golden assertion");
}

// Golden mapping from the supplied old extension. F1.3 additionally uses the PO's
// confirmed January-2026 values for province 53.
const GOLDEN_MAPPING_CASE = {
  "/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc": { shortName: "F4.1", cells: { 10: "200000", 27: "190000" }, total: 200000, onTime: 190000 },
  "/kpi/chat-luong-toan-trinh-buu-gui-noi-tinh": { shortName: "F1.1", cells: { 11: "120000", 18: "100000", 19: "95000" }, rawTotal: 120000, total: 100000, onTime: 95000 },
  "/kpi/chat-luong-thu-gom-buu-lien-tinh": { shortName: "F1.2", cells: { 7: "1358542", 8: "1350000", 9: "1277357" }, total: 1358542, onTime: 1277357 },
  "/kpi/chat-luong-phat-buu-gui-lien-tinh": { shortName: "F1.3", cells: { 10: "102596", 18: "62848", 9: "126228", 11: "99514" }, total: 102596, onTime: 62848, expectedRate: 0.6126 }
};

function makeCaptureSourceHtml(code, name, cellOverrides) {
  const cells = Array(28).fill("");
  cells[1] = String(code);
  cells[2] = name;
  Object.entries(cellOverrides).forEach(([index, value]) => { cells[Number(index)] = value; });
  return `<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`;
}

// Source HTML -> month sheet -> monthlyResults -> DuLieu_Chart -> both dashboards.
async function testGoldenMappingFlowsThroughPipeline() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const selectedLabels = ["F4.1 – Chất lượng phát thành công tại bưu cục", "F1.1 – Nội tỉnh", "F1.2 – Thu gom bưu gửi đi liên tỉnh", "F1.3 – Chất lượng phát bưu gửi liên tỉnh"];
  const sandbox = createSandbox(async (requestUrl) => {
    const shape = GOLDEN_MAPPING_CASE[new URL(requestUrl).pathname];
    assert.ok(shape, "only golden V2 KPI endpoints may be fetched");
    return {
      ok: true,
      text: async () => JSON.stringify({
        data: makeCaptureSourceHtml(selectedProvince.code, selectedProvince.name, shape.cells)
      })
    };
  }, selectedProvince);

  vm.runInContext(
    `getSelectedMetricLabels = () => ${JSON.stringify(selectedLabels)};`,
    sandbox
  );

  const rows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "Tháng 01/2026");
  const focusRow = rows.find((row) => String(row[1]) === String(selectedProvince.code));
  assert.ok(focusRow, "the golden mapping keeps the focus province");

  const months = sandbox.generateMonthRanges("2026-01", "2026-01");
  const monthlyResults = months.map((monthObj) => ({ monthObj, rows }));
  const metrics = JSON.parse(vm.runInContext("JSON.stringify(getActiveMultiMonthMetrics())", sandbox));
  assert.equal(metrics.length, 4, "only the ticked KPIs reach the dashboard");

  const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  const native = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList: months,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus,
    toNumberValue: sandbox.toNumberValue
  });
  const packageFiles = sandbox.fflate.unzipSync(native.bytes);
  assert.ok(packageFiles["xl/worksheets/sheet2.xml"], "package contains the DuLieu_Chart sheet");
  const chartSheet = Buffer.from(packageFiles["xl/worksheets/sheet2.xml"]).toString("utf8");
  assert.ok(chartSheet.includes("Sản lượng"), "sheet2 is DuLieu_Chart");

  Object.values(GOLDEN_MAPPING_CASE).forEach((shape) => {
    const metric = metrics.find((item) => item.shortName === shape.shortName);
    const expectedRate = shape.onTime / shape.total;
    if (shape.expectedRate) {
      assert.ok(
        Math.abs(expectedRate - shape.expectedRate) < 0.0001,
        `${shape.shortName}: fixture must reproduce the PO-confirmed golden ratio`
      );
    }

    // 1. month sheet
    assert.equal(sandbox.toNumberValue(focusRow[metric.totalCol]), shape.total, `${shape.shortName}: month sheet denominator comes from the golden offset`);
    assert.equal(sandbox.toNumberValue(focusRow[metric.onTimeCol]), shape.onTime, `${shape.shortName}: month sheet numerator comes from the golden offset`);
    if (shape.rawTotal != null) assert.equal(sandbox.toNumberValue(focusRow[21]), shape.rawTotal, "F1.1 retains raw total separately from its denominator");
    const rateColumn = { "F4.1": 19, "F1.1": 24, "F1.2": 29, "F1.3": 33 }[shape.shortName];
    assert.equal(Number(focusRow[rateColumn]), expectedRate, `${shape.shortName}: month sheet rate`);

    // 2. monthlyResults -> native model
    const nativeMetric = native.model.metrics.find((item) => item.shortName === shape.shortName);
    assert.ok(nativeMetric, `${shape.shortName}: present in the native model`);
    assert.equal(nativeMetric.total, shape.total, `${shape.shortName}: native total`);
    assert.equal(nativeMetric.onTime, shape.onTime, `${shape.shortName}: native onTime`);
    assert.equal(nativeMetric.cumulativeRate, expectedRate, `${shape.shortName}: native cumulative rate`);
    assert.equal(nativeMetric.points[0].rate, expectedRate, `${shape.shortName}: native monthly rate`);

    // 3. DuLieu_Chart carries the same volumes the rate was derived from
    assert.ok(chartSheet.includes(`>${shape.total}<`), `${shape.shortName}: DuLieu_Chart carries the denominator`);
    assert.ok(chartSheet.includes(`>${shape.onTime}<`), `${shape.shortName}: DuLieu_Chart carries the numerator`);

    // 4. legacy .xls dashboard renders the same rate
    const legacyLine = legacyXml.split(/<\/Row>/).find((chunk) => chunk.includes(`${shape.shortName} – `));
    assert.ok(legacyLine, `${shape.shortName}: legacy dashboard renders a trend row`);
    const legacyNumbers = [...legacyLine.matchAll(/ss:Type="Number">([^<]*)<\/Data>/g)].map((match) => Number(match[1]));
    assert.equal(legacyNumbers[2], expectedRate, `${shape.shortName}: legacy dashboard monthly rate matches the month sheet`);
  });

  // No zeroed denominator may survive anywhere in the chain.
  assert.doesNotMatch(legacyXml, /SL: 0 bg/, "no KPI may render a zero volume after golden mapping");
}

async function testF12ProductionRecordAcceptsSubTotalAboveTotal() {
  const selectedProvince = { code: 70, name: "TP Hồ Chí Minh" };
  const source = { total: 1350604, subTotal: 1350654, onTime: 1316616, sourceRate: 0.9748 };
  const cells = Array(11).fill("");
  cells[1] = String(selectedProvince.code);
  cells[2] = selectedProvince.name;
  cells[7] = String(source.total);
  cells[8] = String(source.subTotal);
  cells[9] = String(source.onTime);
  cells[10] = "97.48%";
  const html = `<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`;
  const sandbox = createSandbox(async () => ({ ok: true, text: async () => JSON.stringify({ data: html }) }), selectedProvince);
  vm.runInContext('getSelectedMetricLabels = () => ["F1.2 – Thu gom bưu gửi đi liên tỉnh"];', sandbox);

  const rows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 7, 1), new Date(2026, 7, 31), {}, "Tháng 08/2026");
  const focus = rows.find((row) => String(row[1]) === "70");
  const expectedRate = source.onTime / source.total;
  assert.ok(focus, "production F1.2 row is retained");
  assert.deepEqual([focus[26], focus[27], focus[28]], [String(source.total), String(source.subTotal), String(source.onTime)]);
  assert.equal(Number(focus[29]), expectedRate, "F1.2 rate uses golden denominator total, not subTotal");
  assert.ok(Math.abs(expectedRate - source.sourceRate) < 0.0001, "onTime / total rounds to the source rate 97.48%");

  const months = sandbox.generateMonthRanges("2026-08", "2026-08");
  const monthlyResults = months.map((monthObj) => ({ monthObj, rows }));
  const metrics = JSON.parse(vm.runInContext("JSON.stringify(getActiveMultiMonthMetrics())", sandbox));
  const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  const native = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList: months,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus,
    toNumberValue: sandbox.toNumberValue
  });
  const nativeMetric = native.model.metrics.find((metric) => metric.shortName === "F1.2");
  assert.deepEqual([nativeMetric.total, nativeMetric.onTime, nativeMetric.cumulativeRate], [source.total, source.onTime, expectedRate], "production F1.2 values reach native XLSX");
  assert.ok(legacyXml.includes(String(expectedRate)), "production F1.2 rate reaches legacy dashboard");

  const parsed = sandbox.parseV2ApiRows(html, vm.runInContext('V2_SOURCE_CONFIGS.find(config => config.key === "thuGomLienTinhF12")', sandbox), "TINH", new Set(["70"]));
  assert.doesNotThrow(() => sandbox.assertV2RecordIntegrity(parsed[0], vm.runInContext('V2_SOURCE_CONFIGS.find(config => config.key === "thuGomLienTinhF12")', sandbox), {}));
  vm.runInContext(`
    originalF12Integrity = assertV2RecordIntegrity;
    assertV2RecordIntegrity = (record, config, context) => {
      originalF12Integrity(record, config, context);
      if (Number(record.subTotal) > Number(record.total)) throw new Error("legacy unproven subTotal guard");
    };
  `, sandbox);
  await assert.rejects(
    () => sandbox.fetchV2ReportRows("TINH", new Date(2026, 7, 1), new Date(2026, 7, 31), {}, "Tháng 08/2026"),
    /legacy unproven subTotal guard/,
    "restoring subTotal > total => throw must fail the approved production-record pipeline"
  );
  vm.runInContext('assertV2RecordIntegrity = originalF12Integrity;', sandbox);

  vm.runInContext('V2_SOURCE_CONFIGS.find(config => config.key === "thuGomLienTinhF12").values = { total: 5, subTotal: 6, onTime: 7 };', sandbox);
  await assert.rejects(
    () => sandbox.fetchV2ReportRows("TINH", new Date(2026, 7, 1), new Date(2026, 7, 31), {}, "Tháng 08/2026"),
    /cột total trống/,
    "mutating the golden F1.2 offsets must fail the production-record regression"
  );
}

async function testGoldenMappingSupportsTwoPeriodComparison() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const requestedPeriods = [];
  const sandbox = createSandbox(async (requestUrl) => {
    const url = new URL(requestUrl);
    requestedPeriods.push([url.searchParams.get("iFrom"), url.searchParams.get("iTo")]);
    const shape = GOLDEN_MAPPING_CASE[url.pathname];
    return { ok: true, text: async () => JSON.stringify({ data: makeCaptureSourceHtml(selectedProvince.code, selectedProvince.name, shape.cells) }) };
  }, selectedProvince);
  vm.runInContext('getSelectedMetricLabels = () => ["F4.1 – Chất lượng phát thành công tại bưu cục", "F1.1 – Nội tỉnh", "F1.2 – Thu gom bưu gửi đi liên tỉnh", "F1.3 – Chất lượng phát bưu gửi liên tỉnh"];', sandbox);

  const currentRows = await sandbox.fetchV2ReportRows("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31), {}, "kỳ báo cáo");
  const compareRows = await sandbox.fetchV2ReportRows("TINH", new Date(2025, 11, 1), new Date(2025, 11, 31), {}, "kỳ so sánh");
  assert.equal(requestedPeriods.length, 8, "two-period comparison performs one complete golden source pass per period");
  const focus = currentRows.find((row) => String(row[1]) === "53");
  assert.deepEqual([focus[31], focus[32], focus[33]], ["102596", "62848", String(62848 / 102596)], "comparison current period retains F1.3 golden values");
  const comparisonRows = sandbox.buildV2ComparisonRows(currentRows, compareRows);
  const f13 = comparisonRows.find((row) => String(row.code) === "53").metrics.find((metric) => metric.label.includes("F1.3"));
  assert.equal(f13.currentRate, 62848 / 102596, "comparison model reads the same F1.3 rate as the source row");
  const worksheet = sandbox.renderV2ComparisonWorksheet(comparisonRows, new Date(2026, 0, 1), new Date(2026, 0, 31), new Date(2025, 11, 1), new Date(2025, 11, 31));
  assert.ok(worksheet.includes(String(62848 / 102596)), "comparison worksheet carries the golden F1.3 rate");
}

function testGoldenV2RequestFilters() {
  ["53", "55"].forEach((provinceCode) => {
    const sandbox = createSandbox(async () => { throw new Error("not used"); }, { code: provinceCode, name: "Test" });
    const paramsByBuilder = JSON.parse(vm.runInContext(`JSON.stringify([
      buildPhatThanhCongV2Params,
      buildNoiTinhF11Params,
      buildThuGomLienTinhF12Params,
      buildPhatLienTinhF13Params
    ].map(builder => ({ tinh: builder("TINH", new Date(2026, 0, 1), new Date(2026, 0, 31)), bc: builder("BC", new Date(2026, 0, 1), new Date(2026, 0, 31)) })))`, sandbox));
    paramsByBuilder.forEach(({ tinh, bc }, index) => {
      assert.equal(tinh.iFrom, "01/01/2026");
      assert.equal(tinh.iTo, "01/31/2026");
      assert.equal(tinh.iPageSize, "50000");
      assert.equal(bc.iPageSize, "50000");
      const provinceKey = index === 1 ? "stMaTinhChapNhan" : index === 2 ? "stMaTinhNhan" : "stMaTinhPhat";
      assert.equal(tinh[provinceKey], "ALL", "TINH keeps the golden nationwide filter");
      assert.equal(bc[provinceKey], provinceCode, "BC filter follows the dropdown instead of a hardcoded province");
    });
    assert.equal(paramsByBuilder[0].bc.stMaLoaiBCPhat, "NULL");
    assert.equal(paramsByBuilder[0].bc.stMaBuuCucPhat, "ALL");
  });
}

function testGoldenBcValueMappings() {
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, { code: 55, name: "TP. Đà Nẵng" });
  const configs = vm.runInContext("V2_SOURCE_CONFIGS", sandbox);
  const expected = {
    phatThanhCongBuuCuc: [200000, null, 190000],
    noiTinhF11: [120000, 100000, 95000],
    thuGomLienTinhF12: [1358542, 1350000, 1277357],
    phatLienTinhF13: [102596, null, 62848]
  };
  configs.forEach((config) => {
    const cells = Array(30).fill("");
    cells[config.bcCodeIndex] = "9001";
    cells[config.bcNameIndex] = "BCVH Golden";
    const [total, subTotal, onTime] = expected[config.key];
    cells[config.bcValues.total] = String(total);
    if (Number.isInteger(config.bcValues.subTotal)) cells[config.bcValues.subTotal] = String(subTotal);
    cells[config.bcValues.onTime] = String(onTime);
    const records = sandbox.parseV2ApiRows(`<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`, config, "BC", new Set(["9001"]));
    assert.equal(records.length, 1, `${config.key}: BC row is selected with its golden code column`);
    assert.deepEqual(
      [Number(records[0].total), Number(records[0].subTotal || 0), Number(records[0].onTime)],
      [total, subTotal || 0, onTime],
      `${config.key}: BC uses golden bcValues rather than TINH values`
    );
  });
}

async function testMultiMonthMissingKpiIsNaInsteadOfComparisonFailure() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const sandbox = createSandbox(async (requestUrl) => {
    const pathName = new URL(requestUrl).pathname;
    const missingF11 = pathName.includes("chat-luong-toan-trinh-buu-gui-noi-tinh");
    return {
      ok: true,
      text: async () => JSON.stringify({
        data: makeSourceHtml(missingF11 ? [10] : [selectedProvince.code, 10], {
          [selectedProvince.code]: selectedProvince.name,
          10: "TP. Hà Nội"
        })
      })
    };
  }, selectedProvince);
  vm.runInContext(`getSelectedMetricLabels = () => ${JSON.stringify([
    "F1.1 – Nội tỉnh",
    "F1.2 – Thu gom bưu gửi đi liên tỉnh",
    "F1.3 – Chất lượng phát bưu gửi liên tỉnh",
    "F4.1 – Chất lượng phát thành công tại bưu cục"
  ])};`, sandbox);
  const months = sandbox.generateMonthRanges("2026-01", "2026-02");
  const results = await sandbox.fetchMultiMonthV2Report(months, "TINH", null, null);
  const focusRows = results.map(({ rows }) => rows.find((row) => String(row[1]) === String(selectedProvince.code)));
  focusRows.forEach((row, monthIndex) => {
    assert.equal(row[22], "", `missing F1.1 month ${monthIndex + 1} keeps its denominator blank`);
    assert.equal(row[24], "", `missing F1.1 month ${monthIndex + 1} keeps its rate N/A`);
    assert.notEqual(row[26], "", `other KPI data remains available in month ${monthIndex + 1}`);
  });
  const dashboard = sandbox.renderV2MultiMonthDashboardWorksheet(results, months, "TINH");
  assert.match(getRowContaining(dashboard, "F1.1 – Chất lượng phát bưu gửi nội tỉnh"), />N\/A<\/Data>/, "multi-month Dashboard renders missing F1.1 as N/A");
}

function testMonthlyDashboardUsesAdjacentMonthsAndMonthlyRanks() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const sandbox = createSandbox(async () => { throw new Error("not used"); }, selectedProvince);
  vm.runInContext(`getSelectedMetricLabels = () => ${JSON.stringify([
    "F1.1 – Nội tỉnh",
    "F1.2 – Thu gom bưu gửi đi liên tỉnh",
    "F1.3 – Chất lượng phát bưu gửi liên tỉnh",
    "F4.1 – Chất lượng phát thành công tại bưu cục"
  ])};`, sandbox);
  const provinceNames = JSON.parse(vm.runInContext("JSON.stringify(PROVINCE_NAMES)", sandbox));
  const metrics = JSON.parse(vm.runInContext("JSON.stringify(getActiveMultiMonthMetrics())", sandbox));
  const months = sandbox.generateMonthRanges("2026-01", "2026-03");
  const focusSeries = {
    "F1.1": [{ volume: 10000, rate: 0.90 }, { volume: 20000, rate: 0.95 }, { volume: 30000, rate: 0.95 }],
    "F1.2": [{ volume: 11000, rate: 0.91 }, { volume: 22000, rate: 0.92 }, { volume: 33000, rate: 0.93 }],
    "F1.3": [{ volume: 12000, rate: 0.88 }, { volume: 24000, rate: 0.90 }, { volume: 36000, rate: 0.92 }],
    "F4.1": [{ volume: 13000, rate: 0.89 }, { volume: 26000, rate: 0.90 }, { volume: 39000, rate: 0.91 }]
  };

  function makeRow(code, name, monthIndex, mutateF11Middle = false) {
    const row = Array(37).fill("");
    row[1] = String(code);
    row[2] = name;
    metrics.forEach((metric) => {
      const isFocus = String(code) === String(selectedProvince.code);
      if (isFocus && metric.shortName === "F1.2" && monthIndex === 1) return;
      const point = isFocus
        ? { ...focusSeries[metric.shortName][monthIndex] }
        : { volume: 10000 + monthIndex, rate: metric.shortName === "F1.1" ? [0.89, 0.96, 0.94][monthIndex] : 0.80 };
      if (mutateF11Middle && isFocus && metric.shortName === "F1.1" && monthIndex === 1) point.rate = 0.80;
      row[metric.totalCol] = String(point.volume);
      row[metric.onTimeCol] = String(Math.round(point.volume * point.rate));
    });
    return row;
  }

  function makeMonthlyResults(mutateF11Middle = false) {
    return months.map((monthObj, monthIndex) => ({
      monthObj,
      rows: Object.entries(provinceNames).map(([code, name]) => makeRow(code, name, monthIndex, mutateF11Middle))
    }));
  }

  function buildNative(monthlyResults) {
    return sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
      monthlyResults,
      monthsList: months,
      tuyChonGR: "TINH",
      selectedProvinceCode: selectedProvince.code,
      selectedProvinceName: selectedProvince.name,
      metrics,
      columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
      classifyKpi: sandbox.getKpiStatus,
      toNumberValue: sandbox.toNumberValue
    });
  }

  const monthlyResults = makeMonthlyResults();
  const native = buildNative(monthlyResults);
  const legacy = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, months, "TINH");
  const f11 = native.model.metrics.find((metric) => metric.shortName === "F1.1");
  assert.deepEqual(Array.from(f11.points, point => point.volume), [10000, 20000, 30000], "F1.1 keeps a separate volume for every month");
  assert.deepEqual(Array.from(f11.points, point => point.rate), [0.90, 0.95, 0.95], "F1.1 keeps a separate rate for every month");
  assert.deepEqual(Array.from(f11.points, point => point.rank), [1, 34, 1], "F1.1 ranks against all 34 valid provinces for each individual month");
  assert.deepEqual(Array.from(f11.points, point => point.rankCount), [34, 34, 34], "monthly rank denominator is the current month's valid province set");
  assert.deepEqual(Array.from(f11.points, point => point.rateDelta == null ? null : Number(point.rateDelta.toFixed(6))), [null, 0.05, 0], "F1.1 compares each month only with its immediate predecessor");
  assert.deepEqual(Array.from(f11.points, point => point.trend), [null, 1, 0], "F1.1 trend is up then unchanged");

  const f12 = native.model.metrics.find((metric) => metric.shortName === "F1.2");
  assert.deepEqual(Array.from(f12.points, point => point.volume), [11000, null, 33000], "a missing monthly F1.2 record remains N/A rather than zero");
  assert.deepEqual(Array.from(f12.points, point => point.rank), [1, null, 1], "a province without valid data is excluded from that month's ranking");
  assert.deepEqual(Array.from(f12.points, point => point.rateDelta), [null, null, null], "N/A never creates a fabricated adjacent-month delta");

  ["F1.1", "F1.2", "F1.3", "F4.1"].forEach((shortName) => {
    const metric = native.model.metrics.find((item) => item.shortName === shortName);
    assert.ok(metric.points.every((point) => point.rank == null || point.rankCount === 34), `${shortName} uses its own monthly ranking population`);
  });
  const legacyHeader = getRowContaining(legacy, "01/2026 Sản lượng");
  assert.ok(legacyHeader && legacyHeader.includes("03/2026 Xu hướng"), "legacy Dashboard has five monthly fields for every selected month");
  assert.doesNotMatch(legacyHeader, /Mục tiêu|Kế hoạch|BQ|lũy kế|Chênh mục tiêu|Trạng thái/i, "legacy monthly table has no target, cumulative-average or plan comparison columns");
  const legacyF11 = getRowContaining(legacy, "F1.1 – Chất lượng phát bưu gửi nội tỉnh");
  assert.match(legacyF11, />10000<\/Data>/, "legacy Dashboard reads the first month volume directly");
  assert.match(legacyF11, />20000<\/Data>/, "legacy Dashboard reads the second month volume directly");
  assert.match(legacyF11, />30000<\/Data>/, "legacy Dashboard reads the third month volume directly");
  assert.match(legacyF11, />0\.04[0-9]*<\/Data>/, "legacy Dashboard renders the adjacent-month rate delta");
  assert.ok(legacyF11.includes("▲") && legacyF11.includes("→"), "legacy Dashboard renders up and unchanged indicators");

  const nativeFiles = sandbox.fflate.unzipSync(native.bytes);
  const nativeDashboard = Buffer.from(nativeFiles["xl/worksheets/sheet1.xml"]).toString("utf8");
  assert.ok(nativeDashboard.includes("01/2026 Sản lượng") && nativeDashboard.includes("03/2026 Xu hướng"), "native Dashboard carries the per-month table");
  assert.doesNotMatch(nativeDashboard, /Tỷ lệ lũy kế|Chênh mục tiêu|>Trạng thái<|>Mục tiêu</, "native Dashboard removes cumulative and target comparison columns");
  assert.ok(nativeDashboard.includes(">N/A<"), "native Dashboard displays missing monthly data as N/A");

  const mutated = buildNative(makeMonthlyResults(true));
  const mutatedF11 = mutated.model.metrics.find((metric) => metric.shortName === "F1.1");
  assert.equal(mutatedF11.points[0].rate, f11.points[0].rate, "changing February leaves January untouched");
  assert.equal(mutatedF11.points[1].rate, 0.80, "changing February changes February only");
  assert.equal(mutatedF11.points[2].rate, f11.points[2].rate, "changing February leaves March source data untouched");
  assert.deepEqual(Array.from(mutatedF11.points, point => point.rateDelta == null ? null : Number(point.rateDelta.toFixed(6))), [null, -0.10, 0.15], "only February and the March-vs-February comparison change");
  assert.deepEqual(Array.from(mutatedF11.points, point => point.trend), [null, -1, 1], "mutated adjacent trends are down then up");
  assert.deepEqual(
    Array.from(mutated.model.metrics.find((metric) => metric.shortName === "F1.3").points, point => point.rate),
    Array.from(native.model.metrics.find((metric) => metric.shortName === "F1.3").points, point => point.rate),
    "changing one F1.1 month does not alter another KPI"
  );
}


// PO report: ticking F1.1 + F1.2 + F1.3 (F4.1 left off) produced a workbook that carried
// essentially only F1.3. This drives the WHOLE multi-month pipeline for that exact
// combination -- selection -> fetch -> monthlyResults -> legacy .xls dashboard, native
// .xlsx model, matrix sheet and DuLieu_Chart -- and requires real data for all three.
async function testSelectedKpiComboSurvivesWholeMultiMonthPipeline() {
  const selectedProvince = { code: 53, name: "Thừa Thiên Huế" };
  const chosenLabels = [
    "F1.1 – Nội tỉnh",
    "F1.2 – Thu gom bưu gửi đi liên tỉnh",
    "F1.3 – Chất lượng phát bưu gửi liên tỉnh"
  ];
  const chosenShortNames = ["F1.1", "F1.2", "F1.3"];
  const omitted = "F4.1";

  const provinceNames = JSON.parse(vm.runInContext(
    "JSON.stringify(PROVINCE_NAMES)",
    createSandbox(async () => { throw new Error("not used"); }, selectedProvince)
  ));
  const codes = Object.keys(provinceNames).map(Number);

  const sandbox = createSandbox(async () => ({
    ok: true,
    text: async () => JSON.stringify({ data: makeSourceHtml(codes, provinceNames) })
  }), selectedProvince);
  vm.runInContext(
    `document.querySelectorAll = (selector) => selector.includes(":checked")
       ? ${JSON.stringify(chosenLabels)}.map((value) => ({ value, checked: true }))
       : [];`,
    sandbox
  );

  const activeMetrics = JSON.parse(vm.runInContext("JSON.stringify(getActiveMultiMonthMetrics('TINH'))", sandbox));
  assert.deepEqual(
    activeMetrics.map((metric) => metric.shortName),
    chosenShortNames,
    "the three ticked KPIs, and only those, reach the exporter"
  );

  const monthsList = sandbox.generateMonthRanges("2026-01", "2026-03");
  const monthlyResults = await sandbox.fetchMultiMonthV2Report(monthsList, "TINH", null, null);
  assert.equal(monthlyResults.length, 3, "three months are fetched");

  // 1. monthlyResults: every ticked KPI carries real volume and on-time numbers.
  monthlyResults.forEach(({ rows }, monthIndex) => {
    const focusRow = rows.find((row) => String(row[1]) === String(selectedProvince.code));
    assert.ok(focusRow, `month ${monthIndex + 1}: the focus province is present`);
    activeMetrics.forEach((metric) => {
      assert.ok(
        sandbox.toNumberValue(focusRow[metric.totalCol]) > 0,
        `month ${monthIndex + 1}: ${metric.shortName} must carry a volume, not an empty cell`
      );
      assert.ok(
        sandbox.toNumberValue(focusRow[metric.onTimeCol]) > 0,
        `month ${monthIndex + 1}: ${metric.shortName} must carry an on-time count`
      );
    });
  });

  // 2. Legacy .xls dashboard: all three KPIs, and no trace of the unticked one.
  const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, monthsList, "TINH");
  chosenShortNames.forEach((shortName) => {
    assert.ok(legacyXml.includes(shortName), `legacy dashboard renders ${shortName}`);
  });
  assert.ok(!legacyXml.includes(omitted), `legacy dashboard must not mention the unticked ${omitted}`);
  assertWellFormedXml(legacyXml);

  // 3. Native .xlsx model: three metrics, each with a non-zero total and a real rate.
  const native = sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList,
    tuyChonGR: "TINH",
    selectedProvinceCode: selectedProvince.code,
    selectedProvinceName: selectedProvince.name,
    metrics: activeMetrics,
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus,
    toNumberValue: sandbox.toNumberValue
  });
  assert.deepEqual(
    Array.from(native.model.metrics, (metric) => metric.shortName),
    chosenShortNames,
    "native model exports exactly the ticked KPIs"
  );
  native.model.metrics.forEach((metric) => {
    assert.ok(metric.total > 0, `${metric.shortName}: native total must not be 0`);
    assert.ok(metric.cumulativeRate > 0, `${metric.shortName}: native rate must not be 0`);
    assert.equal(metric.points.length, monthsList.length, `${metric.shortName}: one point per month`);
    metric.points.forEach((point, monthIndex) => {
      assert.equal(point.hasData, true, `${metric.shortName} month ${monthIndex + 1}: has data`);
      assert.ok(point.volume > 0, `${metric.shortName} month ${monthIndex + 1}: volume`);
    });
  });

  // 4. Every sheet of the package: three KPIs present, unticked one absent everywhere.
  const files = sandbox.fflate.unzipSync(native.bytes);
  const partOf = (name) => Buffer.from(files[name]).toString("utf8");
  const chartData = partOf("xl/worksheets/sheet2.xml");
  const matrix = partOf("xl/worksheets/sheet3.xml");
  chosenShortNames.forEach((shortName) => {
    assert.ok(chartData.includes(shortName), `DuLieu_Chart carries the ${shortName} block`);
    assert.ok(matrix.includes(shortName), `matrix sheet carries the ${shortName} column`);
  });
  Object.entries(files)
    .filter(([name]) => name.endsWith(".xml"))
    .forEach(([name, bytes]) => {
      assert.ok(
        !Buffer.from(bytes).toString("utf8").includes(omitted),
        `${name} must not mention the unticked ${omitted}`
      );
    });

  // 5. One combo chart per ticked KPI, and the package still opens without repair.
  const chartParts = Object.keys(files).filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
  assert.equal(chartParts.length, chosenShortNames.length, "one chart per ticked KPI");
  chartParts.forEach((name) => {
    const xml = partOf(name);
    assertWellFormedXml(xml);
    assertNoDuplicateXmlAttributes(xml, name);
  });
  Object.entries(files)
    .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
    .forEach(([name, bytes]) => {
      const xml = Buffer.from(bytes).toString("utf8");
      assertWellFormedXml(xml);
      assertNoDuplicateXmlAttributes(xml, name);
    });
}

(async () => {
  for (const selectedProvince of [
    { code: 53, name: "Thừa Thiên Huế" },
    { code: 10, name: "TP. Hà Nội" }
  ]) {
    await testProvinceJoin(selectedProvince);
    await testZeroValuedSelectedProvince(selectedProvince);
    testDynamicDashboardFocus(selectedProvince);
  }
  await testGoldenF13MutationFails();
  await testGoldenMappingFlowsThroughPipeline();
  await testF12ProductionRecordAcceptsSubTotalAboveTotal();
  await testGoldenMappingSupportsTwoPeriodComparison();
  testGoldenV2RequestFilters();
  testGoldenBcValueMappings();
  await testMultiMonthMissingKpiIsNaInsteadOfComparisonFailure();
  testMonthlyDashboardUsesAdjacentMonthsAndMonthlyRanks();
  testValidationHasNoHardcodedFocusProvince();
  testComparisonDashboardXml();
  testKpiSsotColorChartAndCommentary();
  testNativeXlsxPackageAndCharts();
  testNativeXlsxExportsOnlySelectedKpis();
  await testSelectedKpiComboSurvivesWholeMultiMonthPipeline();
  for (const selectedProvince of [
    { code: 53, name: "Thừa Thiên Huế" },
    { code: 10, name: "TP. Hà Nội" }
  ]) {
    testNativeXlsxMatchesLegacyXlsNumbers(selectedProvince);
  }
  console.log("V2 report regressions passed");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
