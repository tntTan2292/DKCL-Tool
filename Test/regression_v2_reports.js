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
    const cells = Array(26).fill("");
    cells[1] = String(code);
    cells[2] = provinceNames[code] || "Ngoài danh mục";
    cells[5] = zeroValues ? "0" : "100";
    cells[6] = zeroValues ? "0" : "90";
    cells[7] = zeroValues ? "0" : "80";
    cells[8] = zeroValues ? "0" : "70";
    cells[10] = zeroValues ? "0" : "120";
    cells[11] = zeroValues ? "0" : "110";
    cells[12] = zeroValues ? "0" : "100";
    cells[16] = zeroValues ? "0" : "95";
    cells[25] = zeroValues ? "0" : "115";
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
  assert.ok(dashboardXml.includes("(SL: 0 bg)"), "zero-valued dashboard displays zero volume");
  assert.match(getRowContaining(dashboardXml, "F4.1 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiRed"/, "valid zero rate is retained and classified against its KPI target");
}

function makeMonthlyRow(code, name, seed) {
  const row = Array(37).fill("");
  row[1] = String(code);
  row[2] = name;
  [17, 21, 26, 31].forEach((index) => { row[index] = String(100 + seed); });
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
  assert.equal((f11TargetRow.match(/ss:Type="Number">0\.95<\/Data>/g) || []).length, months.length + 1, "F1.1 target line stays at 95% for every month and summary");
  assert.equal((f13TargetRow.match(/ss:Type="Number">0\.9<\/Data>/g) || []).length, months.length + 1, "F1.3 target line stays at 90% for every month and summary");
  assert.match(getRowContaining(dashboard, "F1.1 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiGreen"/);
  assert.match(getRowContaining(dashboard, "F1.2 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiYellow"/);
  assert.match(getRowContaining(dashboard, "F1.3 – Cột tỷ lệ thực tế"), /ss:StyleID="KpiRed"/);
  assert.ok(dashboard.includes("Ngưỡng SSOT F1.1/F1.2 = 95%, F1.3/F4.1 = 90%"));
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

(async () => {
  for (const selectedProvince of [
    { code: 53, name: "Thừa Thiên Huế" },
    { code: 10, name: "TP. Hà Nội" }
  ]) {
    await testProvinceJoin(selectedProvince);
    await testZeroValuedSelectedProvince(selectedProvince);
    testDynamicDashboardFocus(selectedProvince);
  }
  testValidationHasNoHardcodedFocusProvince();
  testComparisonDashboardXml();
  testKpiSsotColorChartAndCommentary();
  testNativeXlsxPackageAndCharts();
  console.log("V2 report regressions passed");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
