// Differential harness: "V2 -> So sánh 2 kỳ -> Bưu cục" is a legacy feature that worked in
// the original extension. It must keep producing the same business values after the
// multi-month feature was added. Both extensions are driven by ONE fixture and compared on
// business values, not on whether the .xls opens.
//
// GOLDEN BASELINE: D:\Antigravity - Project\SOSANH-Cu\SOSANH\BDHN_DKCL
// CURRENT:         this repository
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const GOLDEN_DIR = process.env.DKCL_GOLDEN_DIR
  || "D:/Antigravity - Project/SOSANH-Cu/SOSANH/BDHN_DKCL";

const currentPopupSource = fs.readFileSync(path.join(__dirname, "..", "popup.js"), "utf8");
const currentExporterSource = fs.readFileSync(path.join(__dirname, "..", "xlsx-export.js"), "utf8");
const fflateSource = fs.readFileSync(path.join(__dirname, "..", "vendor", "fflate.min.js"), "utf8");
const goldenPopupSource = fs.readFileSync(path.join(GOLDEN_DIR, "popup.js"), "utf8");

const SELECTED_PROVINCE = { code: 53, name: "Thừa Thiên Huế" };

const V1_PATH = "/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc";
const F41_PATH = "/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc";
const F11_PATH = "/kpi/chat-luong-toan-trinh-buu-gui-noi-tinh";
const F12_PATH = "/kpi/chat-luong-thu-gom-buu-lien-tinh";
const F13_PATH = "/kpi/chat-luong-phat-buu-gui-lien-tinh";

// ---------------------------------------------------------------------------
// Sandboxes
// ---------------------------------------------------------------------------

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
    textContent: "",
    innerHTML: "",
    dataset: {},
    classList: { add() {}, remove() {}, toggle() {} },
    selectedOptions: [],
    addEventListener() {},
    setAttribute() {},
    getAttribute() { return "false"; },
    querySelector() { return null; },
    scrollIntoView() {},
    ...overrides
  };
}

function createSandbox(source, { fetchImpl, withExporter = false }) {
  const elements = new Map();
  const document = {
    getElementById(id) {
      if (!elements.has(id)) {
        const overrides = id === "province-select" ? {
          value: String(SELECTED_PROVINCE.code),
          selectedOptions: [{ textContent: `${SELECTED_PROVINCE.code} - ${SELECTED_PROVINCE.name}` }]
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
    localStorage: { getItem() { return null; }, setItem() {} },
    DOMParser: TestDomParser,
    console: { log() {}, error() {}, warn() {} },
    fetch: fetchImpl,
    Blob: class { constructor(parts) { this.parts = parts; } },
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
    isNaN,
    setTimeout
  };
  sandbox.window.parent = sandbox.window;
  vm.createContext(sandbox);
  vm.runInContext(fflateSource, sandbox);
  if (withExporter) vm.runInContext(currentExporterSource, sandbox);
  vm.runInContext(source, sandbox);
  return sandbox;
}

const createGolden = (fetchImpl) => createSandbox(goldenPopupSource, { fetchImpl });
const createCurrent = (fetchImpl) => createSandbox(currentPopupSource, { fetchImpl, withExporter: true });

// ---------------------------------------------------------------------------
// Fixture: one post-office dataset, served identically to both extensions
// ---------------------------------------------------------------------------
//
// Cell layout per endpoint, taken from the golden config (bcCodeIndex/bcNameIndex/bcValues):
//   F4.1 + V1 : code@5  name@6  total@10 onTime@27  (V1 success8h also mirrored at @25)
//   F1.1      : code@3  name@4  total@13 subTotal@20 onTime@21
//   F1.2      : code@3  name@4  total@9  subTotal@10 onTime@11
//   F1.3      : code@3  name@4  total@10 onTime@18
//
// The V1 `success8h` offset differs between the extensions (golden 27, current 25). That is
// a separate, pre-existing divergence outside the BC-compare scope, so the fixture writes
// the same number into both cells and this harness stays focused on the BC regressions.

const BC_UNITS = {
  // code: [name in period 1, name in period 2]
  101: ["BCVH Huế 1", "BCVH Huế 1"],
  102: ["BCVH Huế 2", "BCVH Huế 2"],
  103: ["BCVH Huế 3", null],                                   // only period 1
  104: [null, "BCVH Huế 4"],                                   // only period 2
  // same code, name differs by whitespace + Unicode form -> join must use the code
  105: ["BCVH Phú  Vang", "BCVH Phú Vang\u00a0"]
};

const BC_VOLUMES = {
  // code: { period: { kpi: [total, onTime] } }
  101: { 1: { F41: [1000, 900], F11: [800, 760], F12: [600, 570], F13: [500, 440] },
         2: { F41: [1100, 1045], F11: [820, 770], F12: [640, 608], F13: [520, 452] } },
  102: { 1: { F41: [2000, 1700], F11: [1500, 1380], F12: [1200, 1080], F13: [900, 738] },
         2: { F41: [2100, 1869], F11: [1550, 1410], F12: [1250, 1150], F13: [950, 798] } },
  103: { 1: { F41: [300, 240], F11: [250, 225], F12: [200, 180], F13: [150, 120] } },
  104: { 2: { F41: [400, 360], F11: [350, 322], F12: [300, 279], F13: [250, 210] } },
  105: { 1: { F41: [5000, 4250], F11: [4000, 3760], F12: [3000, 2760], F13: [2500, 2000] },
         2: { F41: [5200, 4576], F11: [4100, 3813], F12: [3100, 2883], F13: [2600, 2158] } }
};

function unitsInPeriod(period) {
  return Object.keys(BC_UNITS)
    .map(Number)
    .filter((code) => BC_UNITS[code][period - 1] && BC_VOLUMES[code][period])
    .sort((a, b) => a - b);
}

function bcRow(cells) {
  return `<tr>${Array.from({ length: 30 }, (_, i) => `<td>${cells[i] ?? ""}</td>`).join("")}</tr>`;
}

function buildFixtureHtml(sourcePath, period) {
  return unitsInPeriod(period).map((code) => {
    const name = BC_UNITS[code][period - 1];
    const volumes = BC_VOLUMES[code][period];
    const cells = [];
    cells[0] = "1";
    cells[1] = String(SELECTED_PROVINCE.code);
    cells[2] = SELECTED_PROVINCE.name;

    if (sourcePath === F41_PATH) {
      cells[5] = String(code);
      cells[6] = name;
      cells[10] = String(volumes.F41[0]);
      cells[25] = String(volumes.F41[1]);
      cells[27] = String(volumes.F41[1]);
    } else {
      cells[3] = String(code);
      cells[4] = name;
      if (sourcePath === F11_PATH) {
        cells[13] = String(volumes.F11[0]);
        cells[20] = String(volumes.F11[1]);
        cells[21] = String(volumes.F11[1]);
      } else if (sourcePath === F12_PATH) {
        cells[9] = String(volumes.F12[0]);
        cells[10] = String(volumes.F12[1]);
        cells[11] = String(volumes.F12[1]);
      } else if (sourcePath === F13_PATH) {
        cells[10] = String(volumes.F13[0]);
        cells[18] = String(volumes.F13[1]);
      }
    }
    return bcRow(cells);
  }).join("");
}

function makeFetch(period, { record } = {}) {
  return async (requestUrl) => {
    const url = new URL(requestUrl);
    if (record) record.push({ path: url.pathname, params: Object.fromEntries(url.searchParams) });
    return { ok: true, text: async () => JSON.stringify({ data: buildFixtureHtml(url.pathname, period) }) };
  };
}

// ---------------------------------------------------------------------------
// Business-value extraction
// ---------------------------------------------------------------------------

async function runBcCompare(sandbox, { record1, record2 } = {}) {
  const from = new Date(2026, 7, 25);
  const to = new Date(2026, 7, 25);
  const compareFrom = new Date(2026, 7, 24);
  const compareTo = new Date(2026, 7, 24);

  sandbox.fetch = makeFetch(1, { record: record1 });
  const currentRows = await sandbox.fetchV2ReportRows("BC", from, to, {}, "kỳ báo cáo");
  sandbox.fetch = makeFetch(2, { record: record2 });
  const compareRows = await sandbox.fetchV2ReportRows("BC", compareFrom, compareTo, {}, "kỳ so sánh");

  // The golden signature takes tuyChonGR; the current one (before remediation) does not.
  const comparisonRows = sandbox.buildV2ComparisonRows(currentRows, compareRows, "BC");
  return { currentRows, compareRows, comparisonRows, from, to, compareFrom, compareTo };
}

function summarise(run) {
  // Values crossing back from the vm realm keep that realm's Array prototype, which
  // deepStrictEqual rejects. Round-trip everything into this realm first.
  return JSON.parse(JSON.stringify(summariseRaw(run)));
}

function summariseRaw({ currentRows, compareRows, comparisonRows }) {
  const codesOf = (rows) => Array.from(rows, (row) => String(row[1]));
  return {
    period1Codes: codesOf(currentRows),
    period2Codes: codesOf(compareRows),
    period1Names: currentRows.map((row) => row[2]),
    comparisonCodes: comparisonRows.map((row) => String(row.code)),
    comparisonNames: comparisonRows.map((row) => row.name),
    totalUnits: comparisonRows.map((row) => row.totalUnits),
    dataRowCount: comparisonRows.reduce((sum, row) => sum + row.metrics.length, 0),
    metrics: comparisonRows.map((row) => ({
      code: String(row.code),
      labels: row.metrics.map((metric) => metric.label),
      currentRates: row.metrics.map((metric) => metric.currentRate),
      compareRates: row.metrics.map((metric) => metric.compareRate),
      rateDiffPoints: row.metrics.map((metric) => metric.rateDiffPoints)
    }))
  };
}

function bcWorksheetOf(sandbox, run) {
  const xml = sandbox.downloadV2ComparisonExcel ? captureWorkbookXml(sandbox, run) : "";
  return xml;
}

function captureWorkbookXml(sandbox, run) {
  let captured = "";
  const originalSave = sandbox.saveBlobAsFile;
  sandbox.saveBlobAsFile = (blob) => { captured = String(blob.parts[0]); };
  sandbox.downloadV2ComparisonExcel({
    currentRows: run.currentRows,
    compareRows: run.compareRows,
    comparisonRows: run.comparisonRows,
    tuyChonGR: "BC",
    from: run.from,
    to: run.to,
    compareFrom: run.compareFrom,
    compareTo: run.compareTo,
    fileName: "differential.xls"
  });
  sandbox.saveBlobAsFile = originalSave;
  return captured;
}

function worksheetNames(xml) {
  return [...xml.matchAll(/<Worksheet ss:Name="([^"]+)"/g)].map((match) => match[1]);
}

function comparisonSheetHeader(xml, sheetName) {
  const sheet = xml.match(new RegExp(`<Worksheet ss:Name="${sheetName}">[\\s\\S]*?</Worksheet>`));
  if (!sheet) return null;
  const rows = [...sheet[0].matchAll(/<Row>([\s\S]*?)<\/Row>/g)];
  const headerRow = rows.find((row) => row[1].includes('ss:StyleID="Header"'));
  if (!headerRow) return null;
  return [...headerRow[1].matchAll(/<Data ss:Type="String">([\s\S]*?)<\/Data>/g)].map((match) => match[1]);
}

// ---------------------------------------------------------------------------
// Differential assertions
// ---------------------------------------------------------------------------

function diffReport(label, golden, current) {
  const lines = [];
  const keys = new Set([...Object.keys(golden), ...Object.keys(current)]);
  for (const key of keys) {
    const a = JSON.stringify(golden[key]);
    const b = JSON.stringify(current[key]);
    if (a !== b) lines.push(`  ${key}\n    golden : ${a}\n    current: ${b}`);
  }
  return lines.length ? `${label}\n${lines.join("\n")}` : "";
}

async function testBcCompareMatchesGolden() {
  const goldenRequests = [];
  const currentRequests = [];

  const golden = createGolden(makeFetch(1));
  const current = createCurrent(makeFetch(1));

  const goldenRun = await runBcCompare(golden, { record1: goldenRequests, record2: [] });
  const currentRun = await runBcCompare(current, { record1: currentRequests, record2: [] });

  const goldenSummary = summarise(goldenRun);
  const currentSummary = summarise(currentRun);

  // --- fixture sanity: the dataset really does exercise the join edge cases ---
  assert.ok(goldenSummary.period1Codes.includes("103"), "fixture: unit 103 exists only in period 1");
  assert.ok(!goldenSummary.period1Codes.includes("104"), "fixture: unit 104 is absent from period 1");
  assert.ok(goldenSummary.comparisonCodes.includes("105"), "fixture: unit 105 joins across periods despite a different name spelling");

  const report = diffReport("BC compare 2 kỳ khác Extension gốc:", goldenSummary, currentSummary);
  assert.equal(report, "", report || "BC compare matches the golden extension");

  // Absolute oracle derived from the fixture, not from golden. Parity alone cannot catch a
  // fault that hits both extensions the same way (e.g. period 1 data served for period 2),
  // so the expected numbers are stated outright.
  const expectedPerUnit = {
    101: { current: [0.9, 0.9, 0.9, 440 / 500], compare: [0.95, 0.95, 0.95, 452 / 520] },
    102: { current: [0.85, 0.85, 0.85, 738 / 900], compare: [0.89, 0.89, 0.89, 798 / 950] },
    105: { current: [0.85, 0.85, 0.85, 2000 / 2500], compare: [4576 / 5200, 4576 / 5200, 4576 / 5200, 2158 / 2600] }
  };
  assert.deepEqual(currentSummary.comparisonCodes, ["101", "102", "105"], "only units present in both periods are compared");
  currentSummary.metrics.forEach((entry) => {
    const expected = expectedPerUnit[entry.code];
    assert.ok(expected, `unit ${entry.code} has an expected result`);
    entry.currentRates.forEach((rate, index) => {
      assert.ok(Math.abs(rate - expected.current[index]) < 1e-9, `unit ${entry.code} metric #${index + 1}: period 1 rate is ${rate}, expected ${expected.current[index]}`);
    });
    entry.compareRates.forEach((rate, index) => {
      assert.ok(Math.abs(rate - expected.compare[index]) < 1e-9, `unit ${entry.code} metric #${index + 1}: period 2 rate is ${rate}, expected ${expected.compare[index]}`);
    });
    entry.currentRates.forEach((rate, index) => {
      assert.notEqual(rate, entry.compareRates[index], `unit ${entry.code} metric #${index + 1}: the two periods must not carry identical data`);
    });
  });

  // Units present in only one period must survive in that period's own data sheet.
  assert.deepEqual(currentSummary.period1Codes, ["101", "102", "103", "105"], "period 1 keeps the unit that exists only in period 1");
  assert.deepEqual(currentSummary.period2Codes, ["101", "102", "104", "105"], "period 2 keeps the unit that exists only in period 2");

  // --- worksheet structure ---
  const goldenXml = captureWorkbookXml(golden, goldenRun);
  const currentXml = captureWorkbookXml(current, currentRun);

  assert.ok(worksheetNames(goldenXml).includes("SoSanhV2_BC"), "golden emits the BC comparison sheet");
  assert.deepEqual(
    worksheetNames(currentXml).filter((name) => name.startsWith("SoSanhV2")),
    worksheetNames(goldenXml).filter((name) => name.startsWith("SoSanhV2")),
    "current must emit the same V2 comparison sheet(s) as the golden extension"
  );
  assert.deepEqual(
    comparisonSheetHeader(currentXml, "SoSanhV2_BC"),
    comparisonSheetHeader(goldenXml, "SoSanhV2_BC"),
    "BC comparison sheet must keep the golden column titles and order"
  );

  // --- request shape: same endpoints, same filters ---
  const goldenPaths = [...new Set(goldenRequests.map((entry) => entry.path))].sort();
  const currentPaths = [...new Set(currentRequests.map((entry) => entry.path))].sort();
  assert.deepEqual(currentPaths, goldenPaths, "BC must call exactly the endpoints the golden extension calls");

  // Two sanctioned divergences are normalised away:
  //  * the province filter: golden hardcodes its home province, current reads the dropdown.
  //    Both must resolve to the province actually selected.
  //  * the legacy V1 service filters, which have differed between the two extensions since
  //    commit 420854a — long before the multi-month feature — so they are not a regression
  //    of this change. They are listed explicitly so a NEW divergence still fails.
  const PROVINCE_FILTER_KEYS = ["stMaTinhPhat", "stMaTinhThuGom", "stMaTinhChapNhan", "stMaTinhNhan"];
  const PRE_EXISTING_V1_KEYS = ["stLoaiDichVu", "stNhomLoaiBuuGui[]", "stMaLoaiBCPhat"];

  goldenRequests.forEach((goldenEntry, index) => {
    const currentEntry = currentRequests[index];
    assert.ok(currentEntry, `current issues request #${index + 1}`);

    PROVINCE_FILTER_KEYS.forEach((key) => {
      if (!(key in goldenEntry.params) && !(key in currentEntry.params)) return;
      assert.equal(
        currentEntry.params[key],
        String(SELECTED_PROVINCE.code),
        `request #${index + 1}: ${key} must come from the province dropdown, not a hardcoded province`
      );
    });

    const strip = (params) => Object.fromEntries(
      Object.entries(params).filter(([key]) => !PROVINCE_FILTER_KEYS.includes(key) && !PRE_EXISTING_V1_KEYS.includes(key))
    );
    assert.deepEqual(
      strip(currentEntry.params),
      strip(goldenEntry.params),
      `request #${index + 1} filters must match the golden extension`
    );
  });
}

// Per-KPI selection has no golden oracle (the original extension has no tick boxes). The
// golden extension fetches only F4.1 and F1.3 for BC, so those two must produce
// post-office-keyed rows, while F1.1/F1.2 — which have no post-office breakdown — must fail
// loudly instead of silently emitting province-level rows.
async function testBcPerKpiKeepsPostOfficeIdentity() {
  const supported = {
    "F4.1 – Chất lượng phát thành công tại bưu cục": "F4.1",
    "F1.3 – Chất lượng phát bưu gửi liên tỉnh": "F1.3"
  };
  const unsupported = {
    "F1.1 – Nội tỉnh": "F1.1",
    "F1.2 – Thu gom bưu gửi đi liên tỉnh": "F1.2"
  };

  for (const [label, shortName] of Object.entries(supported)) {
    const sandbox = createCurrent(makeFetch(1));
    vm.runInContext(`getSelectedMetricLabels = () => ${JSON.stringify([label])};`, sandbox);
    const rows = await sandbox.fetchV2ReportRows("BC", new Date(2026, 7, 25), new Date(2026, 7, 25), {}, "kỳ báo cáo");
    const codes = Array.from(rows, (row) => String(row[1])).sort();

    assert.ok(
      !codes.includes(String(SELECTED_PROVINCE.code)),
      `${shortName}: BC rows must be keyed by post-office code, never by the province code ${SELECTED_PROVINCE.code}`
    );
    assert.deepEqual(codes, ["101", "102", "103", "105"], `${shortName}: BC export keeps every post office of the period`);
    assert.deepEqual(
      Array.from(rows, (row) => row[2]),
      ["BCVH Huế 1", "BCVH Huế 2", "BCVH Huế 3", "BCVH Phú Vang"],
      `${shortName}: BC rows carry the post-office name, not the province name`
    );
  }

  for (const [label, shortName] of Object.entries(unsupported)) {
    const sandbox = createCurrent(makeFetch(1));
    vm.runInContext(`getSelectedMetricLabels = () => ${JSON.stringify([label])};`, sandbox);
    await assert.rejects(
      () => sandbox.fetchV2ReportRows("BC", new Date(2026, 7, 25), new Date(2026, 7, 25), {}, "kỳ báo cáo"),
      /chỉ tiêu chất lượng/,
      `${shortName}: BC has no post-office breakdown, so it must refuse instead of inventing rows`
    );
  }
}

// Record-level mapping contract. The end-to-end BC flow survives a broken bcCodeIndex
// because pickV2JoinCodeText scans for any known unit code as a fallback, and bcValues
// happens to equal values for F4.1/F1.3. That redundancy must not be allowed to hide a
// province-level mapping leaking into BC, so the contract is pinned here directly, for all
// four KPIs, against the golden config table.
function testBcRecordMappingUsesPostOfficeConfig() {
  const current = createCurrent(async () => { throw new Error("not used"); });
  const golden = createGolden(async () => { throw new Error("not used"); });

  const goldenConfigs = JSON.parse(vm.runInContext(
    "JSON.stringify(V2_SOURCE_CONFIGS.map(c => ({ key: c.key, bcCodeIndex: c.bcCodeIndex, bcNameIndex: c.bcNameIndex, bcValues: c.bcValues })))",
    golden
  ));
  const currentConfigs = JSON.parse(vm.runInContext(
    "JSON.stringify(V2_SOURCE_CONFIGS.map(c => ({ key: c.key, bcCodeIndex: c.bcCodeIndex, bcNameIndex: c.bcNameIndex, bcValues: c.bcValues })))",
    current
  ));
  assert.deepEqual(currentConfigs, goldenConfigs, "the BC config table must match the golden extension");

  goldenConfigs.forEach((config) => {
    // Distinct sentinel values at every offset, so reading a province offset by mistake is
    // impossible to confuse with reading the post-office one.
    const cells = Array.from({ length: 30 }, (_, index) => String(900 + index));
    cells[1] = "53";                       // province code (the wrong key for BC)
    cells[2] = "Thừa Thiên Huế";           // province name (the wrong name for BC)
    cells[config.bcCodeIndex] = "101";
    cells[config.bcNameIndex] = "BCVH Huế 1";

    const html = `<tr>${cells.map((value) => `<td>${value}</td>`).join("")}</tr>`;
    const records = vm.runInContext(
      `JSON.stringify(parseV2ApiRows(${JSON.stringify(html)}, V2_SOURCE_CONFIGS.find(c => c.key === ${JSON.stringify(config.key)}), "BC", null))`,
      current
    );
    const record = JSON.parse(records)[0];
    assert.ok(record, `${config.key}: BC row is parsed`);

    assert.equal(record.provinceCodeInt, 101, `${config.key}: BC record is keyed by the post-office code, not the province code`);
    assert.equal(record.provinceName, "BCVH Huế 1", `${config.key}: BC record carries the post-office name, not the province name`);
    assert.equal(record.total, cells[config.bcValues.total], `${config.key}: BC total is read from bcValues.total`);
    assert.equal(record.onTime, cells[config.bcValues.onTime], `${config.key}: BC onTime is read from bcValues.onTime`);
    if (Number.isInteger(config.bcValues.subTotal)) {
      assert.equal(record.subTotal, cells[config.bcValues.subTotal], `${config.key}: BC subTotal is read from bcValues.subTotal`);
    }
  });
}

// Protection: restoring BC must not move the province comparison. buildV2ComparisonRows is
// pure, so both extensions are fed identical rows and their TINH output must be identical.
function testTinhCompareUnchanged() {
  const current = createCurrent(async () => { throw new Error("not used"); });
  const golden = createGolden(async () => { throw new Error("not used"); });

  const makeRows = (seed) => [10, 53, 90].map((code, index) => {
    const row = Array.from({ length: 37 }, () => "");
    row[0] = String(index + 1);
    row[1] = String(code);
    row[2] = `Tỉnh ${code}`;
    [5, 9, 13, 15, 19, 24, 29, 33, 35].forEach((column, position) => {
      row[column] = String(0.8 + (position + seed) / 100);
      row[column + 1] = String(position + 1);
    });
    return row;
  });

  const currentRows = makeRows(0);
  const compareRows = makeRows(1);
  const run = (sandbox) => JSON.parse(JSON.stringify(
    sandbox.buildV2ComparisonRows(currentRows, compareRows, "TINH")
  ));

  assert.deepEqual(run(current), run(golden), "TINH comparison rows must stay identical to the golden extension");

  // ...and the province path must still be free of every BC-only rule.
  const fetchSource = current.fetchV2ReportRows.toString();
  ["isBcOperatingRowWithData", "BCVH"].forEach((token) => {
    const bcGated = fetchSource.split("\n")
      .filter((line) => !line.trim().startsWith("//"))
      .filter((line) => line.includes(token));
    bcGated.forEach((line) => {
      assert.ok(
        /tuyChonGR === "BC"/.test(line) || /isBcOperatingRowWithData\(row, activeConfigs\)/.test(line),
        `BC-only rule "${token}" must stay behind a BC guard: ${line.trim()}`
      );
    });
  });
}

// Structural guards for the two failure modes that cannot be reached by data alone.
function testCompareFlowStaysIndependentOfMultiMonth() {
  const sandbox = createCurrent(async () => { throw new Error("not used"); });

  const exportSource = sandbox.exportReportV2.toString();
  assert.doesNotMatch(exportSource, /monthlyResults|fetchMultiMonthV2Report|generateMonthRanges/,
    "So sánh 2 kỳ must not reach into the multi-month pipeline");

  const compareSource = sandbox.buildV2ComparisonRows.toString();
  assert.match(compareSource, /compareRows\.map\(\(row\) => \[String\(row\[PROVINCE_CODE_INDEX\]\), row\]\)/,
    "the two periods are joined by unit code");
  assert.doesNotMatch(compareSource, /PROVINCE_NAME_INDEX\]\), row\]|\[String\(\w+\[PROVINCE_NAME_INDEX\]\)/,
    "the two periods must never be joined by unit name");
}

(async () => {
  testCompareFlowStaysIndependentOfMultiMonth();
  testTinhCompareUnchanged();
  testBcRecordMappingUsesPostOfficeConfig();
  await testBcCompareMatchesGolden();
  await testBcPerKpiKeepsPostOfficeIdentity();
  console.log("BC compare differential passed");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

if (typeof module !== "undefined") module.exports = { createCurrent, createGolden, makeFetch, BC_UNITS, BC_VOLUMES, buildFixtureHtml };
