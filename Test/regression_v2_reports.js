const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const popupSource = fs.readFileSync(path.join(__dirname, "..", "popup.js"), "utf8");

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

function createSandbox(fetchImpl, selectedProvince = { code: "", name: "" }) {
  const elements = new Map();
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
    localStorage: { getItem() { return null; }, setItem() {} },
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
}

function makeMonthlyRow(code, name, seed) {
  const row = Array(37).fill("");
  row[1] = String(code);
  row[2] = name;
  [17, 21, 26, 31].forEach((index) => { row[index] = String(100 + seed); });
  [18, 23, 28, 32].forEach((index) => { row[index] = String(90 + seed); });
  return row;
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
  console.log("V2 report regressions passed");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
