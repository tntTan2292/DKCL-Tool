// Characterization harness for "V2 -> So sánh lũy kế tháng -> Bưu cục".
//
// The golden extension has no multi-month feature, so there is no golden oracle here. The
// oracle is the fixture itself plus the rule the PO confirmed: the report must list every
// post office of the selected province, never a single province row.
const assert = require("node:assert/strict");
const path = require("node:path");
const vm = require("node:vm");

const { createCurrent } = require(path.join(__dirname, "differential_bc_compare.js"));

const SELECTED_PROVINCE = { code: 53, name: "Thừa Thiên Huế" };

const F41_PATH = "/kpi/chat-luong-phat-thanh-cong-cua-buu-cuc";
const F11_PATH = "/kpi/chat-luong-toan-trinh-buu-gui-noi-tinh";
const F12_PATH = "/kpi/chat-luong-thu-gom-buu-lien-tinh";
const F13_PATH = "/kpi/chat-luong-phat-buu-gui-lien-tinh";

const MONTH_KEYS = ["2026-01", "2026-02", "2026-03"];

// Post-office names differ by spacing / Unicode form across months while the code stays
// put, so any join or grouping by name would visibly break.
const BC_NAMES = {
  101: ["BCVH Huế 1", "BCVH  Huế 1", "BCVH Huế 1 "],
  102: ["BCVH Huế 2", "BCVH Huế 2", "BCVH Huế 2"],
  103: ["BCVH Huế 3", "BCVH Huế 3", "BCVH Huế 3"],
  105: [null, "BCVH Phú Vang", "BCVH Phú  Vang"]        // appears only from month 2
};

// [total, onTime] per KPI. null month = no data for that post office that month.
const BC_MONTHLY = {
  101: [
    { F41: [1000, 900], F13: [500, 440] },
    { F41: [1100, 1045], F13: [520, 452] },
    { F41: [1200, 1080], F13: [540, 486] }
  ],
  102: [
    { F41: [2000, 1700], F13: [900, 738] },
    { F41: [2100, 1869], F13: [950, 798] },
    { F41: [2200, 1958], F13: [1000, 850] }
  ],
  // 103 has no data in month 2: it must survive the whole period as N/A for that month.
  103: [
    { F41: [300, 240], F13: [150, 120] },
    null,
    { F41: [320, 272], F13: [160, 136] }
  ],
  105: [
    null,
    { F41: [5000, 4250], F13: [2500, 2000] },
    { F41: [5200, 4576], F13: [2600, 2158] }
  ]
};

const BC_CODES = Object.keys(BC_MONTHLY).map(Number).sort((a, b) => a - b);

function unitsInMonth(monthIndex) {
  return BC_CODES.filter((code) => BC_MONTHLY[code][monthIndex] && BC_NAMES[code][monthIndex]);
}

function buildMonthHtml(sourcePath, monthIndex) {
  return unitsInMonth(monthIndex).map((code) => {
    const name = BC_NAMES[code][monthIndex];
    const volumes = BC_MONTHLY[code][monthIndex];
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
      if (sourcePath === F13_PATH) {
        cells[10] = String(volumes.F13[0]);
        cells[18] = String(volumes.F13[1]);
      } else if (sourcePath === F11_PATH || sourcePath === F12_PATH) {
        // F1.1 / F1.2 have no post-office breakdown; the source returns nothing for BC.
        return "";
      }
    }
    return `<tr>${Array.from({ length: 30 }, (_, i) => `<td>${cells[i] ?? ""}</td>`).join("")}</tr>`;
  }).join("");
}

function monthAwareFetch() {
  return async (requestUrl) => {
    const url = new URL(requestUrl);
    const from = url.searchParams.get("iFrom") || "";           // MM/DD/YYYY
    const monthKey = `${from.slice(6, 10)}-${from.slice(0, 2)}`;
    const monthIndex = MONTH_KEYS.indexOf(monthKey);
    assert.ok(monthIndex >= 0, `fixture covers the requested month (${monthKey})`);
    return { ok: true, text: async () => JSON.stringify({ data: buildMonthHtml(url.pathname, monthIndex) }) };
  };
}

// --- expected values, derived from the fixture -----------------------------------------

function expectedRate(code, monthIndex, kpi) {
  const volumes = BC_MONTHLY[code][monthIndex];
  if (!volumes) return null;
  const [total, onTime] = volumes[kpi];
  return onTime / total;
}

function expectedVolume(code, monthIndex, kpi) {
  const volumes = BC_MONTHLY[code][monthIndex];
  return volumes ? volumes[kpi][0] : null;
}

// ---------------------------------------------------------------------------------------

async function runMultiMonthBc(sandbox, { labels } = {}) {
  if (labels) vm.runInContext(`getSelectedMetricLabels = () => ${JSON.stringify(labels)};`, sandbox);
  sandbox.fetch = monthAwareFetch();
  const monthsList = sandbox.generateMonthRanges(MONTH_KEYS[0], MONTH_KEYS[MONTH_KEYS.length - 1]);
  const monthlyResults = await sandbox.fetchMultiMonthV2Report(monthsList, "BC", SELECTED_PROVINCE.code, SELECTED_PROVINCE.name);
  return { monthsList, monthlyResults };
}

function nativePackage(sandbox, monthlyResults, monthsList) {
  const result = buildNative(sandbox, monthlyResults, monthsList);
  // Values coming back from the vm realm keep that realm's Array prototype, which
  // deepStrictEqual rejects. Round-trip the model into this realm.
  return { ...result, model: JSON.parse(JSON.stringify(result.model)) };
}

function buildNative(sandbox, monthlyResults, monthsList) {
  return sandbox.DkclXlsxExport.buildNativeV2MultiMonthXlsx({
    monthlyResults,
    monthsList,
    tuyChonGR: "BC",
    selectedProvinceCode: SELECTED_PROVINCE.code,
    selectedProvinceName: SELECTED_PROVINCE.name,
    metrics: sandbox.getActiveMultiMonthMetrics("BC"),
    columns: JSON.parse(vm.runInContext("JSON.stringify(V2_EXCEL_COLUMNS)", sandbox)),
    classifyKpi: sandbox.getKpiStatus,
    toNumberValue: sandbox.toNumberValue
  });
}

// ---------------------------------------------------------------------------------------
// 1. monthlyResults must be keyed by post-office code, per month
// ---------------------------------------------------------------------------------------

async function testMonthlyResultsKeyedByPostOffice() {
  const sandbox = createCurrent(monthAwareFetch());
  const { monthlyResults } = await runMultiMonthBc(sandbox);

  assert.equal(monthlyResults.length, MONTH_KEYS.length, "one result per selected month");

  monthlyResults.forEach(({ rows }, monthIndex) => {
    const codes = Array.from(rows, (row) => String(row[1])).sort();
    const names = Array.from(rows, (row) => row[2]);
    const expectedCodes = unitsInMonth(monthIndex).map(String).sort();

    assert.ok(
      !codes.includes(String(SELECTED_PROVINCE.code)),
      `month ${MONTH_KEYS[monthIndex]}: rows must not be keyed by the province code ${SELECTED_PROVINCE.code}`
    );
    assert.deepEqual(codes, expectedCodes, `month ${MONTH_KEYS[monthIndex]}: every post office of that month is listed`);
    names.forEach((name) => {
      assert.notEqual(name, SELECTED_PROVINCE.name, `month ${MONTH_KEYS[monthIndex]}: the province name must never stand in for a post office`);
      assert.match(name, /BCVH/, `month ${MONTH_KEYS[monthIndex]}: rows carry post-office names`);
    });

    // No post office may absorb another one's volume.
    rows.forEach((row) => {
      const code = Number(row[1]);
      assert.equal(sandbox.toNumberValue(row[17]), expectedVolume(code, monthIndex, "F41"), `BC ${code} month ${monthIndex + 1}: F4.1 volume is its own`);
      assert.equal(sandbox.toNumberValue(row[31]), expectedVolume(code, monthIndex, "F13"), `BC ${code} month ${monthIndex + 1}: F1.3 volume is its own`);
      assert.ok(Math.abs(Number(row[19]) - expectedRate(code, monthIndex, "F41")) < 1e-9, `BC ${code} month ${monthIndex + 1}: F4.1 rate`);
      assert.ok(Math.abs(Number(row[33]) - expectedRate(code, monthIndex, "F13")) < 1e-9, `BC ${code} month ${monthIndex + 1}: F1.3 rate`);
    });
  });
}

// ---------------------------------------------------------------------------------------
// 2. The exported workbook must list every post office, with per-month detail
// ---------------------------------------------------------------------------------------

async function testExportListsEveryPostOffice() {
  const sandbox = createCurrent(monthAwareFetch());
  const { monthsList, monthlyResults } = await runMultiMonthBc(sandbox);

  const metrics = sandbox.getActiveMultiMonthMetrics("BC");
  assert.deepEqual(
    Array.from(metrics, (metric) => metric.shortName).sort(),
    ["F1.3", "F4.1"],
    "BC uses only the KPIs the original extension supports"
  );

  const native = nativePackage(sandbox, monthlyResults, monthsList);
  const matrixCodes = Array.from(native.model.matrix, (unit) => String(unit.code)).sort();
  assert.deepEqual(matrixCodes, BC_CODES.map(String).sort(), "the .xlsx matrix lists every post office of the whole period");
  assert.ok(
    !matrixCodes.includes(String(SELECTED_PROVINCE.code)),
    "the province must never appear as a unit row in a BC report"
  );

  // A post office missing one month stays in the report, N/A for that month only.
  const unit103 = native.model.matrix.find((unit) => String(unit.code) === "103");
  assert.ok(unit103, "post office 103 survives the period despite a month with no data");
  assert.ok(Array.isArray(unit103.monthlyByMetric), "BC matrix carries a per-month breakdown");
  assert.equal(unit103.monthlyByMetric.length, metrics.length, "one series per KPI");
  assert.equal(unit103.monthlyByMetric[0].length, MONTH_KEYS.length, "one entry per selected month");

  const metricIndexOf = (shortName) => metrics.findIndex((metric) => metric.shortName === shortName);

  ["F4.1", "F1.3"].forEach((shortName) => {
    const kpi = shortName === "F4.1" ? "F41" : "F13";
    const metricIndex = metricIndexOf(shortName);

    BC_CODES.forEach((code) => {
      const unit = native.model.matrix.find((entry) => String(entry.code) === String(code));
      assert.ok(unit, `post office ${code} is present`);
      assert.match(unit.name, /BCVH/, `post office ${code} keeps its own name`);

      unit.monthlyByMetric[metricIndex].forEach((point, monthIndex) => {
        const expectedVol = expectedVolume(code, monthIndex, kpi);
        if (expectedVol == null) {
          assert.equal(point.hasData, false, `${shortName} BC ${code} month ${monthIndex + 1} is N/A`);
          assert.equal(point.rate, null, `${shortName} BC ${code} month ${monthIndex + 1} rate is N/A, not 0`);
          assert.equal(point.trend, null, `${shortName} BC ${code} month ${monthIndex + 1} trend is N/A`);
          assert.equal(point.rank, null, `${shortName} BC ${code} month ${monthIndex + 1}: N/A never receives a rank`);
          return;
        }
        assert.equal(point.hasData, true, `${shortName} BC ${code} month ${monthIndex + 1} has data`);
        assert.equal(point.volume, expectedVol, `${shortName} BC ${code} month ${monthIndex + 1} volume`);
        assert.ok(Math.abs(point.rate - expectedRate(code, monthIndex, kpi)) < 1e-9, `${shortName} BC ${code} month ${monthIndex + 1} rate`);

        const previousRate = expectedRate(code, monthIndex - 1, kpi);
        if (monthIndex === 0 || previousRate == null) {
          assert.equal(point.rateDelta, null, `${shortName} BC ${code} month ${monthIndex + 1}: no previous month -> no delta`);
        } else {
          const expectedDelta = expectedRate(code, monthIndex, kpi) - previousRate;
          assert.ok(Math.abs(point.rateDelta - expectedDelta) < 1e-9, `${shortName} BC ${code} month ${monthIndex + 1}: delta vs previous month`);
          assert.equal(point.trend, Math.sign(Number(expectedDelta.toFixed(10))), `${shortName} BC ${code} month ${monthIndex + 1}: trend direction`);
        }

        // The first month still carries a real rank of its own.
        assert.ok(Number.isInteger(point.rank) && point.rank >= 1, `${shortName} BC ${code} month ${monthIndex + 1}: has a rank (including the first month)`);
      });
    });
  });

  assert.match(
    native.model.rankingSample,
    /Bưu cục có dữ liệu hợp lệ thuộc tỉnh được chọn/,
    "the model states the BC ranking sample"
  );
  assert.doesNotMatch(native.model.rankingSample, /34|toàn quốc/i, "the BC ranking sample never claims a national scale");
}

// ---------------------------------------------------------------------------------------
// Ranking rules confirmed by the PO
// ---------------------------------------------------------------------------------------

function expectedRanking(monthIndex, kpi) {
  // Sample = post offices of the selected province with valid data, that month, that KPI.
  // Ordinal ranking (project rule, see assignRankByColumn); ties by ascending unit code.
  return BC_CODES
    .filter((code) => BC_MONTHLY[code][monthIndex])
    .map((code) => ({ code, rate: expectedRate(code, monthIndex, kpi) }))
    .sort((a, b) => b.rate - a.rate || a.code - b.code)
    .map((entry, index) => ({ code: entry.code, rank: index + 1 }));
}

async function testBcRankingRules() {
  const sandbox = createCurrent(monthAwareFetch());
  const { monthsList, monthlyResults } = await runMultiMonthBc(sandbox);
  const native = nativePackage(sandbox, monthlyResults, monthsList);
  const metrics = sandbox.getActiveMultiMonthMetrics("BC");

  ["F4.1", "F1.3"].forEach((shortName) => {
    const kpi = shortName === "F4.1" ? "F41" : "F13";
    const metricIndex = metrics.findIndex((metric) => metric.shortName === shortName);

    MONTH_KEYS.forEach((monthKey, monthIndex) => {
      const expected = expectedRanking(monthIndex, kpi);

      // 3 + 4: sample holds only this province's post offices that have data that month.
      const ranked = native.model.matrix
        .map((unit) => ({ code: Number(unit.code), point: unit.monthlyByMetric[metricIndex][monthIndex] }))
        .filter(({ point }) => point.rank != null);
      assert.deepEqual(
        ranked.map(({ code }) => code).sort((a, b) => a - b),
        expected.map((entry) => entry.code).sort((a, b) => a - b),
        `${shortName} ${monthKey}: only post offices with valid data are in the ranking sample`
      );
      ranked.forEach(({ point }) => {
        assert.equal(point.rankCount, expected.length, `${shortName} ${monthKey}: rank denominator is the sample size`);
      });

      // 1 + 2: the rank is computed per month and per KPI, highest rate first.
      expected.forEach(({ code, rank }) => {
        const unit = native.model.matrix.find((entry) => Number(entry.code) === code);
        assert.equal(
          unit.monthlyByMetric[metricIndex][monthIndex].rank,
          rank,
          `${shortName} ${monthKey}: BC ${code} rank`
        );
      });
    });
  });
}

// 6. Moving one post office's rate in one month must only reorder that KPI, that month.
async function testRankingIsIsolatedPerKpiAndMonth() {
  const baseline = createCurrent(monthAwareFetch());
  const baselineRun = await runMultiMonthBc(baseline);
  const baselineNative = nativePackage(baseline, baselineRun.monthlyResults, baselineRun.monthsList);
  const metrics = baseline.getActiveMultiMonthMetrics("BC");

  const snapshot = (native) => native.model.matrix.map((unit) => ({
    code: String(unit.code),
    ranks: unit.monthlyByMetric.map((points) => points.map((point) => point.rank))
  }));

  const before = snapshot(baselineNative);

  // Make BC 103 the best F4.1 performer in month 3 only.
  const original = BC_MONTHLY[103][2].F41;
  BC_MONTHLY[103][2].F41 = [320, 320];
  try {
    const mutated = createCurrent(monthAwareFetch());
    const mutatedRun = await runMultiMonthBc(mutated);
    const after = snapshot(nativePackage(mutated, mutatedRun.monthlyResults, mutatedRun.monthsList));

    const f41 = metrics.findIndex((metric) => metric.shortName === "F4.1");
    const f13 = metrics.findIndex((metric) => metric.shortName === "F1.3");
    const changedMonth = 2;

    assert.deepEqual(
      after.map((u) => u.ranks[f13]),
      before.map((u) => u.ranks[f13]),
      "changing an F4.1 rate must not move any F1.3 rank"
    );
    MONTH_KEYS.forEach((monthKey, monthIndex) => {
      if (monthIndex === changedMonth) return;
      assert.deepEqual(
        after.map((u) => u.ranks[f41][monthIndex]),
        before.map((u) => u.ranks[f41][monthIndex]),
        `changing month ${changedMonth + 1} must not move the F4.1 ranks of ${monthKey}`
      );
    });
    assert.equal(
      after.find((u) => u.code === "103").ranks[f41][changedMonth],
      1,
      "the post office whose rate became the highest takes rank 1 in that month"
    );
  } finally {
    BC_MONTHLY[103][2].F41 = original;
  }
}

// 5. A valid 0/0 stays in the sample under the project's existing rule (rate 0), and is
//    never turned into N/A — nor is N/A ever turned into 0.
async function testZeroOverZeroStaysInSample() {
  const original = BC_MONTHLY[102][0].F41;
  BC_MONTHLY[102][0].F41 = [0, 0];
  try {
    const sandbox = createCurrent(monthAwareFetch());
    const { monthsList, monthlyResults } = await runMultiMonthBc(sandbox);
    const native = nativePackage(sandbox, monthlyResults, monthsList);
    const metrics = sandbox.getActiveMultiMonthMetrics("BC");
    const f41 = metrics.findIndex((metric) => metric.shortName === "F4.1");

    const unit102 = native.model.matrix.find((unit) => String(unit.code) === "102");
    const point = unit102.monthlyByMetric[f41][0];
    assert.equal(point.hasData, true, "0/0 is valid data, not N/A");
    assert.equal(point.volume, 0, "0/0 keeps its zero volume");
    assert.equal(point.rate, 0, "0/0 rates as 0 under the current rule");
    assert.ok(Number.isInteger(point.rank), "0/0 takes part in the ranking");

    // 103 has no data at all in month 2 and must remain rankless there.
    const unit103 = native.model.matrix.find((unit) => String(unit.code) === "103");
    assert.equal(unit103.monthlyByMetric[f41][1].hasData, false, "a month with no data stays N/A");
    assert.equal(unit103.monthlyByMetric[f41][1].rank, null, "N/A is excluded from the ranking sample");
  } finally {
    BC_MONTHLY[102][0].F41 = original;
  }
}

// ---------------------------------------------------------------------------------------
// 3. Legacy .xls and native .xlsx must agree on the business values
// ---------------------------------------------------------------------------------------

async function testLegacyXlsMatchesNativeXlsx() {
  const sandbox = createCurrent(monthAwareFetch());
  const { monthsList, monthlyResults } = await runMultiMonthBc(sandbox);

  const legacyXml = sandbox.renderV2MultiMonthDashboardWorksheet(monthlyResults, monthsList, "BC");
  const native = nativePackage(sandbox, monthlyResults, monthsList);

  BC_CODES.forEach((code) => {
    assert.ok(legacyXml.includes(`>${code}<`), `legacy .xls lists post office ${code}`);
  });
  assert.ok(!/34 Tỉnh|34 tỉnh|Xếp hạng 34|toàn quốc/i.test(legacyXml), "a BC report must not claim a 34-province or national scope");
  assert.match(legacyXml, /Hạng trong các Bưu cục có dữ liệu hợp lệ thuộc tỉnh được chọn/, "the BC sheet states its ranking sample");

  // Every volume and every rank the native model reports must appear in the legacy sheet.
  native.model.matrix.forEach((unit) => {
    unit.monthlyByMetric.forEach((points) => {
      points.forEach((point) => {
        if (!point.hasData) return;
        assert.ok(
          legacyXml.includes(`>${point.volume}<`),
          `legacy .xls carries volume ${point.volume} of post office ${unit.code}`
        );
        assert.ok(
          legacyXml.includes(`>${point.rank}/${point.rankCount}<`),
          `legacy .xls carries rank ${point.rank}/${point.rankCount} of post office ${unit.code}`
        );
      });
    });
  });

  // Rank parity, cell by cell: the legacy row for each post office must carry exactly the
  // ranks the native model computed, in the same order.
  const legacyRows = [...legacyXml.matchAll(/<Row[^>]*>[\s\S]*?<\/Row>/g)].map((match) => match[0]);
  native.model.matrix.forEach((unit) => {
    const row = legacyRows.find((candidate) => candidate.includes(`>${unit.code}<`) && candidate.includes(unit.name));
    assert.ok(row, `legacy .xls has a row for post office ${unit.code}`);
    const legacyRanks = [...row.matchAll(/>(\d+)\/(\d+)</g)].map((match) => `${match[1]}/${match[2]}`);
    const nativeRanks = unit.monthlyByMetric
      .flat()
      .filter((point) => point.hasData)
      .map((point) => `${point.rank}/${point.rankCount}`);
    assert.deepEqual(legacyRanks, nativeRanks, `post office ${unit.code}: legacy .xls and native .xlsx agree on every rank`);
  });
}

// ---------------------------------------------------------------------------------------
// 4. F1.1 / F1.2 at BC level must refuse, not fabricate
// ---------------------------------------------------------------------------------------

async function testUnsupportedKpisRefuse() {
  for (const label of ["F1.1 – Nội tỉnh", "F1.2 – Thu gom bưu gửi đi liên tỉnh"]) {
    const sandbox = createCurrent(monthAwareFetch());
    await assert.rejects(
      () => runMultiMonthBc(sandbox, { labels: [label] }),
      /không hỗ trợ|chỉ tiêu chất lượng/i,
      `${label}: BC must say it is unsupported instead of emitting zero rows`
    );
  }
}

(async () => {
  await testMonthlyResultsKeyedByPostOffice();
  await testExportListsEveryPostOffice();
  await testBcRankingRules();
  await testRankingIsIsolatedPerKpiAndMonth();
  await testZeroOverZeroStaysInSample();
  await testLegacyXlsMatchesNativeXlsx();
  await testUnsupportedKpisRefuse();
  console.log("BC multi-month characterization passed");
})().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
