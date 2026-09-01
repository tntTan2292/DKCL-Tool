(function initDkclNativeXlsx(global) {
  "use strict";

  const XMLNS = {
    main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    packageRel: "http://schemas.openxmlformats.org/package/2006/relationships",
    chart: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    drawing: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    drawingMain: "http://schemas.openxmlformats.org/drawingml/2006/main"
  };

  function xmlEscape(value) {
    return String(value ?? "")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&apos;");
  }

  function columnName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name;
  }

  function quoteSheetName(name) {
    return `'${String(name).replaceAll("'", "''")}'`;
  }

  function sanitizeSheetName(name, fallback) {
    const cleaned = String(name || fallback || "Sheet")
      .replace(/[\\/*?:\[\]]/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 31);
    return cleaned || fallback || "Sheet";
  }

  function toNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const normalized = String(value ?? "").replace(/\s/g, "").replaceAll(".", "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function localKpiStatus(rate, target) {
    if (rate == null || !Number.isFinite(Number(rate))) return { key: "gray", label: "N/A", gapPoints: null };
    const gapPoints = (Number(rate) - Number(target)) * 100;
    if (gapPoints >= 0) return { key: "green", label: "Đạt mục tiêu", gapPoints };
    if (gapPoints >= -5 - 1e-9) return { key: "yellow", label: "Thấp hơn ≤ 5 điểm %", gapPoints };
    return { key: "red", label: "Thấp hơn > 5 điểm %", gapPoints };
  }

  function normalizeKpiStatus(status) {
    return { ...status, label: String(status?.label || "N/A").replace(/^[🟢🟡🔴⚪]\s*/u, "") };
  }

  function inlineCell(ref, value, styleId = 0) {
    if (value == null || value === "") return `<c r="${ref}" s="${styleId}"/>`;
    if (typeof value === "number" && Number.isFinite(value)) {
      return `<c r="${ref}" s="${styleId}"><v>${value}</v></c>`;
    }
    return `<c r="${ref}" s="${styleId}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  function formulaCell(ref, formula, cachedValue, styleId = 0) {
    const cached = cachedValue == null || !Number.isFinite(Number(cachedValue)) ? "" : Number(cachedValue);
    return `<c r="${ref}" s="${styleId}"><f>${xmlEscape(formula)}</f><v>${cached}</v></c>`;
  }

  function rowXml(rowNumber, cells, options = {}) {
    const height = options.height ? ` ht="${options.height}" customHeight="1"` : "";
    return `<row r="${rowNumber}"${height}>${cells.join("")}</row>`;
  }

  function worksheetXml({ rows, columns = [], merges = [], drawingRelId = "", freezeRow = 0, freezeCol = 0, autoFilter = "" }) {
    const colsXml = columns.length
      ? `<cols>${columns.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join("")}</cols>`
      : "";
    const paneParts = [];
    if (freezeRow) paneParts.push(`ySplit="${freezeRow}"`);
    if (freezeCol) paneParts.push(`xSplit="${freezeCol}"`);
    if (freezeRow || freezeCol) paneParts.push(`topLeftCell="${columnName(freezeCol)}${freezeRow + 1}"`);
    const pane = paneParts.length ? `<pane ${paneParts.join(" ")} state="frozen"/>` : "";
    const mergeXml = merges.length ? `<mergeCells count="${merges.length}">${merges.map(ref => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>` : "";
    const filterXml = autoFilter ? `<autoFilter ref="${autoFilter}"/>` : "";
    const drawingXml = drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${XMLNS.main}" xmlns:r="${XMLNS.rel}">
  <sheetViews><sheetView showGridLines="0" workbookViewId="0">${pane}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  ${colsXml}
  <sheetData>${rows.join("")}</sheetData>
  ${filterXml}${mergeXml}${drawingXml}
</worksheet>`;
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${XMLNS.main}">
  <fonts count="7">
    <font><sz val="10"/><name val="Arial"/><family val="2"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><color rgb="FF155724"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><color rgb="FF856404"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><color rgb="FF721C24"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><color rgb="FF5A6268"/><name val="Arial"/></font>
  </fonts>
  <fills count="9">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1B365D"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD4EDDA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF3CD"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8D7DA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2E3E5"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E1F2"/></left><right style="thin"><color rgb="FFD9E1F2"/></right><top style="thin"><color rgb="FFD9E1F2"/></top><bottom style="thin"><color rgb="FFD9E1F2"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="10" fontId="3" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="10" fontId="4" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="10" fontId="5" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="10" fontId="6" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="8" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }

  function statusStyle(status, text = false) {
    const styles = text
      ? { green: 9, yellow: 10, red: 11, gray: 12 }
      : { green: 5, yellow: 6, red: 7, gray: 8 };
    return styles[status?.key] ?? styles.gray;
  }

  function buildModel(options) {
    const {
      monthlyResults,
      monthsList,
      tuyChonGR,
      selectedProvinceCode,
      selectedProvinceName,
      metrics,
      classifyKpi = localKpiStatus
    } = options;

    if (!Array.isArray(monthlyResults) || !monthlyResults.length) throw new Error("Không có dữ liệu tháng để tạo .xlsx.");
    if (!Array.isArray(monthsList) || monthsList.length !== monthlyResults.length) throw new Error("Danh sách tháng không khớp dữ liệu .xlsx.");
    if (!Array.isArray(metrics) || metrics.length !== 4) throw new Error("Dashboard .xlsx yêu cầu đủ bốn KPI F1.1/F1.2/F1.3/F4.1.");

    const selectedCode = String(selectedProvinceCode);
    const unitMap = new Map();
    monthlyResults.forEach(({ rows }, monthIndex) => {
      (rows || []).forEach(row => {
        const code = String(row[1] ?? "");
        if (!unitMap.has(code)) unitMap.set(code, { code, name: row[2] || code, rows: Array(monthsList.length).fill(null) });
        unitMap.get(code).rows[monthIndex] = row;
      });
    });

    const units = [...unitMap.values()].sort((a, b) => Number(a.code) - Number(b.code) || a.code.localeCompare(b.code));
    if (tuyChonGR === "TINH" && units.length !== 34) {
      throw new Error(`Dashboard .xlsx cấp TINH yêu cầu đúng 34 tỉnh, nhận được ${units.length}.`);
    }

    const focus = unitMap.get(selectedCode);
    if (!focus) throw new Error(`Không tìm thấy tỉnh tiêu điểm mã ${selectedCode} (${selectedProvinceName}) trong dữ liệu .xlsx.`);
    const missingMonthIndex = focus.rows.findIndex(row => !row);
    if (missingMonthIndex >= 0) {
      throw new Error(`Tỉnh tiêu điểm mã ${selectedCode} (${selectedProvinceName}) không có bản ghi tháng ${monthsList[missingMonthIndex]?.label || missingMonthIndex + 1}.`);
    }

    const focusMetrics = metrics.map(metric => {
      let total = 0;
      let onTime = 0;
      const points = focus.rows.map((row, monthIndex) => {
        const volume = toNumber(row[metric.totalCol]);
        const completed = toNumber(row[metric.onTimeCol]);
        const rate = volume > 0 ? completed / volume : 0;
        total += volume;
        onTime += completed;
        return { month: monthsList[monthIndex].monthKey || monthsList[monthIndex].label, volume, onTime: completed, rate, target: metric.target };
      });
      const cumulativeRate = total > 0 ? onTime / total : 0;
      return { ...metric, points, total, onTime, cumulativeRate, status: normalizeKpiStatus(classifyKpi(cumulativeRate, metric.target)) };
    });

    const matrix = units.map(unit => ({
      code: unit.code,
      name: unit.name,
      metrics: metrics.map(metric => {
        let total = 0;
        let onTime = 0;
        let hasData = false;
        unit.rows.forEach(row => {
          if (!row) return;
          hasData = true;
          total += toNumber(row[metric.totalCol]);
          onTime += toNumber(row[metric.onTimeCol]);
        });
        const rate = hasData ? (total > 0 ? onTime / total : 0) : null;
        return { total, onTime, rate, status: normalizeKpiStatus(classifyKpi(rate, metric.target)) };
      })
    }));

    return { selectedCode, selectedName: selectedProvinceName || focus.name, metrics: focusMetrics, units, matrix };
  }

  function buildDashboardSheet(model, monthsList) {
    const rows = [];
    rows.push(rowXml(1, [inlineCell("A1", `DASHBOARD BI LŨY KẾ KPI – ${model.selectedName.toUpperCase()} (MÃ ${model.selectedCode})`, 1)], { height: 28 }));
    rows.push(rowXml(2, [inlineCell("A2", `Combo Chart native: cột sản lượng, đường tỷ lệ và đường mục tiêu riêng | ${monthsList[0].label} – ${monthsList[monthsList.length - 1].label}`, 13)], { height: 24 }));
    rows.push(rowXml(4, [
      inlineCell("A4", "KPI", 2), inlineCell("B4", "Mục tiêu", 2), inlineCell("C4", "Tỷ lệ lũy kế", 2),
      inlineCell("D4", "Chênh mục tiêu", 2), inlineCell("E4", "Trạng thái", 2), inlineCell("F4", "Sản lượng", 2)
    ], { height: 24 }));

    model.metrics.forEach((metric, index) => {
      const row = 5 + index;
      rows.push(rowXml(row, [
        inlineCell(`A${row}`, metric.shortName, 2),
        formulaCell(`B${row}`, `${quoteSheetName("DuLieu_Chart")}!${columnName(index + 1)}2`, metric.target, 4),
        formulaCell(`C${row}`, `IF(F${row}=0,0,SUM(${quoteSheetName("DuLieu_Chart")}!${columnName(2 + index * 4)}5:${columnName(2 + index * 4)}${4 + monthsList.length})/F${row})`, metric.cumulativeRate, statusStyle(metric.status)),
        formulaCell(`D${row}`, `C${row}-B${row}`, metric.status.gapPoints == null ? null : metric.status.gapPoints / 100, statusStyle(metric.status)),
        inlineCell(`E${row}`, metric.status.label, statusStyle(metric.status, true)),
        formulaCell(`F${row}`, `SUM(${quoteSheetName("DuLieu_Chart")}!${columnName(1 + index * 4)}5:${columnName(1 + index * 4)}${4 + monthsList.length})`, metric.total, 3)
      ], { height: 24 }));
    });

    rows.push(rowXml(10, [inlineCell("A10", "Biểu đồ native OOXML – mỗi KPI dùng trục phụ tỷ lệ 0–100%", 13)], { height: 22 }));
    return worksheetXml({
      rows,
      columns: [13, 13, 16, 18, 26, 16, 3, 12, 12, 12, 12, 3, 12, 12, 12, 12],
      merges: ["A1:P1", "A2:P2", "A10:P10"],
      drawingRelId: "rId1",
      freezeRow: 4
    });
  }

  function buildChartDataSheet(model, monthsList) {
    const rows = [];
    const targetHeaderCells = [inlineCell("A1", "Ngưỡng KPI SSOT", 2)];
    const targetValueCells = [inlineCell("A2", "Mục tiêu", 2)];
    model.metrics.forEach((metric, index) => {
      const col = columnName(index + 1);
      targetHeaderCells.push(inlineCell(`${col}1`, metric.shortName, 2));
      targetValueCells.push(inlineCell(`${col}2`, metric.target, 4));
    });
    rows.push(rowXml(1, targetHeaderCells, { height: 22 }));
    rows.push(rowXml(2, targetValueCells, { height: 22 }));

    const headers = [inlineCell("A4", "Tháng", 2)];
    model.metrics.forEach((metric, index) => {
      const base = 1 + index * 4;
      headers.push(inlineCell(`${columnName(base)}4`, `${metric.shortName} Sản lượng`, 2));
      headers.push(inlineCell(`${columnName(base + 1)}4`, `${metric.shortName} Đúng hạn`, 2));
      headers.push(inlineCell(`${columnName(base + 2)}4`, `${metric.shortName} Tỷ lệ`, 2));
      headers.push(inlineCell(`${columnName(base + 3)}4`, `${metric.shortName} Mục tiêu`, 2));
    });
    rows.push(rowXml(4, headers, { height: 32 }));

    monthsList.forEach((month, monthIndex) => {
      const row = 5 + monthIndex;
      const cells = [inlineCell(`A${row}`, month.monthKey || month.label, 0)];
      model.metrics.forEach((metric, metricIndex) => {
        const base = 1 + metricIndex * 4;
        const point = metric.points[monthIndex];
        cells.push(inlineCell(`${columnName(base)}${row}`, point.volume, 3));
        cells.push(inlineCell(`${columnName(base + 1)}${row}`, point.onTime, 3));
        cells.push(formulaCell(`${columnName(base + 2)}${row}`, `IF(${columnName(base)}${row}=0,0,${columnName(base + 1)}${row}/${columnName(base)}${row})`, point.rate, 4));
        cells.push(formulaCell(`${columnName(base + 3)}${row}`, `=${columnName(metricIndex + 1)}$2`, metric.target, 4));
      });
      rows.push(rowXml(row, cells, { height: 20 }));
    });

    return worksheetXml({
      rows,
      columns: [14, ...Array(16).fill(15)],
      freezeRow: 4,
      autoFilter: `A4:${columnName(16)}${4 + monthsList.length}`
    });
  }

  function buildMatrixSheet(model) {
    const rows = [];
    const headers = [inlineCell("A1", "Mã", 2), inlineCell("B1", "Tên đơn vị", 2)];
    model.metrics.forEach((metric, index) => {
      headers.push(inlineCell(`${columnName(2 + index * 3)}1`, `${metric.shortName} Sản lượng`, 2));
      headers.push(inlineCell(`${columnName(3 + index * 3)}1`, `${metric.shortName} Tỷ lệ`, 2));
      headers.push(inlineCell(`${columnName(4 + index * 3)}1`, `${metric.shortName} Trạng thái`, 2));
    });
    rows.push(rowXml(1, headers, { height: 32 }));
    model.matrix.forEach((unit, index) => {
      const row = index + 2;
      const cells = [inlineCell(`A${row}`, unit.code, 0), inlineCell(`B${row}`, unit.name, 0)];
      unit.metrics.forEach((metric, metricIndex) => {
        cells.push(inlineCell(`${columnName(2 + metricIndex * 3)}${row}`, metric.total, 3));
        cells.push(inlineCell(`${columnName(3 + metricIndex * 3)}${row}`, metric.rate, statusStyle(metric.status)));
        cells.push(inlineCell(`${columnName(4 + metricIndex * 3)}${row}`, metric.status.label, statusStyle(metric.status, true)));
      });
      rows.push(rowXml(row, cells, { height: 20 }));
    });
    return worksheetXml({
      rows,
      columns: [10, 28, ...Array(12).fill(16)],
      freezeRow: 1,
      freezeCol: 2,
      autoFilter: `A1:${columnName(13)}${model.matrix.length + 1}`
    });
  }

  function buildRawSheet(rowsData, columns) {
    const rows = [];
    const headers = (columns || Array.from({ length: 37 }, (_, index) => `Cột ${index + 1}`))
      .map((label, index) => inlineCell(`${columnName(index)}1`, label, 2));
    rows.push(rowXml(1, headers, { height: 36 }));
    (rowsData || []).forEach((sourceRow, rowIndex) => {
      const rowNumber = rowIndex + 2;
      const cells = headers.map((_, colIndex) => {
        const value = sourceRow[colIndex] ?? "";
        const ref = `${columnName(colIndex)}${rowNumber}`;
        if (colIndex > 2 && value !== "" && Number.isFinite(Number(value))) return inlineCell(ref, Number(value), 0);
        return inlineCell(ref, value, 0);
      });
      rows.push(rowXml(rowNumber, cells));
    });
    return worksheetXml({ rows, columns: [7, 10, 28, ...Array(Math.max(0, headers.length - 3)).fill(14)], freezeRow: 1, freezeCol: 3, autoFilter: `A1:${columnName(headers.length - 1)}${Math.max(1, rowsData.length + 1)}` });
  }

  function chartStringCache(values) {
    return `<c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${xmlEscape(value)}</c:v></c:pt>`).join("")}</c:strCache>`;
  }

  function chartNumberCache(values, formatCode) {
    return `<c:numCache><c:formatCode>${xmlEscape(formatCode)}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${Number(value)}</c:v></c:pt>`).join("")}</c:numCache>`;
  }

  function chartTitle(text) {
    return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="vi-VN" sz="1200" b="1"/><a:t>${xmlEscape(text)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`;
  }

  function seriesText(formula, text) {
    return `<c:tx><c:strRef><c:f>${xmlEscape(formula)}</c:f>${chartStringCache([text])}</c:strRef></c:tx>`;
  }

  function buildChartXml(metric, metricIndex, monthsList) {
    const sheetRef = quoteSheetName("DuLieu_Chart");
    const firstRow = 5;
    const lastRow = 4 + monthsList.length;
    const base = 1 + metricIndex * 4;
    const volumeCol = columnName(base);
    const rateCol = columnName(base + 2);
    const targetCol = columnName(base + 3);
    const categories = metric.points.map(point => point.month);
    const volumes = metric.points.map(point => point.volume);
    const rates = metric.points.map(point => point.rate);
    const targets = metric.points.map(point => point.target);
    const catFormula = `${sheetRef}!$A$${firstRow}:$A$${lastRow}`;
    const volumeFormula = `${sheetRef}!$${volumeCol}$${firstRow}:$${volumeCol}$${lastRow}`;
    const rateFormula = `${sheetRef}!$${rateCol}$${firstRow}:$${rateCol}$${lastRow}`;
    const targetFormula = `${sheetRef}!$${targetCol}$${firstRow}:$${targetCol}$${lastRow}`;
    const categoryAxisId = 100000 + metricIndex * 10 + 1;
    const volumeAxisId = 100000 + metricIndex * 10 + 2;
    const percentAxisId = 100000 + metricIndex * 10 + 3;
    const statusColors = { green: "70AD47", yellow: "FFC000", red: "C00000", gray: "A6A6A6" };
    const rateColor = statusColors[metric.status.key] || statusColors.gray;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${XMLNS.chart}" xmlns:a="${XMLNS.drawingMain}" xmlns:r="${XMLNS.rel}">
  <c:date1904 val="0"/><c:lang val="vi-VN"/><c:roundedCorners val="0"/>
  <c:chart>
    ${chartTitle(`${metric.shortName}: Sản lượng & tỷ lệ so với mục tiêu ${(metric.target * 100).toFixed(0)}%`)}
    <c:plotArea><c:layout/>
      <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
        <c:ser><c:idx val="0"/><c:order val="0"/>${seriesText(`${sheetRef}!$${volumeCol}$4`, `${metric.shortName} Sản lượng`)}
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>${xmlEscape(catFormula)}</c:f>${chartStringCache(categories)}</c:strRef></c:cat>
          <c:val><c:numRef><c:f>${xmlEscape(volumeFormula)}</c:f>${chartNumberCache(volumes, "#,##0")}</c:numRef></c:val>
        </c:ser><c:gapWidth val="90"/><c:axId val="${categoryAxisId}"/><c:axId val="${volumeAxisId}"/>
      </c:barChart>
      <c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
        <c:ser><c:idx val="1"/><c:order val="1"/>${seriesText(`${sheetRef}!$${rateCol}$4`, `${metric.shortName} Tỷ lệ`)}
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${rateColor}"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>
          <c:cat><c:strRef><c:f>${xmlEscape(catFormula)}</c:f>${chartStringCache(categories)}</c:strRef></c:cat>
          <c:val><c:numRef><c:f>${xmlEscape(rateFormula)}</c:f>${chartNumberCache(rates, "0%")}</c:numRef></c:val>
        </c:ser>
        <c:ser><c:idx val="2"/><c:order val="2"/>${seriesText(`${sheetRef}!$${targetCol}$4`, `${metric.shortName} Mục tiêu ${(metric.target * 100).toFixed(0)}%`)}
          <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="1F4E78"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>
          <c:cat><c:strRef><c:f>${xmlEscape(catFormula)}</c:f>${chartStringCache(categories)}</c:strRef></c:cat>
          <c:val><c:numRef><c:f>${xmlEscape(targetFormula)}</c:f>${chartNumberCache(targets, "0%")}</c:numRef></c:val>
        </c:ser><c:smooth val="0"/><c:axId val="${categoryAxisId}"/><c:axId val="${percentAxisId}"/>
      </c:lineChart>
      <c:catAx><c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="${volumeAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>
      <c:valAx><c:axId val="${volumeAxisId}"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorGridlines/><c:tickLblPos val="nextTo"/><c:crossAx val="${categoryAxisId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>
      <c:valAx><c:axId val="${percentAxisId}"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="1"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="${categoryAxisId}"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/>
  </c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>
</c:chartSpace>`;
  }

  function buildDrawingXml(metrics) {
    const anchors = [
      { fromCol: 0, fromRow: 10, toCol: 8, toRow: 27 },
      { fromCol: 8, fromRow: 10, toCol: 16, toRow: 27 },
      { fromCol: 0, fromRow: 28, toCol: 8, toRow: 45 },
      { fromCol: 8, fromRow: 28, toCol: 16, toRow: 45 }
    ];
    const frames = metrics.map((metric, index) => {
      const anchor = anchors[index];
      return `<xdr:twoCellAnchor>
        <xdr:from><xdr:col>${anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>${anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="${xmlEscape(metric.shortName)} Combo Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/>
          <a:graphic><a:graphicData uri="${XMLNS.chart}"><c:chart xmlns:c="${XMLNS.chart}" xmlns:r="${XMLNS.rel}" r:id="rId${index + 1}"/></a:graphicData></a:graphic>
        </xdr:graphicFrame><xdr:clientData/>
      </xdr:twoCellAnchor>`;
    }).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${XMLNS.drawing}" xmlns:a="${XMLNS.drawingMain}">${frames}</xdr:wsDr>`;
  }

  function buildDrawingRelationships(metrics) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${XMLNS.packageRel}">${metrics.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${XMLNS.rel}/chart" Target="../charts/chart${index + 1}.xml"/>`).join("")}</Relationships>`;
  }

  function buildWorkbookXml(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${XMLNS.main}" xmlns:r="${XMLNS.rel}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="15000"/></bookViews><sheets>${sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"${sheet.hidden ? ' state="hidden"' : ""}/>`).join("")}</sheets><calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`;
  }

  function buildWorkbookRelationships(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${XMLNS.packageRel}">${sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${XMLNS.rel}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="${XMLNS.rel}/styles" Target="styles.xml"/></Relationships>`;
  }

  function buildContentTypes(sheets, chartCount) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>${Array.from({ length: chartCount }, (_, index) => `<Override PartName="/xl/charts/chart${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  }

  function buildNativeV2MultiMonthXlsx(options) {
    if (!global.fflate || typeof global.fflate.zipSync !== "function" || typeof global.fflate.strToU8 !== "function") {
      throw new Error("fflate offline chưa được nạp; không thể đóng gói .xlsx native.");
    }

    const model = buildModel(options);
    const rawSheetNames = options.monthsList.map((month, index) => sanitizeSheetName(month.sheetName || month.monthKey || `Thang_${index + 1}`, `Thang_${index + 1}`));
    const usedNames = new Set(["Dashboard", "DuLieu_Chart", options.tuyChonGR === "TINH" ? "MaTran_34_Tinh" : "MaTran_BuuCuc"]);
    const uniqueRawNames = rawSheetNames.map((name, index) => {
      let candidate = name;
      let suffix = 2;
      while (usedNames.has(candidate)) candidate = `${name.slice(0, 27)}_${suffix++}`;
      usedNames.add(candidate);
      return candidate;
    });
    const sheets = [
      { name: "Dashboard", xml: buildDashboardSheet(model, options.monthsList) },
      { name: "DuLieu_Chart", xml: buildChartDataSheet(model, options.monthsList) },
      { name: options.tuyChonGR === "TINH" ? "MaTran_34_Tinh" : "MaTran_BuuCuc", xml: buildMatrixSheet(model) },
      ...options.monthlyResults.map(({ rows }, index) => ({ name: uniqueRawNames[index], xml: buildRawSheet(rows || [], options.columns) }))
    ];

    const files = {
      "[Content_Types].xml": buildContentTypes(sheets, model.metrics.length),
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${XMLNS.packageRel}"><Relationship Id="rId1" Type="${XMLNS.rel}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${XMLNS.rel}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
      "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>DKCL Tool</dc:creator><cp:lastModifiedBy>DKCL Tool</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified></cp:coreProperties>`,
      "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>DKCL Tool</Application><AppVersion>1.0</AppVersion></Properties>`,
      "xl/workbook.xml": buildWorkbookXml(sheets),
      "xl/_rels/workbook.xml.rels": buildWorkbookRelationships(sheets),
      "xl/styles.xml": buildStylesXml(),
      "xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${XMLNS.packageRel}"><Relationship Id="rId1" Type="${XMLNS.rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
      "xl/drawings/drawing1.xml": buildDrawingXml(model.metrics),
      "xl/drawings/_rels/drawing1.xml.rels": buildDrawingRelationships(model.metrics)
    };
    sheets.forEach((sheet, index) => { files[`xl/worksheets/sheet${index + 1}.xml`] = sheet.xml; });
    model.metrics.forEach((metric, index) => { files[`xl/charts/chart${index + 1}.xml`] = buildChartXml(metric, index, options.monthsList); });

    const zipped = {};
    Object.entries(files).forEach(([path, content]) => { zipped[path] = global.fflate.strToU8(content); });
    return {
      bytes: global.fflate.zipSync(zipped, { level: 6 }),
      model,
      parts: Object.keys(files).sort()
    };
  }

  global.DkclXlsxExport = Object.freeze({
    buildNativeV2MultiMonthXlsx,
    isAvailable: () => Boolean(global.fflate && typeof global.fflate.zipSync === "function")
  });
})(globalThis);
