const state = {
  line1File: null,
  line2File: null,
  materialFile: null,
  workbook: null,
  outputFileName: "",
  result: null,
};

const APP_VERSION = "20260722_1715";

const ids = {
  line1File: document.getElementById("line1File"),
  line2File: document.getElementById("line2File"),
  materialFile: document.getElementById("materialFile"),
  line1Name: document.getElementById("line1Name"),
  line2Name: document.getElementById("line2Name"),
  materialName: document.getElementById("materialName"),
  buildBtn: document.getElementById("buildBtn"),
  resetBtn: document.getElementById("resetBtn"),
  downloadBtn: document.getElementById("downloadBtn"),
  messages: document.getElementById("messages"),
  periodText: document.getElementById("periodText"),
  line1Count: document.getElementById("line1Count"),
  line2Count: document.getElementById("line2Count"),
  planCount: document.getElementById("planCount"),
  exceptionCount: document.getElementById("exceptionCount"),
  previewBody: document.getElementById("previewBody"),
};

const MONTHLY_REQUIRED = ["생산번호", "영업모델", "차대호기", "순번", "연번", "라인", "착수일", "CWT", "설계모델", "비고"];
const MATERIAL_REQUIRED = ["물류번호", "자재번호", "품명", "소요량", "발주량", "발주번호", "납기일자"];

ids.line1File.addEventListener("change", event => setFile("line1File", "line1Name", event));
ids.line2File.addEventListener("change", event => setFile("line2File", "line2Name", event));
ids.materialFile.addEventListener("change", event => setFile("materialFile", "materialName", event));
ids.buildBtn.addEventListener("click", buildPlan);
ids.downloadBtn.addEventListener("click", downloadWorkbook);
ids.resetBtn.addEventListener("click", () => location.reload());

function setFile(key, labelId, event) {
  const file = event.target.files[0] || null;
  state[key] = file;
  ids[labelId].textContent = file ? file.name : "파일을 선택하세요";
  ids.buildBtn.disabled = !(state.line1File && state.line2File && state.materialFile);
  ids.downloadBtn.disabled = true;
  state.workbook = null;
}

async function buildPlan() {
  try {
    setMessage("엑셀 파일을 읽고 있습니다...");
    const [line1Rows, line2Rows, materialRows] = await Promise.all([
      readExcelRows(state.line1File, 1, ["생산번호", "영업모델", "착수일"]),
      readExcelRows(state.line2File, 1, ["생산번호", "영업모델", "착수일"]),
      readExcelRows(state.materialFile, 0, ["물류번호", "자재번호", "품명"]),
    ]);

    const result = processRows(line1Rows, line2Rows, materialRows);
    state.result = result;
    state.workbook = createWorkbook(result);
    state.outputFileName = createOutputFileName(result.period);
    renderResult(result);
  } catch (error) {
    console.error(error);
    setMessage(`처리 중 오류가 발생했습니다.\n${error.message || error}`);
  }
}

function readExcelRows(file, headerRowIndex, requiredHeaders = []) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const workbook = XLSX.read(event.target.result, { type: "array", cellDates: true, raw: false });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
        const detectedHeaderRowIndex = detectHeaderRow(matrix, headerRowIndex, requiredHeaders);
        const headers = (matrix[detectedHeaderRowIndex] || []).map(normalizeHeader);
        const rows = matrix.slice(detectedHeaderRowIndex + 1)
          .filter(row => row.some(value => String(value ?? "").trim() !== ""))
          .map((row, index) => ({ rowNumber: detectedHeaderRowIndex + index + 2, values: row, headers }));
        resolve({ fileName: file.name, sheetName, headers, rows, headerRowIndex: detectedHeaderRowIndex });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function processRows(line1Source, line2Source, materialSource) {
  const exceptions = [];
  const summary = [];
  const line1Map = buildMonthlyMap(line1Source.headers);
  const line2Map = buildMonthlyMap(line2Source.headers);
  const materialMap = buildMaterialMap(materialSource.headers);

  validateColumns("월확정 통합1라인", line1Map, MONTHLY_REQUIRED, exceptions);
  validateColumns("월확정 통합2라인", line2Map, MONTHLY_REQUIRED, exceptions);
  validateColumns("물류번호별 자재소요현황", materialMap, MATERIAL_REQUIRED, exceptions);

  const line1Monthly = parseMonthlyRows(line1Source, line1Map, "통합1라인", exceptions, () => true);
  const line2Parsed = parseLine2Rows(line2Source, line2Map, exceptions);
  const line2Monthly = line2Parsed.hyundai;
  const materialItems = parseMaterialRows(materialSource, materialMap);

  const materialLine1 = filterMaterial(materialItems, "KPA10");
  const materialLine2 = filterMaterial(materialItems, "KPA20");
  const materialIndex = indexMaterials([...materialLine1, ...materialLine2], exceptions);
  const allMaterialsByLogistics = indexAllMaterials(materialItems);

  const monthly = [...line1Monthly, ...line2Monthly];
  const plans = [];
  for (const item of monthly) {
    const matches = materialIndex.get(item.productionNo) || [];
    if (matches.length === 1) {
      plans.push(toPlanRow(item, matches[0]));
    } else if (matches.length === 0) {
      const candidates = allMaterialsByLogistics.get(item.productionNo) || [];
      if (candidates.length) {
        const materialNos = candidates.map(row => row.materialNo).filter(Boolean).join(", ");
        const names = [...new Set(candidates.map(row => row.materialName).filter(Boolean))].join(" / ");
        exceptions.push(toException("카운터웨이트 품명 미확인", item.lineType, item.productionNo, item.salesModel, materialNos, `동일 물류번호 자재는 있으나 품명에서 카운터웨이트를 확인하지 못했습니다. 품명: ${names}`));
      } else {
        exceptions.push(toException("자재소요 미매칭", item.lineType, item.productionNo, item.salesModel, "", "생산번호와 일치하는 카운터웨이트 자재소요가 없습니다."));
      }
    } else {
      exceptions.push(toException("자재소요 중복", item.lineType, item.productionNo, item.salesModel, matches.map(row => row.materialNo).join(", "), "동일 생산번호에 카운터웨이트 자재소요가 2건 이상입니다."));
    }
  }

  const monthlyKeys = new Set(monthly.map(row => row.productionNo));
  for (const mat of [...materialLine1, ...materialLine2]) {
    if (!monthlyKeys.has(mat.logisticsNo)) {
      exceptions.push(toException("월확정 미매칭", mat.lineType, mat.logisticsNo, "", mat.materialNo, "물류번호와 일치하는 월확정 생산번호가 없습니다."));
    }
  }

  plans.sort(comparePlanRows);
  const period = getPeriod(plans);

  summary.push(["구분", "건수", "비고"]);
  summary.push(["월확정 통합1라인 전체", line1Source.rows.length, line1Source.fileName]);
  summary.push(["월확정 통합1라인 사용", line1Monthly.length, "전체 사용"]);
  summary.push(["월확정 통합2라인 전체", line2Source.rows.length, line2Source.fileName]);
  summary.push(["월확정 통합2라인 현대 제품", line2Monthly.length, "영업모델 HX~/R~"]);
  summary.push(["월확정 통합2라인 DX 제외", line2Parsed.dx.length, "예외목록 기록"]);
  summary.push(["자재소요 전체", materialSource.rows.length, materialSource.fileName]);
  summary.push(["자재소요 KPA10 카운터웨이트", materialLine1.length, "통합1라인 매칭 대상"]);
  summary.push(["자재소요 KPA20 카운터웨이트", materialLine2.length, "통합2라인 매칭 대상"]);
  summary.push(["정상 생산계획", plans.length, "생산계획 시트"]);
  summary.push(["예외", exceptions.length, "예외목록 시트"]);
  summary.push(["착수기간", period.label, "월확정 착수일 기준"]);

  return { plans, exceptions, summary, period, counts: { line1: line1Monthly.length, line2: line2Monthly.length } };
}

function buildMonthlyMap(headers) {
  return {
    productionNo: findHeader(headers, ["생산번호", "물류번호"]),
    salesModel: findHeader(headers, ["영업모델"]),
    machineNo: findHeader(headers, ["차대호기"]),
    seq: findHeader(headers, ["순번"], "last"),
    serial: findHeader(headers, ["연번"]),
    line: findHeader(headers, ["라인"]),
    startDate: findHeader(headers, ["착수일", "착수일자"]),
    radar: findHeader(headers, ["RADAR", "Radar"]),
    cwt: findHeader(headers, ["CWT"]),
    designModel: findHeader(headers, ["설계모델"]),
    remark: findHeader(headers, ["비고"]),
  };
}

function buildMaterialMap(headers) {
  const map = {
    logisticsNo: findHeader(headers, ["물류번호", "생산번호"]),
    materialNo: findHeader(headers, ["자재번호"]),
    materialName: findHeader(headers, ["품명"]),
    requiredQty: findHeader(headers, ["소요량"]),
    orderQty: findHeader(headers, ["발주량"]),
    poNo: findHeader(headers, ["발주번호"]),
    dueDate: findHeader(headers, ["납기일자", "납품예정일"]),
  };
  applyMaterialPositionFallback(map, headers);
  return map;
}

function validateColumns(label, map, required, exceptions) {
  const missing = required.filter(name => {
    if (name === "RADAR") return false;
    const key = toMapKey(name);
    return map[key] === null || map[key] === undefined;
  });
  if (missing.length) {
    exceptions.push(toException("필수 컬럼 누락", label, "", "", "", `${label}: ${missing.join(", ")} 컬럼을 찾을 수 없습니다.`));
  }
}

function toMapKey(koreanName) {
  return ({
    "생산번호": "productionNo",
    "영업모델": "salesModel",
    "차대호기": "machineNo",
    "순번": "seq",
    "연번": "serial",
    "라인": "line",
    "착수일": "startDate",
    "CWT": "cwt",
    "설계모델": "designModel",
    "비고": "remark",
    "물류번호": "logisticsNo",
    "자재번호": "materialNo",
    "품명": "materialName",
    "소요량": "requiredQty",
    "발주량": "orderQty",
    "발주번호": "poNo",
    "납기일자": "dueDate",
  })[koreanName];
}

function parseMonthlyRows(source, map, lineType, exceptions, predicate) {
  const parsed = [];
  for (const row of source.rows) {
    const data = makeMonthly(row, map, lineType);
    if (!data.productionNo) {
      exceptions.push(toException("필수 값 누락", lineType, "", data.salesModel, "", `${source.fileName} ${row.rowNumber}행 생산번호가 비어 있습니다.`));
      continue;
    }
    if (predicate(data)) parsed.push(data);
  }
  return parsed;
}

function parseLine2Rows(source, map, exceptions) {
  const hyundai = [];
  const dx = [];
  for (const row of source.rows) {
    const data = makeMonthly(row, map, "통합2라인");
    if (!data.productionNo) continue;
    if (isHyundaiModel(data.salesModel)) {
      hyundai.push(data);
    } else if (String(data.salesModel).startsWith("DX")) {
      dx.push(data);
      exceptions.push(toException("DX 제품 제외", "통합2라인", data.productionNo, data.salesModel, "", "통합2라인 디벨론 제품은 생산계획 I/F 대상에서 제외합니다."));
    } else {
      exceptions.push(toException("모델 필터 제외", "통합2라인", data.productionNo, data.salesModel, "", "통합2라인 현대 제품 기준 HX~/R~에 해당하지 않습니다."));
    }
  }
  return { hyundai, dx };
}

function makeMonthly(row, map, lineType) {
  return {
    productionNo: normalizeKey(valueAt(row, map.productionNo)),
    salesModel: valueAt(row, map.salesModel),
    machineNo: valueAt(row, map.machineNo),
    seq: valueAt(row, map.seq),
    serial: valueAt(row, map.serial),
    line: valueAt(row, map.line),
    lineType,
    startDate: normalizeDate(valueAt(row, map.startDate)),
    radar: valueAt(row, map.radar),
    cwt: valueAt(row, map.cwt),
    designModel: valueAt(row, map.designModel),
    remark: valueAt(row, map.remark),
  };
}

function parseMaterialRows(source, map) {
  return source.rows.map(row => ({
    logisticsNo: normalizeKey(valueAt(row, map.logisticsNo)),
    materialNo: valueAt(row, map.materialNo),
    materialName: valueAt(row, map.materialName),
    requiredQty: valueAt(row, map.requiredQty),
    orderQty: valueAt(row, map.orderQty),
    poNo: valueAt(row, map.poNo),
    dueDate: normalizeDate(valueAt(row, map.dueDate)),
    lineType: getLineTypeFromLogisticsNo(valueAt(row, map.logisticsNo)),
  })).filter(row => row.logisticsNo);
}

function filterMaterial(rows, prefix) {
  return rows.filter(row => row.logisticsNo.startsWith(prefix) && isCounterweightName(row.materialName));
}

function isCounterweightName(name) {
  return normalizeSearchText(name).includes("카운터웨이트");
}

function indexMaterials(rows, exceptions) {
  const map = new Map();
  for (const row of rows) {
    if (!map.has(row.logisticsNo)) map.set(row.logisticsNo, []);
    map.get(row.logisticsNo).push(row);
  }
  for (const [key, matches] of map.entries()) {
    if (matches.length > 1) {
      exceptions.push(toException("자재소요 중복", matches[0].lineType, key, "", matches.map(m => m.materialNo).join(", "), "동일 물류번호에 카운터웨이트 자재가 2건 이상 존재합니다."));
    }
  }
  return map;
}

function indexAllMaterials(rows) {
  const map = new Map();
  for (const row of rows) {
    if (!map.has(row.logisticsNo)) map.set(row.logisticsNo, []);
    map.get(row.logisticsNo).push(row);
  }
  return map;
}

function toPlanRow(monthly, material) {
  return {
    "착수일": monthly.startDate,
    "라인": monthly.lineType,
    "생산번호(물류번호)": monthly.productionNo,
    "자재번호": material.materialNo,
    "품명": material.materialName,
    "모델명": monthly.salesModel,
    "차대호기": monthly.machineNo,
    "순번": monthly.seq,
    "연번": monthly.serial,
    "CWT": monthly.cwt,
    "RADAR": monthly.radar,
    "설계모델": monthly.designModel,
    "소요량": material.requiredQty,
    "발주량": material.orderQty,
    "발주번호": material.poNo,
    "납기일자": material.dueDate,
    "비고": monthly.remark,
  };
}

function toException(type, line, key, model, materialNo, reason) {
  return {
    "예외유형": type,
    "라인": line,
    "생산번호/물류번호": key,
    "영업모델": model,
    "자재번호": materialNo,
    "사유": reason,
  };
}

function createWorkbook(result) {
  const wb = XLSX.utils.book_new();
  const planSheet = XLSX.utils.json_to_sheet(result.plans);
  const exceptionSheet = XLSX.utils.json_to_sheet(result.exceptions.length ? result.exceptions : [toException("예외 없음", "", "", "", "", "예외가 없습니다.")]);
  const summarySheet = XLSX.utils.aoa_to_sheet(result.summary);
  XLSX.utils.book_append_sheet(wb, planSheet, "생산계획");
  XLSX.utils.book_append_sheet(wb, exceptionSheet, "예외목록");
  XLSX.utils.book_append_sheet(wb, summarySheet, "검증요약");
  return wb;
}

function renderResult(result) {
  ids.line1Count.textContent = `${result.counts.line1}건`;
  ids.line2Count.textContent = `${result.counts.line2}건`;
  ids.planCount.textContent = `${result.plans.length}건`;
  ids.exceptionCount.textContent = `${result.exceptions.length}건`;
  ids.periodText.textContent = result.period.label;
  ids.downloadBtn.disabled = false;
  setMessage(`생산계획 생성이 완료되었습니다. (${APP_VERSION})\n정상 ${result.plans.length}건, 예외 ${result.exceptions.length}건\n출력 파일명: ${state.outputFileName}`);
  renderPreview(result.plans);
}

function renderPreview(plans) {
  const rows = plans.slice(0, 20);
  if (!rows.length) {
    ids.previewBody.innerHTML = '<tr><td colspan="8">정상 생산계획 데이터가 없습니다.</td></tr>';
    return;
  }
  ids.previewBody.innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row["착수일"])}</td>
      <td>${escapeHtml(row["라인"])}</td>
      <td>${escapeHtml(row["생산번호(물류번호)"])}</td>
      <td>${escapeHtml(row["자재번호"])}</td>
      <td>${escapeHtml(row["모델명"])}</td>
      <td>${escapeHtml(row["차대호기"])}</td>
      <td>${escapeHtml(row["CWT"])}</td>
      <td>${escapeHtml(row["RADAR"])}</td>
    </tr>
  `).join("");
}

function downloadWorkbook() {
  if (!state.workbook || !state.outputFileName) return;
  XLSX.writeFile(state.workbook, state.outputFileName);
}

function findHeader(headers, names, mode = "first") {
  const targets = (Array.isArray(names) ? names : [names]).map(name => normalizeHeader(name).toUpperCase());
  const matches = [];
  headers.forEach((header, index) => {
    const current = normalizeHeader(header).toUpperCase();
    if (targets.includes(current)) matches.push(index);
  });
  if (!matches.length) return null;
  return mode === "last" ? matches[matches.length - 1] : matches[0];
}

function detectHeaderRow(matrix, preferredIndex, requiredHeaders) {
  if (!requiredHeaders.length) return preferredIndex;
  const preferredScore = scoreHeaderRow(matrix[preferredIndex] || [], requiredHeaders);
  let bestIndex = preferredIndex;
  let bestScore = preferredScore;

  const scanLimit = Math.min(matrix.length, 12);
  for (let index = 0; index < scanLimit; index += 1) {
    const score = scoreHeaderRow(matrix[index] || [], requiredHeaders);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }

  return bestScore >= Math.min(2, requiredHeaders.length) ? bestIndex : preferredIndex;
}

function scoreHeaderRow(row, requiredHeaders) {
  const normalizedCells = row.map(value => normalizeHeader(value).toUpperCase());
  return requiredHeaders.reduce((score, header) => {
    const target = normalizeHeader(header).toUpperCase();
    return score + (normalizedCells.includes(target) ? 1 : 0);
  }, 0);
}

function applyMaterialPositionFallback(map, headers) {
  const hasEnoughColumns = headers.length >= 24;
  if (!hasEnoughColumns) return;

  // 물류번호별 자재소요현황 표준 다운로드 구조: D/E/F/I/J/U/X
  const fallback = {
    logisticsNo: 3,
    materialNo: 4,
    materialName: 5,
    requiredQty: 8,
    orderQty: 9,
    poNo: 20,
    dueDate: 23,
  };

  for (const [key, index] of Object.entries(fallback)) {
    if (map[key] === null || map[key] === undefined) {
      map[key] = index;
    }
  }
}

function valueAt(row, index) {
  if (index === null || index === undefined) return "";
  const value = row.values[index];
  return String(value ?? "").trim();
}

function normalizeHeader(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function normalizeKey(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function normalizeSearchText(value) {
  return String(value ?? "").replace(/\s+/g, "").toUpperCase();
}

function normalizeDate(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const dateMatch = raw.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (dateMatch) return `${dateMatch[1]}-${dateMatch[2].padStart(2, "0")}-${dateMatch[3].padStart(2, "0")}`;
  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slashMatch) {
    const year = slashMatch[3].length === 2 ? `20${slashMatch[3]}` : slashMatch[3];
    return `${year}-${slashMatch[1].padStart(2, "0")}-${slashMatch[2].padStart(2, "0")}`;
  }
  const koreanMatch = raw.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
  if (koreanMatch) return `${koreanMatch[1]}-${koreanMatch[2].padStart(2, "0")}-${koreanMatch[3].padStart(2, "0")}`;
  return raw;
}

function getPeriod(plans) {
  const dates = plans.map(row => row["착수일"]).filter(Boolean).sort();
  if (!dates.length) return { min: "", max: "", label: "착수일 없음" };
  return { min: dates[0], max: dates[dates.length - 1], label: `${dates[0]} ~ ${dates[dates.length - 1]}` };
}

function createOutputFileName(period) {
  const start = (period.min || "START").replaceAll("-", "");
  const end = (period.max || "END").replaceAll("-", "");
  return `원우ENG_CWT_생산계획_${start}-${end}.xlsx`;
}

function comparePlanRows(a, b) {
  return String(a["착수일"]).localeCompare(String(b["착수일"]))
    || String(a["라인"]).localeCompare(String(b["라인"]))
    || numericCompare(a["순번"], b["순번"])
    || String(a["연번"]).localeCompare(String(b["연번"]))
    || String(a["생산번호(물류번호)"]).localeCompare(String(b["생산번호(물류번호)"]));
}

function numericCompare(a, b) {
  const na = Number(a);
  const nb = Number(b);
  if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
  return String(a).localeCompare(String(b));
}

function isHyundaiModel(model) {
  const value = String(model ?? "").trim();
  return value.startsWith("HX") || value.startsWith("R");
}

function getLineTypeFromLogisticsNo(value) {
  const logisticsNo = normalizeKey(value);
  if (logisticsNo.startsWith("KPA10")) return "통합1라인";
  if (logisticsNo.startsWith("KPA20")) return "통합2라인";
  return "라인 미분류";
}

function setMessage(message) {
  ids.messages.textContent = message;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

