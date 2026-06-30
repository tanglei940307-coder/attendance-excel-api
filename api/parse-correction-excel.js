import ExcelJS from "exceljs";

function parseMaybeJson(value) {
  if (!value) return {};

  if (typeof value === "object") {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

function getApiKey(req) {
  return (
    req.headers["x-api-key"] ||
    req.headers["X-API-Key"] ||
    req.headers["authorization"] ||
    ""
  );
}

function checkApiKey(req) {
  const apiKey = process.env.API_KEY || "";
  const requestKey = String(getApiKey(req) || "").replace(/^Bearer\s+/i, "");

  if (!apiKey) {
    return false;
  }

  return requestKey === apiKey;
}

function getFileUrlFromBody(body) {
  if (!body) return "";

  const parsedBody = parseMaybeJson(body);

  if (parsedBody.file_url) return String(parsedBody.file_url);
  if (parsedBody.url) return String(parsedBody.url);
  if (parsedBody.excel_url) return String(parsedBody.excel_url);
  if (parsedBody.download_url) return String(parsedBody.download_url);

  const file =
    parsedBody.correction_file ||
    parsedBody.file ||
    parsedBody.excel_file ||
    parsedBody.input_file;

  if (typeof file === "string") {
    return file;
  }

  if (Array.isArray(file) && file.length > 0) {
    const first = file[0];

    if (typeof first === "string") return first;

    if (first && typeof first === "object") {
      if (first.url) return String(first.url);
      if (first.file_url) return String(first.file_url);
      if (first.download_url) return String(first.download_url);
      if (first.uri) return String(first.uri);
    }
  }

  if (file && typeof file === "object") {
    if (file.url) return String(file.url);
    if (file.file_url) return String(file.file_url);
    if (file.download_url) return String(file.download_url);
    if (file.uri) return String(file.uri);
  }

  return "";
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "");
}

function normalizeCell(value) {
  if (value === null || value === undefined) return "";

  if (typeof value === "object") {
    if (value.text !== undefined) {
      return String(value.text).trim();
    }

    if (value.result !== undefined) {
      return String(value.result).trim();
    }

    if (value.richText && Array.isArray(value.richText)) {
      return value.richText
        .map(item => item.text || "")
        .join("")
        .trim();
    }

    if (value.hyperlink && value.text) {
      return String(value.text).trim();
    }
  }

  return String(value).trim();
}

function worksheetToRows(worksheet) {
  const rows = [];

  worksheet.eachRow({ includeEmpty: true }, row => {
    const values = [];

    const maxColumn = worksheet.columnCount || row.cellCount || 0;

    for (let col = 1; col <= maxColumn; col++) {
      const cell = row.getCell(col);
      values.push(normalizeCell(cell.value));
    }

    rows.push(values);
  });

  return rows;
}

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const headers = row.map(normalizeHeader);

    const hasAbnormalId =
      headers.includes("异常ID") ||
      headers.includes("异常id") ||
      headers.includes("abnormal_id");

    const hasCorrectValue =
      headers.includes("修正值") ||
      headers.includes("correct_value") ||
      headers.includes("确认值");

    if (hasAbnormalId && hasCorrectValue) {
      return i;
    }
  }

  return -1;
}

function buildHeaderMap(headerRow) {
  const map = {};

  headerRow.forEach((cell, index) => {
    const key = normalizeHeader(cell);

    if (key) {
      map[key] = index;
    }
  });

  return map;
}

function getCell(row, headerMap, names) {
  for (const name of names) {
    const key = normalizeHeader(name);
    const index = headerMap[key];

    if (index !== undefined && index !== null) {
      const value = row[index];

      if (value !== undefined && value !== null) {
        return String(value).trim();
      }
    }
  }

  return "";
}

function getCellNumber(row, headerMap, names) {
  const value = getCell(row, headerMap, names);
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function parseCorrectionRows(workbook) {
  let worksheet = workbook.getWorksheet("异常待确认");

  if (!worksheet) {
    worksheet = workbook.worksheets.find(ws =>
      String(ws.name || "").includes("异常")
    );
  }

  if (!worksheet) {
    return {
      sheet_name: "",
      correction_rows: [],
      correction_count: 0,
      message: "未找到“异常待确认”工作表"
    };
  }

  const rows = worksheetToRows(worksheet);
  const headerRowIndex = findHeaderRow(rows);

  if (headerRowIndex < 0) {
    return {
      sheet_name: worksheet.name,
      correction_rows: [],
      correction_count: 0,
      message: "未找到包含“异常ID”和“修正值”的表头行"
    };
  }

  const headerRow = rows[headerRowIndex];
  const headerMap = buildHeaderMap(headerRow);
  const dataRows = rows.slice(headerRowIndex + 1);

  const correctionRows = [];

  for (const row of dataRows) {
    const abnormalId = getCell(row, headerMap, [
      "异常ID",
      "异常id",
      "abnormal_id"
    ]);

    const correctValue = getCell(row, headerMap, [
      "修正值",
      "correct_value",
      "确认值"
    ]);

    const correctRemark = getCell(row, headerMap, [
      "修正备注",
      "correct_remark",
      "备注"
    ]);

    if (!abnormalId) continue;
    if (!correctValue) continue;

    correctionRows.push({
      abnormal_id: abnormalId,
      batch_id: getCell(row, headerMap, ["批次ID", "batch_id"]),
      month: getCell(row, headerMap, ["月份", "month"]),
      company_name: getCell(row, headerMap, ["外包公司", "company_name"]),
      work_group: getCell(row, headerMap, ["工作组", "work_group"]),
      employee_name: getCell(row, headerMap, ["员工", "employee_name"]),
      attendance_date: getCell(row, headerMap, ["考勤日期", "attendance_date"]),
      day_number: getCellNumber(row, headerMap, ["日", "day_number"]),
      raw_text: getCell(row, headerMap, ["原始内容", "raw_text"]),
      reason: getCell(row, headerMap, ["异常原因", "reason"]),
      status: getCell(row, headerMap, ["当前状态", "status"]),
      source_image_index: getCell(row, headerMap, [
        "来源图片序号",
        "source_image_index"
      ]),
      source_image_name: getCell(row, headerMap, [
        "来源图片名称",
        "source_image_name"
      ]),
      correct_value: correctValue,
      correct_remark: correctRemark
    });
  }

  return {
    sheet_name: worksheet.name,
    correction_rows: correctionRows,
    correction_count: correctionRows.length,
    message: "解析成功"
  };
}

export default async function handler(req, res) {
  try {
    if (req.method === "OPTIONS") {
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key");
      return res.status(204).end();
    }

    if (req.method !== "POST") {
      return res.status(405).json({
        success: false,
        message: "Only POST is allowed"
      });
    }

    if (!checkApiKey(req)) {
      return res.status(401).json({
        success: false,
        message: "Unauthorized"
      });
    }

    const body = parseMaybeJson(req.body);
    const fileUrl = getFileUrlFromBody(body);

    if (!fileUrl) {
      return res.status(400).json({
        success: false,
        message: "缺少 Excel 文件链接 file_url"
      });
    }

    const fileResp = await fetch(fileUrl);

    if (!fileResp.ok) {
      return res.status(400).json({
        success: false,
        message: `Excel 文件下载失败：${fileResp.status}`
      });
    }

    const arrayBuffer = await fileResp.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const parsed = parseCorrectionRows(workbook);

    res.setHeader("Access-Control-Allow-Origin", "*");

    return res.status(200).json({
      success: true,
      message: parsed.message,
      sheet_name: parsed.sheet_name,
      correction_rows: parsed.correction_rows,
      correction_count: parsed.correction_count
    });
  } catch (error) {
    console.error(error);

    return res.status(500).json({
      success: false,
      message: "解析修正 Excel 失败",
      error: error.message
    });
  }
}
