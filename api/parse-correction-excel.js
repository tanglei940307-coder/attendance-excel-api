const XLSX = require("xlsx");

function sendJson(res, statusCode, data) {
  res.statusCode = statusCode;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(data));
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

  if (typeof body === "string") {
    try {
      body = JSON.parse(body);
    } catch (e) {
      return "";
    }
  }

  if (body.file_url) return String(body.file_url);
  if (body.url) return String(body.url);
  if (body.excel_url) return String(body.excel_url);

  const file = body.correction_file || body.file || body.excel_file;

  if (typeof file === "string") return file;

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

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const headers = row.map(normalizeHeader);

    const hasAbnormalId =
      headers.includes("异常ID") ||
      headers.includes("abnormal_id") ||
      headers.includes("异常id");

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
  const sheetName =
    workbook.SheetNames.find(name => String(name).trim() === "异常待确认") ||
    workbook.SheetNames.find(name => String(name).includes("异常")) ||
    "";

  if (!sheetName) {
    return {
      sheet_name: "",
      rows: [],
      message: "未找到“异常待确认”工作表"
    };
  }

  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const headerRowIndex = findHeaderRow(rows);

  if (headerRowIndex < 0) {
    return {
      sheet_name: sheetName,
      rows: [],
      message: "未找到包含“异常ID”和“修正值”的表头行"
    };
  }

  const headerRow = rows[headerRowIndex];
  const headerMap = buildHeaderMap(headerRow);
  const dataRows = rows.slice(headerRowIndex + 1);

  const corrections = [];

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

    corrections.push({
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
    sheet_name: sheetName,
    rows: corrections,
    message: "解析成功"
  };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return sendJson(res, 405, {
        success: false,
        message: "Method Not Allowed"
      });
    }

    if (!checkApiKey(req)) {
      return sendJson(res, 401, {
        success: false,
        message: "Unauthorized"
      });
    }

    const fileUrl = getFileUrlFromBody(req.body);

    if (!fileUrl) {
      return sendJson(res, 400, {
        success: false,
        message: "缺少 Excel 文件链接 file_url"
      });
    }

    const fileResp = await fetch(fileUrl);

    if (!fileResp.ok) {
      return sendJson(res, 400, {
        success: false,
        message: `Excel 文件下载失败：${fileResp.status}`
      });
    }

    const arrayBuffer = await fileResp.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellDates: false,
      cellText: true
    });

    const parsed = parseCorrectionRows(workbook);

    return sendJson(res, 200, {
      success: true,
      message: parsed.message,
      sheet_name: parsed.sheet_name,
      correction_rows: parsed.rows,
      correction_count: parsed.rows.length
    });
  } catch (error) {
    return sendJson(res, 500, {
      success: false,
      message: error && error.message ? error.message : "解析修正 Excel 失败"
    });
  }
};
