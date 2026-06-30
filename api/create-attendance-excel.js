import ExcelJS from "exceljs";
import { put } from "@vercel/blob";

function parseMaybeJson(value) {
  if (!value) return {};
  if (typeof value === "object") return value;

  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

function safeFileName(name) {
  const raw = String(name || "临时工工时待核对表.xlsx").trim();
  const cleaned = raw.replace(/[\\/:*?"<>|]/g, "_");
  return cleaned.endsWith(".xlsx") ? cleaned : `${cleaned}.xlsx`;
}

function safeSheetName(name, usedNames) {
  let s = String(name || "Sheet").trim();

  s = s.replace(/[\\/?*[\]:]/g, "");
  s = s.slice(0, 31) || "Sheet";

  let finalName = s;
  let i = 1;

  while (usedNames.has(finalName)) {
    const suffix = `_${i}`;
    finalName = s.slice(0, 31 - suffix.length) + suffix;
    i += 1;
  }

  usedNames.add(finalName);
  return finalName;
}

function isHeaderRow(row) {
  return Array.isArray(row) && row.includes("序号") && row.includes("姓名");
}

function isGroupRow(row) {
  return Array.isArray(row) && row.length === 1 && String(row[0] || "").startsWith("工作组：");
}

function isTitleRow(row) {
  return Array.isArray(row) && row.length === 1 && String(row[0] || "").includes("工时");
}

function styleWorksheet(ws) {
  ws.views = [{ state: "frozen", ySplit: 1 }];

  ws.eachRow((row) => {
    row.height = 22;

    row.eachCell((cell) => {
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true
      };

      cell.border = {
        top: { style: "thin", color: { argb: "FFD9D9D9" } },
        left: { style: "thin", color: { argb: "FFD9D9D9" } },
        bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
        right: { style: "thin", color: { argb: "FFD9D9D9" } }
      };
    });

    const values = row.values.slice(1);

    if (isTitleRow(values)) {
      row.height = 28;
      row.eachCell((cell) => {
        cell.font = { bold: true, size: 14 };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFE2F0D9" }
        };
      });
    }

    if (isGroupRow(values)) {
      row.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF2CC" }
        };
      });
    }

    if (isHeaderRow(values)) {
      row.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF4472C4" }
        };
      });
    }
  });

  for (let i = 1; i <= ws.columnCount; i++) {
    const col = ws.getColumn(i);

    if (i === 1) {
      col.width = 8;
    } else if (i === 2) {
      col.width = 14;
    } else if (i >= 3 && i <= 33) {
      col.width = 7;
    } else {
      col.width = 14;
    }
  }
}

function addRowsToSheet(ws, rows) {
  rows.forEach((row) => {
    ws.addRow(Array.isArray(row) ? row : [String(row || "")]);
  });

  if (ws.rowCount > 0 && ws.columnCount > 1) {
    try {
      ws.mergeCells(1, 1, 1, Math.min(ws.columnCount, 40));
      const firstRow = ws.getRow(1);
      firstRow.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true
      };
    } catch {
      // 如果合并失败，不影响文件生成
    }
  }
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

    const apiKey = process.env.API_KEY || "";
    if (apiKey) {
      const requestKey = req.headers["x-api-key"];
      if (requestKey !== apiKey) {
        return res.status(401).json({
          success: false,
          message: "Unauthorized"
        });
      }
    }

    const body = parseMaybeJson(req.body);
    const excelData = parseMaybeJson(body.excel_data);

    const fileName = safeFileName(
      body.file_name || excelData.file_name || "临时工工时待核对表.xlsx"
    );

    const sheets = Array.isArray(excelData.sheets) ? excelData.sheets : [];

    if (sheets.length === 0) {
      return res.status(400).json({
        success: false,
        message: "excel_data.sheets 为空，无法生成 Excel"
      });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Coze Attendance Bot";
    workbook.created = new Date();

    const usedSheetNames = new Set();

    for (const sheetData of sheets) {
      const sheetName = safeSheetName(sheetData.sheet_name, usedSheetNames);
      const rows = Array.isArray(sheetData.rows) ? sheetData.rows : [];

      const ws = workbook.addWorksheet(sheetName);
      addRowsToSheet(ws, rows);
      styleWorksheet(ws);
    }

    const buffer = await workbook.xlsx.writeBuffer();

    const timestamp = Date.now();
    const blobName = `attendance/${timestamp}_${fileName}`;

    const blob = await put(blobName, buffer, {
      access: "public",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    res.setHeader("Access-Control-Allow-Origin", "*");

    return res.status(200).json({
      success: true,
      file_name: fileName,
      file_url: blob.url,
      message: "Excel生成成功"
    });
  } catch (error) {
    console.error(error);

    return res.status(500).json({
      success: false,
      message: "Excel生成失败",
      error: error.message
    });
  }
}
