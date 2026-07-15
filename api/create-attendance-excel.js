import ExcelJS from "exceljs";
import { put } from "@vercel/blob";

function parseMaybeJson(value) {
  if (!value) return {};
  if (typeof value === "object" && value !== null) return value;

  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

function text(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function safeFileName(name) {
  const raw = text(name) || "临时工工时待核对表.xlsx";
  const cleaned = raw.replace(/[\\/:*?"<>|]/g, "_");
  return cleaned.toLowerCase().endsWith(".xlsx")
    ? cleaned
    : `${cleaned}.xlsx`;
}

function safeSheetName(name, usedNames) {
  let base = text(name) || "Sheet";
  base = base.replace(/[\\/:?*\[\]]/g, "").slice(0, 31) || "Sheet";

  let finalName = base;
  let index = 1;

  while (usedNames.has(finalName)) {
    const suffix = `_${index}`;
    finalName = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }

  usedNames.add(finalName);
  return finalName;
}

function solidFill(argb) {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb }
  };
}

function thinBorder(color = "FFD6DEE8") {
  return {
    top: { style: "thin", color: { argb: color } },
    left: { style: "thin", color: { argb: color } },
    bottom: { style: "thin", color: { argb: color } },
    right: { style: "thin", color: { argb: color } }
  };
}

function rowValues(row) {
  return row.values.slice(1).map((value) => text(value));
}

function isBlankRow(values) {
  return values.every((value) => value === "");
}

function isTitleRow(values) {
  return text(values[0]).includes("临时工工时待核对表");
}

function isCompanyRow(values) {
  return text(values[0]).startsWith("外包公司：");
}

function isGroupRow(values) {
  return text(values[0]).startsWith("工作组：");
}

function isHeaderRow(values) {
  return values.includes("序号") && values.includes("姓名");
}

function isErrorTitleRow(values) {
  return text(values[0]) === "异常待确认";
}

function isErrorInstructionRow(values) {
  const first = text(values[0]);
  return first.includes("修正值") && first.includes("修正备注");
}

function isErrorHeaderRow(values) {
  return values.includes("异常ID") && values.includes("异常原因");
}

function getWeekendDays(monthValue) {
  const result = new Set();
  const match = text(monthValue).match(/^(\d{4})-(\d{1,2})$/);

  if (!match) return result;

  const year = Number(match[1]);
  const month = Number(match[2]);

  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    return result;
  }

  const daysInMonth = new Date(year, month, 0).getDate();

  for (let day = 1; day <= daysInMonth; day += 1) {
    const weekday = new Date(year, month - 1, day).getDay();
    if (weekday === 0 || weekday === 6) result.add(day);
  }

  return result;
}

function mergeWholeRow(worksheet, rowNumber) {
  if (worksheet.columnCount <= 1) return;

  try {
    worksheet.mergeCells(rowNumber, 1, rowNumber, worksheet.columnCount);
  } catch {
    // 合并失败时继续生成文件
  }
}

function setPageLayout(worksheet) {
  worksheet.pageSetup = {
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    verticalCentered: false,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.55,
      bottom: 0.55,
      header: 0.2,
      footer: 0.2
    }
  };

  worksheet.headerFooter = {
    oddHeader: "&C&14&B临时工考勤核对表",
    oddFooter: "&L考勤识别系统&C第 &P 页 / 共 &N 页&R&F"
  };
}

function styleCellBase(cell, borderColor = "FFD6DEE8") {
  cell.font = {
    name: "Microsoft YaHei",
    size: 10,
    color: { argb: "FF1F2937" }
  };

  cell.alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true
  };

  cell.border = thinBorder(borderColor);
}

function styleAttendanceSheet(worksheet, monthValue) {
  const weekendDays = getWeekendDays(monthValue);
  const dateStartColumn = 3;
  const dateEndColumn = 33;
  const summaryStartColumn = 34;
  const summaryEndColumn = 39;
  let firstHeaderRow = 0;

  setPageLayout(worksheet);

  worksheet.properties = {
    ...worksheet.properties,
    defaultRowHeight: 22,
    tabColor: { argb: "FF2F5597" }
  };

  worksheet.getColumn(1).width = 8;
  worksheet.getColumn(2).width = 24;

  for (let columnNumber = 3; columnNumber <= 33; columnNumber += 1) {
    worksheet.getColumn(columnNumber).width = 5.5;
  }

  worksheet.getColumn(34).width = 11;
  worksheet.getColumn(35).width = 11;
  worksheet.getColumn(36).width = 8;
  worksheet.getColumn(37).width = 8;
  worksheet.getColumn(38).width = 8;
  worksheet.getColumn(39).width = 24;

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values = rowValues(row);

    if (isTitleRow(values) || isCompanyRow(values) || isGroupRow(values)) {
      mergeWholeRow(worksheet, rowNumber);
    }
  }

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values = rowValues(row);
    row.height = 23;

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      styleCellBase(row.getCell(columnNumber));
    }

    if (isBlankRow(values)) {
      row.height = 9;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.border = {};
        cell.fill = solidFill("FFFFFFFF");
      }
      continue;
    }

    if (isTitleRow(values)) {
      row.height = 40;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 18,
          bold: true,
          color: { argb: "FFFFFFFF" }
        };
        cell.fill = solidFill("FF1F4E78");
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: true
        };
        cell.border = thinBorder("FF1F4E78");
      }
      continue;
    }

    if (isCompanyRow(values)) {
      row.height = 28;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 11,
          bold: true,
          color: { argb: "FF1F4E78" }
        };
        cell.fill = solidFill("FFD9EAF7");
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }
      continue;
    }

    if (isGroupRow(values)) {
      row.height = 28;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 11,
          bold: true,
          color: { argb: "FF1F2937" }
        };
        cell.fill = solidFill("FFEAF2F8");
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }
      continue;
    }

    if (isHeaderRow(values)) {
      if (!firstHeaderRow) firstHeaderRow = rowNumber;
      row.height = 38;

      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FFFFFFFF" }
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: true
        };
        cell.border = thinBorder("FF9EB6CE");

        if (columnNumber <= 2) {
          cell.fill = solidFill("FF2F5597");
        } else if (columnNumber >= dateStartColumn && columnNumber <= dateEndColumn) {
          const day = columnNumber - dateStartColumn + 1;
          cell.fill = solidFill(weekendDays.has(day) ? "FFE6A23C" : "FF4472C4");
        } else if (columnNumber >= summaryStartColumn && columnNumber <= summaryEndColumn) {
          cell.fill = solidFill("FF305496");
        } else {
          cell.fill = solidFill("FF4472C4");
        }
      }
      continue;
    }

    const employeeName = text(row.getCell(2).value);
    const pendingName = employeeName.startsWith("待确认姓名");
    const alternatingFill = rowNumber % 2 === 0 ? "FFF7FAFD" : "FFFFFFFF";

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const cell = row.getCell(columnNumber);
      cell.fill = solidFill(pendingName ? "FFFFF2CC" : alternatingFill);

      if (columnNumber === 1) {
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FF44546A" }
        };
      }

      if (columnNumber === 2) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: pendingName ? "FF9C5700" : "FF1F2937" }
        };
        if (pendingName) cell.fill = solidFill("FFFFC000");
      }

      if (columnNumber >= dateStartColumn && columnNumber <= dateEndColumn) {
        const day = columnNumber - dateStartColumn + 1;
        if (weekendDays.has(day) && !pendingName) {
          cell.fill = solidFill("FFFFF7E6");
        }
      }

      if (columnNumber >= summaryStartColumn && columnNumber <= summaryEndColumn) {
        if (!pendingName) cell.fill = solidFill("FFF0F6FC");

        if (columnNumber === 34 || columnNumber === 35) {
          cell.font = {
            name: "Microsoft YaHei",
            size: 10,
            bold: true,
            color: { argb: "FF1F4E78" }
          };
        }
      }

      if (columnNumber === summaryEndColumn) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }
    }

    row.getCell(35).numFmt = "0.##";
  }

  const frozenRow = firstHeaderRow || 1;
  worksheet.views = [
    {
      state: "frozen",
      xSplit: 2,
      ySplit: frozenRow,
      topLeftCell: `C${frozenRow + 1}`,
      activeCell: `C${frozenRow + 1}`
    }
  ];

  if (worksheet.rowCount > 0 && worksheet.columnCount > 0) {
    const lastColumn = worksheet.getColumn(worksheet.columnCount).letter;
    worksheet.pageSetup.printArea = `A1:${lastColumn}${worksheet.rowCount}`;
  }
}

function styleErrorSheet(worksheet) {
  setPageLayout(worksheet);

  worksheet.properties = {
    ...worksheet.properties,
    defaultRowHeight: 22,
    tabColor: { argb: "FFC65911" }
  };

  const widths = [8, 18, 18, 12, 20, 16, 20, 14, 8, 20, 32, 12, 13, 24, 18, 26];
  widths.forEach((width, index) => {
    worksheet.getColumn(index + 1).width = width;
  });

  let headerRowNumber = 0;

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values = rowValues(row);

    if (isErrorTitleRow(values) || isErrorInstructionRow(values)) {
      mergeWholeRow(worksheet, rowNumber);
    }
  }

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values = rowValues(row);
    row.height = 24;

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      styleCellBase(row.getCell(columnNumber), "FFE3C5B5");
    }

    if (isBlankRow(values)) {
      row.height = 9;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.border = {};
        cell.fill = solidFill("FFFFFFFF");
      }
      continue;
    }

    if (isErrorTitleRow(values)) {
      row.height = 40;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 18,
          bold: true,
          color: { argb: "FFFFFFFF" }
        };
        cell.fill = solidFill("FFC65911");
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: true
        };
        cell.border = thinBorder("FFC65911");
      }
      continue;
    }

    if (isErrorInstructionRow(values)) {
      row.height = 34;
      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FF9C5700" }
        };
        cell.fill = solidFill("FFFFF2CC");
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }
      continue;
    }

    if (isErrorHeaderRow(values)) {
      headerRowNumber = rowNumber;
      row.height = 40;

      for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FFFFFFFF" }
        };
        cell.fill = solidFill(columnNumber >= 15 ? "FFE6A23C" : "FFC65911");
        cell.border = thinBorder("FFD99A6C");
      }
      continue;
    }

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const cell = row.getCell(columnNumber);
      cell.fill = solidFill(rowNumber % 2 === 0 ? "FFFFF8F4" : "FFFFFFFF");

      if (columnNumber === 10 || columnNumber === 11 || columnNumber === 14) {
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }

      if (columnNumber === 15 || columnNumber === 16) {
        cell.fill = solidFill("FFFFF2CC");
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FF9C5700" }
        };
        cell.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
          indent: 1
        };
      }

      if (columnNumber === 12 && text(cell.value).toLowerCase() === "pending") {
        cell.fill = solidFill("FFFCE4D6");
        cell.font = {
          name: "Microsoft YaHei",
          size: 10,
          bold: true,
          color: { argb: "FFC00000" }
        };
      }
    }
  }

  const frozenRow = headerRowNumber || 4;
  worksheet.views = [
    {
      state: "frozen",
      xSplit: 1,
      ySplit: frozenRow,
      topLeftCell: `B${frozenRow + 1}`,
      activeCell: `B${frozenRow + 1}`
    }
  ];

  if (headerRowNumber > 0 && worksheet.rowCount > headerRowNumber) {
    worksheet.autoFilter = {
      from: { row: headerRowNumber, column: 1 },
      to: { row: worksheet.rowCount, column: 16 }
    };
  }

  if (worksheet.rowCount > 0 && worksheet.columnCount > 0) {
    const lastColumn = worksheet.getColumn(worksheet.columnCount).letter;
    worksheet.pageSetup.printArea = `A1:${lastColumn}${worksheet.rowCount}`;
  }
}

function addRowsToSheet(worksheet, rows) {
  rows.forEach((row) => {
    worksheet.addRow(Array.isArray(row) ? row : [String(row || "")]);
  });
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

    const monthValue = text(excelData.month);
    const sheets = Array.isArray(excelData.sheets) ? excelData.sheets : [];

    if (sheets.length === 0) {
      return res.status(400).json({
        success: false,
        message: "excel_data.sheets为空，无法生成Excel"
      });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Coze Attendance Bot";
    workbook.lastModifiedBy = "Coze Attendance Bot";
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.title = fileName.replace(/\.xlsx$/i, "");
    workbook.subject = "临时工工时待核对表";

    const usedSheetNames = new Set();

    for (const sheetData of sheets) {
      const sheetName = safeSheetName(sheetData.sheet_name, usedSheetNames);
      const rows = Array.isArray(sheetData.rows) ? sheetData.rows : [];
      const worksheet = workbook.addWorksheet(sheetName);

      addRowsToSheet(worksheet, rows);

      const errorSheet =
        sheetName === "异常待确认" ||
        rows.some((row) => isErrorTitleRow(Array.isArray(row) ? row : [row]));

      if (errorSheet) {
        styleErrorSheet(worksheet);
      } else {
        styleAttendanceSheet(worksheet, monthValue);
      }
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
      error: error?.message || String(error)
    });
  }
}
