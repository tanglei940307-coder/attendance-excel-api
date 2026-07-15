import ExcelJS from "exceljs";
import { put } from "@vercel/blob";

/**
 * 尝试将字符串或对象转为普通对象。
 */
function parseMaybeJson(value) {
  if (!value) return {};

  if (
    typeof value === "object" &&
    value !== null
  ) {
    return value;
  }

  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

/**
 * 清理下载文件名。
 */
function safeFileName(name) {
  const raw = String(
    name || "临时工工时待核对表.xlsx"
  ).trim();

  const cleaned = raw.replace(
    /[\\/:*?"<>|]/g,
    "_"
  );

  return cleaned.endsWith(".xlsx")
    ? cleaned
    : `${cleaned}.xlsx`;
}

/**
 * 清理工作表名称，并保证不重复。
 */
function safeSheetName(name, usedNames) {
  let sheetName = String(
    name || "Sheet"
  ).trim();

  sheetName = sheetName
    .replace(/[\\/?*[\]:]/g, "")
    .slice(0, 31) || "Sheet";

  let finalName = sheetName;
  let index = 1;

  while (usedNames.has(finalName)) {
    const suffix = `_${index}`;

    finalName =
      sheetName.slice(
        0,
        31 - suffix.length
      ) + suffix;

    index += 1;
  }

  usedNames.add(finalName);

  return finalName;
}

function text(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value).trim();
}

function isBlankRow(values) {
  return (
    Array.isArray(values) &&
    values.every((value) => text(value) === "")
  );
}

function isHeaderRow(values) {
  return (
    Array.isArray(values) &&
    values.includes("序号") &&
    values.includes("姓名")
  );
}

function isGroupRow(values) {
  return (
    Array.isArray(values) &&
    values.length >= 1 &&
    text(values[0]).startsWith("工作组：")
  );
}

function isCompanyRow(values) {
  return (
    Array.isArray(values) &&
    values.length >= 1 &&
    text(values[0]).startsWith("外包公司：")
  );
}

function isTitleRow(values) {
  return (
    Array.isArray(values) &&
    values.length >= 1 &&
    text(values[0]).includes(
      "临时工工时待核对表"
    )
  );
}

function isErrorTitleRow(values) {
  return (
    Array.isArray(values) &&
    text(values[0]) === "异常待确认"
  );
}

function isErrorInstructionRow(values) {
  return (
    Array.isArray(values) &&
    text(values[0]).includes(
      "修正值"
    ) &&
    text(values[0]).includes(
      "修正备注"
    )
  );
}

function isErrorHeaderRow(values) {
  return (
    Array.isArray(values) &&
    values.includes("异常ID") &&
    values.includes("异常原因")
  );
}

/**
 * 从 YYYY-MM 中获得周末日期。
 */
function getWeekendDays(monthValue) {
  const result = new Set();

  const match = text(monthValue).match(
    /^(\d{4})-(\d{1,2})$/
  );

  if (!match) {
    return result;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);

  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    return result;
  }

  const daysInMonth = new Date(
    year,
    month,
    0
  ).getDate();

  for (
    let day = 1;
    day <= daysInMonth;
    day++
  ) {
    const weekDay = new Date(
      year,
      month - 1,
      day
    ).getDay();

    if (
      weekDay === 0 ||
      weekDay === 6
    ) {
      result.add(day);
    }
  }

  return result;
}

/**
 * 统一细边框。
 */
function thinBorder(color = "FFD6DEE8") {
  return {
    top: {
      style: "thin",
      color: { argb: color }
    },
    left: {
      style: "thin",
      color: { argb: color }
    },
    bottom: {
      style: "thin",
      color: { argb: color }
    },
    right: {
      style: "thin",
      color: { argb: color }
    }
  };
}

function solidFill(argb) {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb }
  };
}

/**
 * 将一行合并到工作表最大列。
 */
function mergeWholeRow(ws, rowNumber) {
  if (
    ws.columnCount <= 1 ||
    rowNumber < 1
  ) {
    return;
  }

  try {
    ws.mergeCells(
      rowNumber,
      1,
      rowNumber,
      ws.columnCount
    );
  } catch {
    // 合并失败不影响文件生成
  }
}

/**
 * 设置工作簿通用打印属性。
 */
function setPageLayout(ws) {
  ws.pageSetup = {
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

  ws.headerFooter = {
    oddHeader:
      "&C&14&B临时工考勤核对表",
    oddFooter:
      "&L生成自考勤识别系统&C第 &P 页 / 共 &N 页&R&F"
  };

  ws.properties.defaultRowHeight = 22;
}

/**
 * 主考勤工作表样式。
 */
function styleAttendanceSheet(
  ws,
  monthValue
) {
  const weekendDays =
    getWeekendDays(monthValue);

  let firstHeaderRow = 0;

  setPageLayout(ws);

  ws.properties.tabColor = {
    argb: "FF2F5597"
  };

  ws.sheetProperties.pageSetUpPr = {
    fitToPage: true
  };

  /*
   * 日期列：
   * A序号、B姓名、C开始为1日。
   */
  const dateStartColumn = 3;
  const dateEndColumn = 33;

  /*
   * 汇总列：
   * AH至AM。
   */
  const summaryStartColumn = 34;
  const summaryEndColumn = 39;

  ws.eachRow(
    { includeEmpty: true },
    (row, rowNumber) => {
      const values = row.values.slice(1);

      row.height = 23;

      /*
       * 通用单元格样式。
       */
      row.eachCell(
        { includeEmpty: true },
        (cell) => {
          cell.font = {
            name: "Microsoft YaHei",
            size: 10,
            color: {
              argb: "FF1F2937"
            }
          };

          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true
          };

          cell.border = thinBorder();
        }
      );

      if (isBlankRow(values)) {
        row.height = 9;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.border = {};
            cell.fill = solidFill(
              "FFFFFFFF"
            );
          }
        );

        return;
      }

      /*
       * 主标题行。
       */
      if (isTitleRow(values)) {
        row.height = 38;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 18,
              bold: true,
              color: {
                argb: "FFFFFFFF"
              }
            };

            cell.fill = solidFill(
              "FF1F4E78"
            );

            cell.alignment = {
              vertical: "middle",
              horizontal: "center",
              wrapText: true
            };

            cell.border = thinBorder(
              "FF1F4E78"
            );
          }
        );

        return;
      }

      /*
       * 外包公司信息行。
       */
      if (isCompanyRow(values)) {
        row.height = 27;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 11,
              bold: true,
              color: {
                argb: "FF1F4E78"
              }
            };

            cell.fill = solidFill(
              "FFD9EAF7"
            );

            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
              indent: 1
            };
          }
        );

        return;
      }

      /*
       * 工作组标题行。
       */
      if (isGroupRow(values)) {
        row.height = 27;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 11,
              bold: true,
              color: {
                argb: "FF1F2937"
              }
            };

            cell.fill = solidFill(
              "FFEAF2F8"
            );

            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
              indent: 1
            };
          }
        );

        return;
      }

      /*
       * 表头行。
       */
      if (isHeaderRow(values)) {
        if (!firstHeaderRow) {
          firstHeaderRow = rowNumber;
        }

        row.height = 36;

        row.eachCell(
          { includeEmpty: true },
          (cell, columnNumber) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FFFFFFFF"
              }
            };

            cell.alignment = {
              vertical: "middle",
              horizontal: "center",
              wrapText: true
            };

            cell.border = thinBorder(
              "FF9EB6CE"
            );

            /*
             * 序号和姓名。
             */
            if (columnNumber <= 2) {
              cell.fill = solidFill(
                "FF2F5597"
              );

              return;
            }

            /*
             * 日期列。
             */
            if (
              columnNumber >=
                dateStartColumn &&
              columnNumber <=
                dateEndColumn
            ) {
              const day =
                columnNumber -
                dateStartColumn +
                1;

              if (
                weekendDays.has(day)
              ) {
                cell.fill = solidFill(
                  "FFE6A23C"
                );
              } else {
                cell.fill = solidFill(
                  "FF4472C4"
                );
              }

              return;
            }

            /*
             * 汇总列。
             */
            if (
              columnNumber >=
                summaryStartColumn &&
              columnNumber <=
                summaryEndColumn
            ) {
              cell.fill = solidFill(
                "FF305496"
              );

              return;
            }

            cell.fill = solidFill(
              "FF4472C4"
            );
          }
        );

        return;
      }

      /*
       * 普通员工数据行。
       */
      const employeeName =
        text(row.getCell(2).value);

      const isPendingName =
        employeeName.startsWith(
          "待确认姓名"
        );

      const dataFill =
        rowNumber % 2 === 0
          ? "FFF7FAFD"
          : "FFFFFFFF";

      row.eachCell(
        { includeEmpty: true },
        (cell, columnNumber) => {
          cell.fill = solidFill(
            isPendingName
              ? "FFFFF2CC"
              : dataFill
          );

          /*
           * 序号。
           */
          if (columnNumber === 1) {
            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FF44546A"
              }
            };
          }

          /*
           * 姓名。
           */
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
              color: {
                argb: isPendingName
                  ? "FF9C5700"
                  : "FF1F2937"
              }
            };

            if (isPendingName) {
              cell.fill = solidFill(
                "FFFFC000"
              );
            }
          }

          /*
           * 周末日期列。
           */
          if (
            columnNumber >=
              dateStartColumn &&
            columnNumber <=
              dateEndColumn
          ) {
            const day =
              columnNumber -
              dateStartColumn +
              1;

            if (
              weekendDays.has(day) &&
              !isPendingName
            ) {
              cell.fill = solidFill(
                "FFFFF7E6"
              );
            }
          }

          /*
           * 汇总列。
           */
          if (
            columnNumber >=
              summaryStartColumn &&
            columnNumber <=
              summaryEndColumn
          ) {
            if (!isPendingName) {
              cell.fill = solidFill(
                "FFF0F6FC"
              );
            }

            if (
              columnNumber === 34 ||
              columnNumber === 35
            ) {
              cell.font = {
                name:
                  "Microsoft YaHei",
                size: 10,
                bold: true,
                color: {
                  argb: "FF1F4E78"
                }
              };
            }
          }

          /*
           * 备注靠左。
           */
          if (
            columnNumber ===
            summaryEndColumn
          ) {
            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
              indent: 1
            };
          }
        }
      );

      /*
       * 总工时显示为最多两位小数。
       */
      row.getCell(35).numFmt = "0.##";
    }
  );

  /*
   * 合并标题、公司和工作组行。
   */
  ws.eachRow((row, rowNumber) => {
    const values = row.values.slice(1);

    if (
      isTitleRow(values) ||
      isCompanyRow(values) ||
      isGroupRow(values)
    ) {
      mergeWholeRow(ws, rowNumber);
    }
  });

  /*
   * 冻结首个表头上方内容，同时固定序号和姓名列。
   */
  ws.views = [
    {
      state: "frozen",
      xSplit: 2,
      ySplit:
        firstHeaderRow > 0
          ? firstHeaderRow
          : 1,
      topLeftCell:
        firstHeaderRow > 0
          ? `C${firstHeaderRow + 1}`
          : "C2",
      activeCell:
        firstHeaderRow > 0
          ? `C${firstHeaderRow + 1}`
          : "C2"
    }
  ];

  /*
   * 设置列宽。
   */
  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 24;

  for (
    let columnNumber = 3;
    columnNumber <= 33;
    columnNumber++
  ) {
    ws.getColumn(
      columnNumber
    ).width = 5.5;
  }

  ws.getColumn(34).width = 11;
  ws.getColumn(35).width = 11;
  ws.getColumn(36).width = 8;
  ws.getColumn(37).width = 8;
  ws.getColumn(38).width = 8;
  ws.getColumn(39).width = 24;

  /*
   * 设置打印区域。
   */
  if (
    ws.rowCount > 0 &&
    ws.columnCount > 0
  ) {
    ws.pageSetup.printArea =
      `A1:${ws.getColumn(
        ws.columnCount
      ).letter}${ws.rowCount}`;
  }
}

/**
 * 异常待确认工作表样式。
 */
function styleErrorSheet(ws) {
  setPageLayout(ws);

  ws.properties.tabColor = {
    argb: "FFC65911"
  };

  let headerRowNumber = 0;

  ws.eachRow(
    { includeEmpty: true },
    (row, rowNumber) => {
      const values = row.values.slice(1);

      row.height = 24;

      row.eachCell(
        { includeEmpty: true },
        (cell) => {
          cell.font = {
            name: "Microsoft YaHei",
            size: 10,
            color: {
              argb: "FF1F2937"
            }
          };

          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true
          };

          cell.border = thinBorder(
            "FFE3C5B5"
          );
        }
      );

      if (isBlankRow(values)) {
        row.height = 9;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.border = {};
          }
        );

        return;
      }

      if (isErrorTitleRow(values)) {
        row.height = 38;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 18,
              bold: true,
              color: {
                argb: "FFFFFFFF"
              }
            };

            cell.fill = solidFill(
              "FFC65911"
            );

            cell.alignment = {
              vertical: "middle",
              horizontal: "center",
              wrapText: true
            };
          }
        );

        return;
      }

      if (
        isErrorInstructionRow(values)
      ) {
        row.height = 32;

        row.eachCell(
          { includeEmpty: true },
          (cell) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FF9C5700"
              }
            };

            cell.fill = solidFill(
              "FFFFF2CC"
            );

            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
              indent: 1
            };
          }
        );

        return;
      }

      if (isErrorHeaderRow(values)) {
        headerRowNumber = rowNumber;
        row.height = 38;

        row.eachCell(
          { includeEmpty: true },
          (cell, columnNumber) => {
            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FFFFFFFF"
              }
            };

            cell.fill = solidFill(
              columnNumber >= 15
                ? "FFE6A23C"
                : "FFC65911"
            );

            cell.border = thinBorder(
              "FFD99A6C"
            );
          }
        );

        return;
      }

      /*
       * 异常数据行。
       */
      row.eachCell(
        { includeEmpty: true },
        (cell, columnNumber) => {
          cell.fill = solidFill(
            rowNumber % 2 === 0
              ? "FFFFF8F4"
              : "FFFFFFFF"
          );

          /*
           * 修正值、修正备注。
           */
          if (
            columnNumber === 15 ||
            columnNumber === 16
          ) {
            cell.fill = solidFill(
              "FFFFF2CC"
            );

            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FF9C5700"
              }
            };

            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true
            };
          }

          /*
           * 异常原因靠左。
           */
          if (
            columnNumber === 10 ||
            columnNumber === 11 ||
            columnNumber === 14
          ) {
            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
              indent: 1
            };
          }

          /*
           * pending状态高亮。
           */
          if (
            columnNumber === 12 &&
            text(cell.value).toLowerCase() ===
              "pending"
          ) {
            cell.fill = solidFill(
              "FFFCE4D6"
            );

            cell.font = {
              name: "Microsoft YaHei",
              size: 10,
              bold: true,
              color: {
                argb: "FFC00000"
              }
            };
          }
        }
      );
    }
  );

  /*
   * 合并异常标题及说明行。
   */
  ws.eachRow((row, rowNumber) => {
    const values = row.values.slice(1);

    if (
      isErrorTitleRow(values) ||
      isErrorInstructionRow(values)
    ) {
      mergeWholeRow(ws, rowNumber);
    }
  });

  /*
   * 异常表冻结表头。
   */
  ws.views = [
    {
      state: "frozen",
      xSplit: 1,
      ySplit:
        headerRowNumber > 0
          ? headerRowNumber
          : 4,
      topLeftCell:
        headerRowNumber > 0
          ? `B${headerRowNumber + 1}`
          : "B5",
      activeCell:
        headerRowNumber > 0
          ? `B${headerRowNumber + 1}`
          : "B5"
    }
  ];

  /*
   * 异常表列宽。
   */
  const widths = [
    8,
    18,
    18,
    12,
    20,
    16,
    20,
    14,
    8,
    20,
    32,
    12,
    13,
    24,
    18,
    26
  ];

  widths.forEach(
    (width, index) => {
      ws.getColumn(
        index + 1
      ).width = width;
    }
  );

  /*
   * 异常表启用筛选。
   */
  if (
    headerRowNumber > 0 &&
    ws.rowCount > headerRowNumber
  ) {
    ws.autoFilter = {
      from: {
        row: headerRowNumber,
        column: 1
      },
      to: {
        row: ws.rowCount,
        column: 16
      }
    };
  }

  if (
    ws.rowCount > 0 &&
    ws.columnCount > 0
  ) {
    ws.pageSetup.printArea =
      `A1:${ws.getColumn(
        ws.columnCount
      ).letter}${ws.rowCount}`;
  }
}

/**
 * 将二维数组写入工作表。
 */
function addRowsToSheet(ws, rows) {
  rows.forEach((row) => {
    ws.addRow(
      Array.isArray(row)
        ? row
        : [String(row || "")]
    );
  });
}

/**
 * 创建Excel文件接口。
 */
export default async function handler(
  req,
  res
) {
  try {
    /*
     * 跨域预检。
     */
    if (req.method === "OPTIONS") {
      res.setHeader(
        "Access-Control-Allow-Origin",
        "*"
      );

      res.setHeader(
        "Access-Control-Allow-Methods",
        "POST, OPTIONS"
      );

      res.setHeader(
        "Access-Control-Allow-Headers",
        "Content-Type, x-api-key"
      );

      return res.status(204).end();
    }

    if (req.method !== "POST") {
      return res.status(405).json({
        success: false,
        message:
          "Only POST is allowed"
      });
    }

    /*
     * API密钥校验。
     */
    const apiKey =
      process.env.API_KEY || "";

    if (apiKey) {
      const requestKey =
        req.headers["x-api-key"];

      if (requestKey !== apiKey) {
        return res.status(401).json({
          success: false,
          message: "Unauthorized"
        });
      }
    }

    const body = parseMaybeJson(
      req.body
    );

    const excelData = parseMaybeJson(
      body.excel_data
    );

    const fileName = safeFileName(
      body.file_name ||
      excelData.file_name ||
      "临时工工时待核对表.xlsx"
    );

    const monthValue =
      text(excelData.month);

    const sheets =
      Array.isArray(
        excelData.sheets
      )
        ? excelData.sheets
        : [];

    if (sheets.length === 0) {
      return res.status(400).json({
        success: false,
        message:
          "excel_data.sheets为空，无法生成Excel"
      });
    }

    /*
     * 创建工作簿。
     */
    const workbook =
      new ExcelJS.Workbook();

    workbook.creator =
      "Coze Attendance Bot";

    workbook.lastModifiedBy =
      "Coze Attendance Bot";

    workbook.created =
      new Date();

    workbook.modified =
      new Date();

    workbook.company =
      "Attendance Management";

    workbook.subject =
      "临时工工时待核对表";

    workbook.title =
      fileName.replace(
        /\.xlsx$/i,
        ""
      );

    /*
     * 日期格式使用1900日期系统。
     */
    workbook.properties.date1904 =
      false;

    const usedSheetNames =
      new Set();

    for (const sheetData of sheets) {
      const sheetName =
        safeSheetName(
          sheetData.sheet_name,
          usedSheetNames
        );

      const rows =
        Array.isArray(
          sheetData.rows
        )
          ? sheetData.rows
          : [];

      const ws =
        workbook.addWorksheet(
          sheetName,
          {
            properties: {
              defaultRowHeight: 22
            }
          }
        );

      addRowsToSheet(
        ws,
        rows
      );

      if (
        sheetName ===
          "异常待确认" ||
        rows.some((row) =>
          isErrorTitleRow(
            Array.isArray(row)
              ? row
              : [row]
          )
        )
      ) {
        styleErrorSheet(ws);
      } else {
        styleAttendanceSheet(
          ws,
          monthValue
        );
      }
    }

    /*
     * 生成Excel缓冲区。
     */
    const buffer =
      await workbook.xlsx.writeBuffer();

    const timestamp =
      Date.now();

    const blobName =
      `attendance/${timestamp}_${fileName}`;

    const blob = await put(
      blobName,
      buffer,
      {
        access: "public",
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    );

    res.setHeader(
      "Access-Control-Allow-Origin",
      "*"
    );

    return res.status(200).json({
      success: true,
      file_name: fileName,
      file_url: blob.url,
      message:
        "Excel生成成功"
    });
  } catch (error) {
    console.error(error);

    return res.status(500).json({
      success: false,
      message:
        "Excel生成失败",
      error:
        error?.message ||
        String(error)
    });
  }
}
