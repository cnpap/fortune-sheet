import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { Sheet } from "@fortune-sheet/core";

/**
 * 将Sheet数据转换为二维数组
 */
function sheetToArray(sheet: Sheet): any[][] {
  const celldata = sheet.celldata || [];
  const rows = sheet.row || 100;
  const cols = sheet.column || 26;

  // 创建一个二维数组来存储单元格数据
  const sheetData: any[][] = Array(rows)
    .fill(null)
    .map(() => Array(cols).fill(null));

  // 填充单元格数据
  celldata.forEach((cell) => {
    if (cell.r < rows && cell.c < cols) {
      sheetData[cell.r][cell.c] = cell.v?.v !== undefined ? cell.v.v : null;
    }
  });

  return sheetData;
}

/**
 * 将表格数据转换为CSV格式
 */
function arrayToCSV(data: any[][]): string {
  return data
    .map((row) =>
      row
        .map((cell) => {
          if (cell === null || cell === undefined) return "";
          // 处理包含逗号或引号的单元格
          const cellStr = String(cell);
          if (
            cellStr.includes(",") ||
            cellStr.includes('"') ||
            cellStr.includes("\n")
          ) {
            return `"${cellStr.replace(/"/g, '""')}"`;
          }
          return cellStr;
        })
        .join(",")
    )
    .join("\n");
}

/**
 * 导出为CSV格式 (只导出第一个工作表)
 */
function exportCSV(sheets: Sheet[]): string {
  if (sheets.length === 0) return "";
  const sheet = sheets[0]; // 只导出第一个工作表
  const data = sheetToArray(sheet);
  return arrayToCSV(data);
}

/**
 * 将工作表导出为Excel或CSV文件
 * @param sheets 工作表数组
 * @param options 导出选项
 */
export const exportExcel = (
  sheets: Sheet[],
  options?: {
    type?: "xlsx" | "csv";
    fileName?: string;
  }
) => {
  const type = options?.type || "xlsx";
  const fileName = options?.fileName || "fortune-sheet-export";

  if (type === "csv") {
    // CSV导出 (仅导出第一个工作表)
    const csvContent = exportCSV(sheets);
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
    saveAs(blob, `${fileName}.csv`);
    return;
  }

  // Excel导出
  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    // 使用核心库中的函数将表格数据转换为二维数组
    const sheetData = sheetToArray(sheet);

    // 创建工作表
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  });

  // 导出为Excel文件
  const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  saveAs(blob, `${fileName}.xlsx`);
};

/**
 * 扩展WorkbookInstance的类型，添加导出功能
 */
export type ExportAPI = {
  /**
   * 导出工作簿为Excel或CSV
   */
  exportWorkbook: (options?: {
    type?: "xlsx" | "csv";
    fileName?: string;
  }) => void;
};
