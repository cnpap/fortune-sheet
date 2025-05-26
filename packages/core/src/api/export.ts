import { Sheet } from "../types";

/**
 * 将Sheet数据转换为二维数组
 * @param sheet Sheet对象
 * @returns 二维数组表示的数据
 */
export function sheetToArray(sheet: Sheet): any[][] {
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
 * @param data 二维数组数据
 * @returns CSV格式的字符串
 */
export function arrayToCSV(data: any[][]): string {
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
 * 导出为CSV格式
 * @param sheets 工作表数组
 * @returns CSV格式的字符串 (只导出第一个工作表)
 */
export function exportCSV(sheets: Sheet[]): string {
  if (sheets.length === 0) return "";
  const sheet = sheets[0]; // 只导出第一个工作表
  const data = sheetToArray(sheet);
  return arrayToCSV(data);
}

/**
 * 获取支持导出到Excel的工具函数
 * 注意：这些函数需要在前端环境才能执行，因为它们依赖于xlsx库
 */
export type ExportExcelUtils = {
  /**
   * 导出为Excel格式
   * @param sheets 工作表数组
   * @param options 导出选项
   */
  exportExcel: (
    sheets: Sheet[],
    options?: {
      type?: "xlsx" | "csv";
      fileName?: string;
    }
  ) => void;
};
