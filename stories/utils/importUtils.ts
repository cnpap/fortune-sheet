import * as XLSX from "xlsx";
import { Sheet } from "@fortune-sheet/core";

/**
 * 将Excel或CSV文件导入为fortune-sheet可用的数据格式
 * @param file 上传的文件对象
 * @returns Promise<Sheet[]> 转换后的sheet数据数组
 */
export async function importExcel(file: File): Promise<Sheet[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        if (!e.target?.result) {
          reject(new Error("读取文件失败"));
          return;
        }

        // 读取工作簿
        const data = new Uint8Array(e.target.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });

        // 转换为fortune-sheet格式
        const sheets: Sheet[] = [];

        workbook.SheetNames.forEach((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];

          // 获取工作表范围
          const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");
          const rowCount = range.e.r + 1;
          const colCount = range.e.c + 1;

          // 创建celldata
          const celldata: Array<{
            r: number;
            c: number;
            v: { v: any; m?: string };
          }> = [];

          // 遍历单元格
          for (let r = 0; r <= range.e.r; r += 1) {
            for (let c = 0; c <= range.e.c; c += 1) {
              const cellAddress = XLSX.utils.encode_cell({ r, c });
              const cell = worksheet[cellAddress];

              if (cell && cell.v !== undefined && cell.v !== null) {
                // 处理不同类型的单元格值
                let cellValue: any = cell.v;

                // 确保数值类型正确
                if (cell.t === "n") {
                  cellValue = Number(cellValue);
                } else if (cell.t === "b") {
                  cellValue = Boolean(cellValue);
                } else {
                  // 确保字符串类型
                  cellValue = String(cellValue);
                }

                celldata.push({
                  r,
                  c,
                  v: {
                    v: cellValue,
                    m: cell.w || String(cellValue), // 包含格式化显示值，确保有值
                  },
                });
              }
            }
          }

          // 生成唯一的sheet ID
          const sheetId = `sheet_${Date.now()}_${index}_${Math.floor(
            Math.random() * 1000
          )}`;

          // 创建Sheet对象
          sheets.push({
            name: sheetName,
            celldata,
            order: index,
            row: Math.max(50, rowCount), // 至少50行
            column: Math.max(26, colCount), // 至少26列
            config: {},
            id: sheetId,
          });
        });

        resolve(sheets);
      } catch (error) {
        console.error("解析Excel文件失败:", error);
        reject(error);
      }
    };

    reader.onerror = (error) => {
      console.error("读取文件出错:", error);
      reject(error);
    };

    // 读取文件内容
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 导入CSV文件
 * @param file 上传的CSV文件
 * @returns Promise<Sheet> 转换后的sheet数据
 */
export async function importCSV(file: File): Promise<Sheet> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        if (!e.target?.result) {
          reject(new Error("读取文件失败"));
          return;
        }

        // 读取CSV内容
        const content = e.target.result as string;

        // 使用xlsx解析CSV
        const workbook = XLSX.read(content, { type: "string" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // 获取工作表范围
        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1");
        const rowCount = range.e.r + 1;
        const colCount = range.e.c + 1;

        // 创建celldata
        const celldata: Array<{
          r: number;
          c: number;
          v: { v: any; m?: string };
        }> = [];

        // 遍历单元格
        for (let r = 0; r <= range.e.r; r += 1) {
          for (let c = 0; c <= range.e.c; c += 1) {
            const cellAddress = XLSX.utils.encode_cell({ r, c });
            const cell = worksheet[cellAddress];

            if (cell && cell.v !== undefined && cell.v !== null) {
              // 处理不同类型的单元格值
              let cellValue: any = cell.v;

              // 确保数值类型正确
              if (cell.t === "n") {
                cellValue = Number(cellValue);
              } else if (cell.t === "b") {
                cellValue = Boolean(cellValue);
              } else {
                // 确保字符串类型
                cellValue = String(cellValue);
              }

              celldata.push({
                r,
                c,
                v: {
                  v: cellValue,
                  m: cell.w || String(cellValue), // 包含格式化显示值
                },
              });
            }
          }
        }

        // 生成唯一的sheet ID
        const sheetId = `sheet_${Date.now()}_${Math.floor(
          Math.random() * 1000
        )}`;

        // 创建Sheet对象
        const sheet: Sheet = {
          name: file.name.replace(/\.[^/.]+$/, "") || "导入的数据",
          celldata,
          order: 0,
          row: Math.max(50, rowCount),
          column: Math.max(26, colCount),
          config: {},
          id: sheetId,
        };

        resolve(sheet);
      } catch (error) {
        console.error("解析CSV文件失败:", error);
        reject(error);
      }
    };

    reader.onerror = (error) => {
      console.error("读取文件出错:", error);
      reject(error);
    };

    // 读取文件内容
    reader.readAsText(file);
  });
}

/**
 * 检测文件类型
 * @param file 文件对象
 */
export function detectFileType(file: File): "xlsx" | "csv" | "unknown" {
  const fileName = file.name.toLowerCase();

  if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
    return "xlsx";
  }

  if (fileName.endsWith(".csv")) {
    return "csv";
  }

  return "unknown";
}

/**
 * 自动检测文件类型并导入
 * @param file 上传的文件
 */
export async function importFile(file: File): Promise<Sheet[]> {
  const fileType = detectFileType(file);

  if (fileType === "xlsx") {
    return importExcel(file);
  }

  if (fileType === "csv") {
    const sheet = await importCSV(file);
    return [sheet];
  }

  throw new Error(`不支持的文件类型: ${file.type}`);
}
