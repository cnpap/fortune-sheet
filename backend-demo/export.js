// const _ = require("lodash");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

/**
 * 将sheet数据转换为二维数组
 * @param {Object} sheet Sheet对象
 * @returns {Array} 二维数组
 */
function sheetToArray(sheet) {
  const celldata = sheet.celldata || [];
  const rows = sheet.row || 100;
  const cols = sheet.column || 26;

  // 创建一个二维数组来存储单元格数据
  const sheetData = Array(rows)
    .fill(null)
    .map(() => Array(cols).fill(null));

  // 填充单元格数据
  celldata.forEach((cell) => {
    if (cell.r < rows && cell.c < cols) {
      if (cell.v && typeof cell.v === "object" && "v" in cell.v) {
        sheetData[cell.r][cell.c] = cell.v.v;
      } else {
        sheetData[cell.r][cell.c] = cell.v;
      }
    }
  });

  return sheetData;
}

/**
 * 将数据导出为Excel文件
 * @param {Array} sheets 工作表数组
 * @param {string} fileName 文件名
 * @param {string} format 导出格式 (xlsx 或 csv)
 * @returns {Object} { filePath, fileName } 文件路径和名称
 */
function exportToExcel(
  sheets,
  fileName = "fortune-sheet-export",
  format = "xlsx"
) {
  // 创建临时目录
  const exportDir = path.join(__dirname, "exports");
  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir);
  }

  // 创建工作簿
  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    // 转换数据为二维数组
    const data = sheetToArray(sheet);

    // 创建工作表
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name || "Sheet");
  });

  // 生成文件路径
  const timestamp = Date.now();
  const fileNameWithTimestamp = `${fileName}-${timestamp}`;
  const filePath = path.join(exportDir, `${fileNameWithTimestamp}.${format}`);

  // 根据格式导出
  if (format === "csv") {
    // 只导出第一个工作表为CSV
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    fs.writeFileSync(filePath, csv);
  } else {
    // 导出为Excel文件
    XLSX.writeFile(workbook, filePath);
  }

  return {
    filePath,
    fileName: `${fileNameWithTimestamp}.${format}`,
  };
}

/**
 * 删除指定天数前创建的导出文件
 * @param {number} days 天数
 */
function cleanupExportFiles(days = 1) {
  const exportDir = path.join(__dirname, "exports");
  if (!fs.existsSync(exportDir)) return;

  const files = fs.readdirSync(exportDir);
  const now = Date.now();
  const dayInMs = 24 * 60 * 60 * 1000;

  files.forEach((file) => {
    const filePath = path.join(exportDir, file);
    const stats = fs.statSync(filePath);
    const fileCreationTime = stats.birthtime.getTime();

    // 如果文件创建时间超过指定天数，则删除
    if (now - fileCreationTime > days * dayInMs) {
      fs.unlinkSync(filePath);
    }
  });
}

// 定期清理过期的导出文件
setInterval(() => {
  cleanupExportFiles();
}, 12 * 60 * 60 * 1000); // 每12小时执行一次

module.exports = {
  exportToExcel,
  sheetToArray,
};
