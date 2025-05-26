import React, { useCallback, useRef, useState } from "react";
import { Meta, StoryFn } from "@storybook/react";
import { Workbook, WorkbookInstance } from "@fortune-sheet/react";
import { Sheet } from "@fortune-sheet/core";
import { importFile } from "./utils/importUtils";

export default {
  component: Workbook,
} as Meta<typeof Workbook>;

// 创建一个文件上传按钮组件
const FileUploadButton: React.FC<{
  onUpload: (file: File) => void;
  accept?: string;
  disabled?: boolean;
  children: React.ReactNode;
}> = ({ onUpload, accept = ".xlsx,.xls,.csv", disabled = false, children }) => {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleClick = () => {
    if (inputRef.current) {
      inputRef.current.click();
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { files } = e.target;
    if (files && files.length > 0) {
      onUpload(files[0]);
      // 重置input，允许重复上传相同文件
      if (inputRef.current) {
        inputRef.current.value = "";
      }
    }
  };

  return (
    <>
      <button
        type="button"
        onClick={handleClick}
        disabled={disabled}
        style={{
          padding: "8px 16px",
          backgroundColor: disabled ? "#ccc" : "#1890ff",
          color: "white",
          border: "none",
          borderRadius: "4px",
          cursor: disabled ? "not-allowed" : "pointer",
        }}
      >
        {children}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onChange={handleChange}
        style={{ display: "none" }}
      />
    </>
  );
};

// 创建一个基础的文件导入示例
const BasicImportDemo: React.FC = () => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const [data, setData] = useState<Sheet[]>([
    {
      name: "Sheet1",
      celldata: [],
      order: 0,
      row: 50,
      column: 26,
      config: {},
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{
    text: string;
    type: "success" | "error";
  } | null>(null);
  const [importId, setImportId] = useState(0); // 添加导入标识，用于触发数据更新

  const handleFileUpload = async (file: File) => {
    setIsLoading(true);
    setMessage(null);

    try {
      console.log(
        "开始导入文件:",
        file.name,
        "类型:",
        file.type,
        "大小:",
        file.size,
        "bytes"
      );

      // 导入文件
      const importedSheets = await importFile(file);

      // 打印导入的Excel数据
      console.log("导入的Excel数据:", importedSheets);
      console.log("sheet数量:", importedSheets.length);

      if (importedSheets.length > 0) {
        console.log("第一个sheet样本数据:", {
          name: importedSheets[0].name,
          rowCount: importedSheets[0].row,
          colCount: importedSheets[0].column,
          cellCount: importedSheets[0].celldata?.length || 0,
        });
      }

      // 深度处理导入的数据，确保格式正确
      const processedSheets = importedSheets.map((sheet) => ({
        ...sheet,
        // 确保id唯一性
        id:
          sheet.id ||
          `sheet_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`,
        // 确保每个sheet有必要的默认值
        config: sheet.config || {},
        status: sheet.status || 1,
        // 确保celldata格式正确
        celldata: (sheet.celldata || []).map((cell) => ({
          ...cell,
          v: cell.v
            ? {
                ...cell.v,
                // 确保m属性存在
                m:
                  cell.v.m ||
                  (cell.v.v !== null && cell.v.v !== undefined
                    ? String(cell.v.v)
                    : ""),
              }
            : { v: "", m: "" },
        })),
      }));

      console.log("处理后的sheet数据:", processedSheets);

      // 如果当前sheet是空白的，直接替换；否则添加新的sheet
      let newData;
      if (
        data.length === 1 &&
        (!data[0].celldata || data[0].celldata.length === 0)
      ) {
        // 创建全新的数据对象而不是修改现有对象
        newData = [...processedSheets];
        console.log("替换空白sheet");
      } else {
        // 确保创建新的数组对象
        newData = [
          ...data.map((sheet) => ({ ...sheet })), // 深拷贝现有sheet
          ...processedSheets.map((sheet, index) => ({
            ...sheet,
            order: data.length + index,
          })),
        ];
        console.log("添加到现有sheets");
      }

      console.log("更新前数据:", data);
      console.log("更新后数据:", newData);

      // 更新数据状态
      setData(newData);

      // 增加导入ID触发更新
      setImportId((prev) => prev + 1);
      console.log("更新importId触发表格刷新");

      setMessage({
        text: `成功导入文件 "${file.name}"`,
        type: "success",
      });
    } catch (error) {
      console.error("导入文件失败:", error);

      // 记录更详细的错误信息
      if (error instanceof Error) {
        console.error("错误类型:", error.name);
        console.error("错误消息:", error.message);
        console.error("错误堆栈:", error.stack);
      }

      setMessage({
        text: error instanceof Error ? error.message : "导入文件失败",
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  };

  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <div
        style={{
          padding: "10px",
          display: "flex",
          gap: "10px",
          alignItems: "center",
        }}
      >
        <FileUploadButton onUpload={handleFileUpload} disabled={isLoading}>
          {isLoading ? "正在导入..." : "导入Excel/CSV"}
        </FileUploadButton>

        {message && (
          <div
            style={{
              padding: "8px 16px",
              borderRadius: "4px",
              backgroundColor:
                message.type === "success" ? "#f6ffed" : "#fff2f0",
              border: `1px solid ${
                message.type === "success" ? "#b7eb8f" : "#ffccc7"
              }`,
              color: message.type === "success" ? "#52c41a" : "#ff4d4f",
            }}
          >
            {message.text}
          </div>
        )}

        {isLoading && (
          <div
            style={{
              padding: "8px 16px",
              borderRadius: "4px",
              backgroundColor: "#e6f7ff",
              border: "1px solid #91d5ff",
              color: "#1890ff",
              display: "flex",
              alignItems: "center",
            }}
          >
            <div
              style={{
                width: "16px",
                height: "16px",
                borderRadius: "50%",
                border: "2px solid #1890ff",
                borderTopColor: "transparent",
                marginRight: "8px",
                animation: "spin 1s linear infinite",
              }}
            />
            正在处理文件，请稍候...
          </div>
        )}
      </div>

      <style>
        {`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}
      </style>

      <div style={{ flex: 1 }}>
        <Workbook
          key={`workbook-${importId}`}
          ref={workbookRef}
          data={data}
          onChange={onChange}
        />
      </div>
    </div>
  );
};

// 创建一个拖放导入示例
const DragDropImportDemo: React.FC = () => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const [data, setData] = useState<Sheet[]>([
    {
      name: "Sheet1",
      celldata: [],
      order: 0,
      row: 50,
      column: 26,
      config: {},
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{
    text: string;
    type: "success" | "error";
  } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [importId, setImportId] = useState(0); // 添加导入标识，用于触发数据更新

  const handleFileUpload = async (file: File) => {
    setIsLoading(true);
    setMessage(null);

    try {
      console.log(
        "开始导入文件:",
        file.name,
        "类型:",
        file.type,
        "大小:",
        file.size,
        "bytes"
      );

      // 导入文件
      const importedSheets = await importFile(file);

      // 打印导入的Excel数据
      console.log("导入的Excel数据:", importedSheets);
      console.log("sheet数量:", importedSheets.length);

      if (importedSheets.length > 0) {
        console.log("第一个sheet样本数据:", {
          name: importedSheets[0].name,
          rowCount: importedSheets[0].row,
          colCount: importedSheets[0].column,
          cellCount: importedSheets[0].celldata?.length || 0,
        });
      }

      // 深度处理导入的数据，确保格式正确
      const processedSheets = importedSheets.map((sheet) => ({
        ...sheet,
        // 确保id唯一性
        id:
          sheet.id ||
          `sheet_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`,
        // 确保每个sheet有必要的默认值
        config: sheet.config || {},
        status: sheet.status || 1,
        // 确保celldata格式正确
        celldata: (sheet.celldata || []).map((cell) => ({
          ...cell,
          v: cell.v
            ? {
                ...cell.v,
                // 确保m属性存在
                m:
                  cell.v.m ||
                  (cell.v.v !== null && cell.v.v !== undefined
                    ? String(cell.v.v)
                    : ""),
              }
            : { v: "", m: "" },
        })),
      }));

      console.log("处理后的sheet数据:", processedSheets);

      // 如果当前sheet是空白的，直接替换；否则添加新的sheet
      let newData;
      if (
        data.length === 1 &&
        (!data[0].celldata || data[0].celldata.length === 0)
      ) {
        // 创建全新的数据对象而不是修改现有对象
        newData = [...processedSheets];
        console.log("替换空白sheet");
      } else {
        // 确保创建新的数组对象
        newData = [
          ...data.map((sheet) => ({ ...sheet })), // 深拷贝现有sheet
          ...processedSheets.map((sheet, index) => ({
            ...sheet,
            order: data.length + index,
          })),
        ];
        console.log("添加到现有sheets");
      }

      console.log("更新前数据:", data);
      console.log("更新后数据:", newData);

      // 更新数据状态
      setData(newData);

      // 增加导入ID触发更新
      setImportId((prev) => prev + 1);
      console.log("更新importId触发表格刷新");

      setMessage({
        text: `成功导入文件 "${file.name}"`,
        type: "success",
      });
    } catch (error) {
      console.error("导入文件失败:", error);

      // 记录更详细的错误信息
      if (error instanceof Error) {
        console.error("错误类型:", error.name);
        console.error("错误消息:", error.message);
        console.error("错误堆栈:", error.stack);
      }

      setMessage({
        text: error instanceof Error ? error.message : "导入文件失败",
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  };

  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);

    const { files } = e.dataTransfer;
    if (files && files.length > 0) {
      handleFileUpload(files[0]);
    }
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        position: "relative",
      }}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <div
        style={{
          padding: "10px",
          display: "flex",
          gap: "10px",
          alignItems: "center",
        }}
      >
        <FileUploadButton onUpload={handleFileUpload} disabled={isLoading}>
          {isLoading ? "正在导入..." : "导入Excel/CSV"}
        </FileUploadButton>

        <div
          style={{
            marginLeft: "10px",
            color: "#666",
          }}
        >
          或将文件拖放到此处
        </div>

        {message && (
          <div
            style={{
              padding: "8px 16px",
              borderRadius: "4px",
              backgroundColor:
                message.type === "success" ? "#f6ffed" : "#fff2f0",
              border: `1px solid ${
                message.type === "success" ? "#b7eb8f" : "#ffccc7"
              }`,
              color: message.type === "success" ? "#52c41a" : "#ff4d4f",
            }}
          >
            {message.text}
          </div>
        )}

        {isLoading && (
          <div
            style={{
              padding: "8px 16px",
              borderRadius: "4px",
              backgroundColor: "#e6f7ff",
              border: "1px solid #91d5ff",
              color: "#1890ff",
              display: "flex",
              alignItems: "center",
            }}
          >
            <div
              style={{
                width: "16px",
                height: "16px",
                borderRadius: "50%",
                border: "2px solid #1890ff",
                borderTopColor: "transparent",
                marginRight: "8px",
                animation: "spin 1s linear infinite",
              }}
            />
            正在处理文件，请稍候...
          </div>
        )}
      </div>

      <style>
        {`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}
      </style>

      <div
        style={{
          flex: 1,
          position: "relative",
        }}
      >
        <Workbook
          key={`workbook-${importId}`}
          ref={workbookRef}
          data={data}
          onChange={onChange}
        />

        {isDragging && (
          <div
            style={{
              position: "absolute",
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              backgroundColor: "rgba(24, 144, 255, 0.1)",
              border: "2px dashed #1890ff",
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              zIndex: 100,
            }}
          >
            <div
              style={{
                padding: "20px",
                backgroundColor: "white",
                borderRadius: "8px",
                boxShadow: "0 2px 8px rgba(0, 0, 0, 0.15)",
              }}
            >
              释放鼠标以导入文件
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export const 基础导入示例: StoryFn<typeof Workbook> = () => <BasicImportDemo />;
export const 拖放导入示例: StoryFn<typeof Workbook> = () => (
  <DragDropImportDemo />
);
