import React, { useCallback, useRef, useState, useEffect } from "react";
import { Meta, StoryFn } from "@storybook/react";
import { Workbook, WorkbookInstance } from "@fortune-sheet/react";
import { Sheet } from "@fortune-sheet/core";
import { exportExcel } from "./utils/exportUtils";

export default {
  component: Workbook,
} as Meta<typeof Workbook>;

// 创建一个增强的Workbook组件，包含导出功能
const EnhancedWorkbook = React.forwardRef<
  WorkbookInstance & {
    exportToExcel: (type?: "xlsx" | "csv", fileName?: string) => void;
  },
  React.ComponentProps<typeof Workbook>
>((props, ref) => {
  const internalRef = useRef<WorkbookInstance>(null);

  // 将内部ref暴露给外部
  useEffect(() => {
    if (!ref || !internalRef.current) return;

    if (typeof ref === "function") {
      ref({
        ...internalRef.current,
        exportToExcel: (type = "xlsx", fileName = "fortune-sheet-export") => {
          const sheets = internalRef.current?.getAllSheets() || [];
          exportExcel(sheets, { type, fileName });
        },
      });
    } else {
      // @ts-ignore 扩展ref对象
      ref.current = {
        ...internalRef.current,
        exportToExcel: (type = "xlsx", fileName = "fortune-sheet-export") => {
          const sheets = internalRef.current?.getAllSheets() || [];
          exportExcel(sheets, { type, fileName });
        },
      };
    }
  }, [ref]);

  return <Workbook {...props} ref={internalRef} />;
});

// 创建一个包含样本数据的组件
const SimpleExportDemo: React.FC = () => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const [data, setData] = useState<Sheet[]>([
    {
      name: "Sheet1",
      celldata: [
        { r: 0, c: 0, v: { v: "姓名" } },
        { r: 0, c: 1, v: { v: "年龄" } },
        { r: 0, c: 2, v: { v: "城市" } },
        { r: 1, c: 0, v: { v: "张三" } },
        { r: 1, c: 1, v: { v: 25 } },
        { r: 1, c: 2, v: { v: "北京" } },
        { r: 2, c: 0, v: { v: "李四" } },
        { r: 2, c: 1, v: { v: 30 } },
        { r: 2, c: 2, v: { v: "上海" } },
      ],
      order: 0,
      row: 100,
      column: 26,
    },
  ]);

  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  const handleExportExcel = () => {
    if (workbookRef.current) {
      const allSheets = workbookRef.current.getAllSheets();
      exportExcel(allSheets, { type: "xlsx", fileName: "导出数据-Excel" });
    }
  };

  const handleExportCSV = () => {
    if (workbookRef.current) {
      const allSheets = workbookRef.current.getAllSheets();
      exportExcel(allSheets, { type: "csv", fileName: "导出数据-CSV" });
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <div style={{ padding: "10px", display: "flex", gap: "10px" }}>
        <button
          type="button"
          onClick={handleExportExcel}
          style={{
            padding: "8px 16px",
            backgroundColor: "#1890ff",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
          }}
        >
          导出Excel
        </button>
        <button
          type="button"
          onClick={handleExportCSV}
          style={{
            padding: "8px 16px",
            backgroundColor: "#52c41a",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
          }}
        >
          导出CSV
        </button>
      </div>
      <div style={{ flex: 1 }}>
        <Workbook ref={workbookRef} data={data} onChange={onChange} />
      </div>
    </div>
  );
};

// 创建一个增强型API的示例
const EnhancedExportDemo: React.FC = () => {
  // 使用增强的Workbook类型
  const workbookRef = useRef<
    WorkbookInstance & {
      exportToExcel: (type?: "xlsx" | "csv", fileName?: string) => void;
    }
  >(null);

  const [data, setData] = useState<Sheet[]>([
    {
      name: "人员数据",
      celldata: [
        { r: 0, c: 0, v: { v: "姓名" } },
        { r: 0, c: 1, v: { v: "年龄" } },
        { r: 0, c: 2, v: { v: "城市" } },
        { r: 1, c: 0, v: { v: "张三" } },
        { r: 1, c: 1, v: { v: 25 } },
        { r: 1, c: 2, v: { v: "北京" } },
        { r: 2, c: 0, v: { v: "李四" } },
        { r: 2, c: 1, v: { v: 30 } },
        { r: 2, c: 2, v: { v: "上海" } },
      ],
      order: 0,
      row: 100,
      column: 26,
    },
    {
      name: "产品数据",
      celldata: [
        { r: 0, c: 0, v: { v: "产品" } },
        { r: 0, c: 1, v: { v: "价格" } },
        { r: 1, c: 0, v: { v: "笔记本" } },
        { r: 1, c: 1, v: { v: 5999 } },
      ],
      order: 1,
      row: 100,
      column: 26,
    },
  ]);

  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  const handleExportExcel = () => {
    if (workbookRef.current) {
      // 直接使用增强的API
      workbookRef.current.exportToExcel("xlsx", "增强API导出-Excel");
    }
  };

  const handleExportCSV = () => {
    if (workbookRef.current) {
      // 直接使用增强的API
      workbookRef.current.exportToExcel("csv", "增强API导出-CSV");
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <div style={{ padding: "10px", display: "flex", gap: "10px" }}>
        <button
          type="button"
          onClick={handleExportExcel}
          style={{
            padding: "8px 16px",
            backgroundColor: "#1890ff",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
          }}
        >
          增强API导出Excel
        </button>
        <button
          type="button"
          onClick={handleExportCSV}
          style={{
            padding: "8px 16px",
            backgroundColor: "#52c41a",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
          }}
        >
          增强API导出CSV
        </button>
      </div>
      <div style={{ flex: 1 }}>
        <EnhancedWorkbook ref={workbookRef} data={data} onChange={onChange} />
      </div>
    </div>
  );
};

export const 基础导出示例: StoryFn<typeof Workbook> = () => (
  <SimpleExportDemo />
);
export const 增强API导出示例: StoryFn<typeof Workbook> = () => (
  <EnhancedExportDemo />
);
