import React, { useCallback, useRef, useState } from "react";
import { Meta, StoryFn } from "@storybook/react";
import { Workbook, WorkbookInstance } from "@fortune-sheet/react";
import { Sheet } from "@fortune-sheet/core";

export default {
  component: Workbook,
} as Meta<typeof Workbook>;

// 后端服务器URL
const SERVER_URL = process.env.STORYBOOK_SERVER_URL || "http://localhost:8081";

// 创建一个包含与服务器交互的导出功能组件
const ServerExportDemo: React.FC = () => {
  const workbookRef = useRef<WorkbookInstance>(null);
  const [data, setData] = useState<Sheet[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{
    text: string;
    type: "success" | "error" | "info";
  } | null>(null);

  // 处理导出请求
  const handleExport = async (format: "xlsx" | "csv", fileName: string) => {
    try {
      setIsLoading(true);
      setMessage({
        text: `正在导出${format.toUpperCase()}文件...`,
        type: "info",
      });

      const response = await fetch(`${SERVER_URL}/export`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ format, fileName }),
      });

      if (!response.ok) {
        throw new Error(`导出${format.toUpperCase()}文件失败`);
      }

      const result = await response.json();

      if (result.success) {
        // 创建一个下载链接
        const downloadUrl = `${SERVER_URL}${result.fileUrl}`;
        const link = document.createElement("a");
        link.href = downloadUrl;
        link.download = result.fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setMessage({
          text: `${format.toUpperCase()}文件导出成功`,
          type: "success",
        });
      } else {
        throw new Error(result.error || `导出${format.toUpperCase()}文件失败`);
      }
    } catch (error) {
      console.error(`导出${format.toUpperCase()}文件出错:`, error);
      setMessage({
        text:
          error instanceof Error
            ? error.message
            : `导出${format.toUpperCase()}文件出错`,
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  };

  // 导出为Excel
  const handleExportExcel = async () => {
    await handleExport("xlsx", "fortune-sheet-数据");
  };

  // 导出为CSV
  const handleExportCSV = async () => {
    await handleExport("csv", "fortune-sheet-数据");
  };

  // 初始化服务器数据
  const initServerData = useCallback(async () => {
    try {
      setIsLoading(true);
      setMessage({ text: "正在初始化服务器数据...", type: "info" });
      const response = await fetch(`${SERVER_URL}/init`);
      if (!response.ok) {
        throw new Error("初始化服务器数据失败");
      }

      // 再次获取数据
      const dataResponse = await fetch(`${SERVER_URL}/`);
      if (!dataResponse.ok) {
        throw new Error("获取初始化数据失败");
      }

      const jsonData = await dataResponse.json();
      setData(jsonData);
      setMessage({ text: "服务器数据初始化成功", type: "success" });
    } catch (error) {
      console.error("初始化服务器数据出错:", error);
      setMessage({
        text: error instanceof Error ? error.message : "初始化服务器数据出错",
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  }, []);

  // 从服务器获取数据
  const fetchData = useCallback(async () => {
    try {
      setIsLoading(true);
      setMessage({ text: "正在加载数据...", type: "info" });
      const response = await fetch(`${SERVER_URL}/`);
      if (!response.ok) {
        throw new Error("获取数据失败");
      }
      const jsonData = await response.json();

      if (Array.isArray(jsonData) && jsonData.length > 0) {
        setData(jsonData);
        setMessage({ text: "数据加载成功", type: "success" });
      } else {
        // 如果服务器没有数据，初始化服务器
        await initServerData();
      }
    } catch (error) {
      console.error("获取数据出错:", error);
      setMessage({
        text: error instanceof Error ? error.message : "获取数据出错",
        type: "error",
      });
    } finally {
      setIsLoading(false);
    }
  }, [initServerData]);

  // 初始加载数据
  React.useEffect(() => {
    fetchData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  // 计算信息提示的样式
  const getMessageStyle = (type: "success" | "error" | "info") => {
    let backgroundColor = "#e6f7ff";
    let borderColor = "#91d5ff";
    let textColor = "#1890ff";

    if (type === "success") {
      backgroundColor = "#f6ffed";
      borderColor = "#b7eb8f";
      textColor = "#52c41a";
    } else if (type === "error") {
      backgroundColor = "#fff2f0";
      borderColor = "#ffccc7";
      textColor = "#ff4d4f";
    }

    return {
      padding: "8px 16px",
      borderRadius: "4px",
      backgroundColor,
      border: `1px solid ${borderColor}`,
      color: textColor,
    };
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <div
        style={{
          padding: "10px",
          display: "flex",
          gap: "10px",
          flexWrap: "wrap",
          alignItems: "center",
        }}
      >
        <button
          type="button"
          onClick={fetchData}
          disabled={isLoading}
          style={{
            padding: "8px 16px",
            backgroundColor: "#1890ff",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isLoading ? "not-allowed" : "pointer",
            opacity: isLoading ? 0.7 : 1,
          }}
        >
          刷新数据
        </button>
        <button
          type="button"
          onClick={initServerData}
          disabled={isLoading}
          style={{
            padding: "8px 16px",
            backgroundColor: "#faad14",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isLoading ? "not-allowed" : "pointer",
            opacity: isLoading ? 0.7 : 1,
          }}
        >
          初始化服务器
        </button>
        <button
          type="button"
          onClick={handleExportExcel}
          disabled={isLoading || data.length === 0}
          style={{
            padding: "8px 16px",
            backgroundColor: "#52c41a",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isLoading || data.length === 0 ? "not-allowed" : "pointer",
            opacity: isLoading || data.length === 0 ? 0.7 : 1,
          }}
        >
          导出为Excel
        </button>
        <button
          type="button"
          onClick={handleExportCSV}
          disabled={isLoading || data.length === 0}
          style={{
            padding: "8px 16px",
            backgroundColor: "#722ed1",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: isLoading || data.length === 0 ? "not-allowed" : "pointer",
            opacity: isLoading || data.length === 0 ? 0.7 : 1,
          }}
        >
          导出为CSV
        </button>

        {message && (
          <div style={getMessageStyle(message.type)}>{message.text}</div>
        )}
      </div>

      <div style={{ flex: 1 }}>
        {data.length > 0 ? (
          <Workbook ref={workbookRef} data={data} onChange={onChange} />
        ) : (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              height: "100%",
              fontSize: "16px",
              color: "#999",
            }}
          >
            {isLoading ? "加载中..." : "数据为空，请点击初始化服务器"}
          </div>
        )}
      </div>
    </div>
  );
};

export const 服务器导出示例: StoryFn<typeof Workbook> = () => (
  <ServerExportDemo />
);
