import React, {
  useState,
  useCallback,
  useEffect,
  useRef,
  useMemo,
} from "react";
import { Meta, StoryFn } from "@storybook/react";
import { Sheet, Op, Selection, colors } from "@fortune-sheet/core";
import { Workbook, WorkbookInstance } from "@fortune-sheet/react";
import { v4 as uuidv4 } from "uuid";
import { hashCode } from "./utils";
import { importFile } from "./utils/importUtils";

export default {
  component: Workbook,
  title: "协同工作簿",
} as Meta<typeof Workbook>;

// 文件上传按钮组件
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
          marginRight: "8px",
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

// 主要的协同工作簿组件
const CollaborativeWorkbook: React.FC = () => {
  const [data, setData] = useState<Sheet[]>();
  const [connectionError, setConnectionError] = useState<string | null>(null);
  const [shareCode, setShareCode] = useState<string>("");
  const [inputShareCode, setInputShareCode] = useState<string>("");
  const [isConnected, setIsConnected] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{
    text: string;
    type: "success" | "error" | "info";
  } | null>(null);
  const [onlineUsers, setOnlineUsers] = useState<any[]>([]);
  const [importId, setImportId] = useState(0); // 添加导入标识，用于触发数据更新

  const wsRef = useRef<WebSocket>();
  const workbookRef = useRef<WorkbookInstance>(null);
  const lastSelection = useRef<any>();
  const { username, userId } = useMemo(() => {
    const _userId = uuidv4();
    return { username: `用户-${_userId.slice(0, 3)}`, userId: _userId };
  }, []);

  // 显示消息
  const showMessage = (text: string, type: "success" | "error" | "info") => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 3000);
  };

  // 连接到工作簿
  const connectToWorkbook = useCallback((code: string) => {
    if (!code.trim()) {
      showMessage("请输入分享码", "error");
      return;
    }

    // 关闭现有连接
    if (wsRef.current) {
      wsRef.current.close();
    }

    setIsLoading(true);
    setConnectionError(null);

    const wsUrl = process.env.STORYBOOK_WS_URL || "ws://localhost:8081/ws";
    const socket = new WebSocket(wsUrl);
    wsRef.current = socket;

    socket.onopen = () => {
      setIsConnected(true);
      setShareCode(code);
      // 加入指定分享码的房间
      socket.send(JSON.stringify({ req: "join", shareCode: code }));
      showMessage(`正在连接到工作簿：${code}...`, "info");
    };

    socket.onmessage = (e) => {
      const msg = JSON.parse(e.data);
      if (msg.req === "getData") {
        if (Array.isArray(msg.data) && msg.data.length > 0) {
          // 深度处理导入的数据，确保格式正确
          const processedSheets = msg.data.map((sheet: any) => ({
            ...sheet,
            // 确保id唯一性
            id: sheet._id || sheet.id || uuidv4(),
            // 确保每个sheet有必要的默认值
            config: sheet.config || {},
            status: sheet.status || 1,
            // 确保celldata格式正确
            celldata: (sheet.celldata || []).map((cell: any) => ({
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
          setData(processedSheets);
          // 增加导入ID触发更新
          setImportId((prev) => prev + 1);
        } else {
          // 如果没有数据，创建默认工作表
          const defaultSheet = {
            id: uuidv4(),
            name: "Sheet1",
            celldata: [],
            row: 100,
            column: 26,
            config: {},
          };
          setData([defaultSheet]);
        }
        setIsLoading(false);
        showMessage(`成功连接到工作簿：${code}`, "success");
      } else if (msg.req === "op") {
        workbookRef.current?.applyOp(msg.data);
      } else if (msg.req === "addPresences") {
        // 过滤掉当前用户，避免显示自己为协同用户
        const filteredUsers = (msg.data || []).filter(
          (user: any) => user.userId !== userId
        );
        setOnlineUsers(filteredUsers);
        workbookRef.current?.addPresences(msg.data);
      } else if (msg.req === "removePresences") {
        workbookRef.current?.removePresences(msg.data);
      }
    };

    socket.onerror = () => {
      setConnectionError("连接失败，请检查分享码是否正确或服务器是否运行");
      setIsConnected(false);
      setIsLoading(false);
      showMessage("连接失败", "error");
    };

    socket.onclose = () => {
      setIsConnected(false);
      setIsLoading(false);
    };
  }, []);

  // 创建新工作簿
  const createNewWorkbook = async () => {
    setIsLoading(true);
    try {
      const response = await fetch("http://localhost:8081/create", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
      });

      const result = await response.json();
      if (result.success) {
        setShareCode(result.shareCode);
        setInputShareCode(result.shareCode);
        showMessage(`工作簿创建成功！分享码：${result.shareCode}`, "success");
        // 自动连接到新创建的工作簿
        connectToWorkbook(result.shareCode);
      } else {
        showMessage(result.error || "创建工作簿失败", "error");
      }
    } catch (err) {
      console.error("创建工作簿失败:", err);
      showMessage("创建工作簿失败，请检查网络连接", "error");
    } finally {
      setIsLoading(false);
    }
  };

  // 断开连接
  const disconnect = () => {
    if (wsRef.current) {
      wsRef.current.close();
    }
    setIsConnected(false);
    setShareCode("");
    setData(undefined);
    setOnlineUsers([]);
    showMessage("已断开连接", "info");
  };

  // 处理操作
  const onOp = useCallback((op: Op[]) => {
    const socket = wsRef.current;
    if (!socket || socket.readyState !== WebSocket.OPEN) return;
    socket.send(JSON.stringify({ req: "op", data: op }));
  }, []);

  // 处理数据变化
  const onChange = useCallback((d: Sheet[]) => {
    setData(d);
  }, []);

  // 处理选择变化
  const afterSelectionChange = useCallback(
    (sheetId: string, selection: Selection) => {
      const socket = wsRef.current;
      if (!socket || socket.readyState !== WebSocket.OPEN) return;

      const s = {
        r: selection.row[0],
        c: selection.column[0],
      };

      if (
        lastSelection.current?.r === s.r &&
        lastSelection.current?.c === s.c
      ) {
        return;
      }

      lastSelection.current = s;
      socket.send(
        JSON.stringify({
          req: "addPresences",
          data: [
            {
              sheetId,
              username,
              userId,
              color: colors[Math.abs(hashCode(userId)) % colors.length],
              selection: s,
            },
          ],
        })
      );
    },
    [userId, username]
  );

  // 文件导入
  const handleFileUpload = async (file: File) => {
    if (!shareCode) {
      showMessage("请先连接到工作簿", "error");
      return;
    }

    setIsLoading(true);
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

      // 发送到服务器
      const response = await fetch(
        `http://localhost:8081/import/${shareCode}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ sheets: processedSheets }),
        }
      );

      const result = await response.json();
      if (result.success) {
        showMessage(`文件 "${file.name}" 导入成功`, "success");
        // 重新获取数据
        if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
          wsRef.current.send(JSON.stringify({ req: "join", shareCode }));
        }
      } else {
        showMessage(result.error || "导入失败", "error");
      }
    } catch (importError) {
      console.error("导入文件失败:", importError);

      // 记录更详细的错误信息
      if (importError instanceof Error) {
        console.error("错误类型:", importError.name);
        console.error("错误消息:", importError.message);
        console.error("错误堆栈:", importError.stack);
      }

      showMessage(
        importError instanceof Error ? importError.message : "导入文件失败",
        "error"
      );
    } finally {
      setIsLoading(false);
    }
  };

  // 导出Excel
  const handleExportExcel = async () => {
    if (!shareCode) {
      showMessage("请先连接到工作簿", "error");
      return;
    }

    try {
      const response = await fetch(
        `http://localhost:8081/export/${shareCode}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            format: "xlsx",
            fileName: `协同工作簿-${shareCode}`,
          }),
        }
      );

      const result = await response.json();
      if (result.success) {
        // 下载文件
        const link = document.createElement("a");
        link.href = `http://localhost:8081${result.fileUrl}`;
        link.download = result.fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        showMessage("Excel导出成功", "success");
      } else {
        showMessage(result.error || "导出失败", "error");
      }
    } catch (exportError) {
      console.error("导出失败:", exportError);
      showMessage("导出失败", "error");
    }
  };

  // 导出CSV
  const handleExportCSV = async () => {
    if (!shareCode) {
      showMessage("请先连接到工作簿", "error");
      return;
    }

    try {
      const response = await fetch(
        `http://localhost:8081/export/${shareCode}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            format: "csv",
            fileName: `协同工作簿-${shareCode}`,
          }),
        }
      );

      const result = await response.json();
      if (result.success) {
        // 下载文件
        const link = document.createElement("a");
        link.href = `http://localhost:8081${result.fileUrl}`;
        link.download = result.fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        showMessage("CSV导出成功", "success");
      } else {
        showMessage(result.error || "导出失败", "error");
      }
    } catch (csvError) {
      console.error("导出失败:", csvError);
      showMessage("导出失败", "error");
    }
  };

  // 清理连接
  useEffect(() => {
    return () => {
      if (wsRef.current) {
        if (
          wsRef.current.readyState === WebSocket.OPEN &&
          lastSelection.current
        ) {
          wsRef.current.send(
            JSON.stringify({
              req: "removePresences",
              data: [{ userId, username }],
            })
          );
        }
        wsRef.current.close();
      }
    };
  }, [userId, username]);

  // 获取消息背景色
  const getMessageBackgroundColor = (type: string) => {
    if (type === "success") return "#f6ffed";
    if (type === "error") return "#fff2f0";
    return "#e6f7ff";
  };

  // 获取消息边框色
  const getMessageBorderColor = (type: string) => {
    if (type === "success") return "#b7eb8f";
    if (type === "error") return "#ffccc7";
    return "#91d5ff";
  };

  // 获取消息文字色
  const getMessageTextColor = (type: string) => {
    if (type === "success") return "#52c41a";
    if (type === "error") return "#ff4d4f";
    return "#1890ff";
  };

  if (connectionError) {
    return (
      <div style={{ padding: 16 }}>
        <h3>连接失败</h3>
        <p>{connectionError}</p>
        <p>请确保：</p>
        <ol>
          <li>后端服务器正在运行 (node backend-demo/index.js)</li>
          <li>MongoDB 正在运行</li>
          <li>分享码正确</li>
        </ol>
        <button
          type="button"
          onClick={() => {
            setConnectionError(null);
            setInputShareCode("");
          }}
          style={{
            padding: "8px 16px",
            backgroundColor: "#1890ff",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
          }}
        >
          重试
        </button>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      {/* 工具栏 */}
      <div
        style={{
          padding: "10px",
          borderBottom: "1px solid #e8e8e8",
          backgroundColor: "#fafafa",
        }}
      >
        {/* 连接控制 */}
        <div
          style={{
            marginBottom: "10px",
            display: "flex",
            alignItems: "center",
            gap: "10px",
          }}
        >
          {!isConnected ? (
            <>
              <input
                type="text"
                placeholder="输入分享码"
                value={inputShareCode}
                onChange={(e) => setInputShareCode(e.target.value)}
                onKeyPress={(e) => {
                  if (
                    e.key === "Enter" &&
                    inputShareCode.trim() &&
                    !isLoading
                  ) {
                    connectToWorkbook(inputShareCode);
                  }
                }}
                style={{
                  padding: "8px",
                  border: "1px solid #d9d9d9",
                  borderRadius: "4px",
                  width: "120px",
                }}
              />
              <button
                type="button"
                onClick={() => connectToWorkbook(inputShareCode)}
                disabled={isLoading}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#52c41a",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: isLoading ? "not-allowed" : "pointer",
                }}
              >
                {isLoading ? "连接中..." : "连接工作簿"}
              </button>
              <span style={{ color: "#999" }}>或</span>
              <button
                type="button"
                onClick={createNewWorkbook}
                disabled={isLoading}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#1890ff",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: isLoading ? "not-allowed" : "pointer",
                }}
              >
                {isLoading ? "创建中..." : "创建新工作簿"}
              </button>
            </>
          ) : (
            <>
              <span style={{ color: "#52c41a", fontWeight: "bold" }}>
                已连接到工作簿：{shareCode}
              </span>
              <button
                type="button"
                onClick={disconnect}
                style={{
                  padding: "8px 16px",
                  backgroundColor: "#ff4d4f",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                }}
              >
                断开连接
              </button>
              {onlineUsers.length > 0 && (
                <span style={{ color: "#666", marginLeft: "20px" }}>
                  在线用户：{onlineUsers.map((u) => u.username).join(", ")}
                </span>
              )}
            </>
          )}
        </div>

        {/* 文件操作 */}
        {isConnected && (
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <FileUploadButton onUpload={handleFileUpload} disabled={isLoading}>
              {isLoading ? "导入中..." : "导入文件"}
            </FileUploadButton>
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
        )}

        {/* 消息显示 */}
        {message && (
          <div
            style={{
              marginTop: "10px",
              padding: "8px 16px",
              borderRadius: "4px",
              backgroundColor: getMessageBackgroundColor(message.type),
              border: `1px solid ${getMessageBorderColor(message.type)}`,
              color: getMessageTextColor(message.type),
            }}
          >
            {message.text}
          </div>
        )}
      </div>

      {/* 工作簿 */}
      <div style={{ flex: 1 }}>
        {isConnected && data ? (
          <Workbook
            key={`workbook-${importId}`}
            ref={workbookRef}
            data={data}
            onChange={onChange}
            onOp={onOp}
            hooks={{
              afterSelectionChange,
            }}
          />
        ) : (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              height: "100%",
              color: "#999",
            }}
          >
            {isLoading ? "加载中..." : "请连接到工作簿或创建新工作簿"}
          </div>
        )}
      </div>
    </div>
  );
};

const Template: StoryFn<typeof Workbook> = () => {
  return <CollaborativeWorkbook />;
};

export const 协同工作簿 = Template.bind({});
