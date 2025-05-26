/* eslint-disable no-console */
const express = require("express");
const { MongoClient } = require("mongodb");
const SocketServer = require("ws").Server;
const uuid = require("uuid");
const _ = require("lodash");
const path = require("path");
const cors = require("cors");
const { applyOp } = require("./op");
const { exportToExcel } = require("./export");

const defaultData = {
  name: "Demo",
  id: uuid.v4(),
  celldata: [{ r: 0, c: 0, v: null }],
  order: 0,
  row: 84,
  column: 60,
  config: {},
  pivotTable: null,
  isPivotTable: false,
  status: 0,
};

const dbName = process.env.MONGODB_DB_NAME || "fortune-sheet";
const collectionName = process.env.MONGODB_COLLECTION_NAME || "workbook";
const uri =
  process.env.MONGODB_URI ||
  "mongodb://root:h0Igwg0qfq21xXyPybJSTJkmvj98MHtdYgynCHZRdK70OceBu8sL0wnHH27sONZ7@119.7.191.6:27017/?directConnection=true";
const client = new MongoClient(uri);

// 存储每个分享码的在线用户
const presencesByShareCode = new Map();

// 生成分享码
function generateShareCode() {
  return Math.random().toString(36).substring(2, 8).toUpperCase();
}

async function initMongoDB() {
  try {
    await client.connect();
    await client.db("admin").command({ ping: 1 });
    console.log("Connected to MongoDB");
  } catch (error) {
    console.error("Failed to connect to MongoDB:", error);
  }
}

initMongoDB();

const app = express();
const port = process.env.PORT || 8081;

// 启用CORS
app.use(cors());
// 解析JSON请求体
app.use(express.json());
// 静态文件目录，用于提供导出的文件
app.use("/exports", express.static(path.join(__dirname, "exports")));

async function getData(shareCode) {
  const db = client.db(dbName);
  const data = await db
    .collection(collectionName)
    .find({ shareCode })
    .toArray();
  data.forEach((sheet) => {
    if (!_.isUndefined(sheet._id)) delete sheet._id;
  });
  return data;
}

// 创建新的工作簿并返回分享码
app.post("/create", async (req, res) => {
  try {
    const shareCode = generateShareCode();
    const db = client.db(dbName);

    // 创建默认工作表
    const defaultSheet = {
      ...defaultData,
      shareCode,
      createdAt: new Date(),
    };

    await db.collection(collectionName).insertOne(defaultSheet);

    res.json({
      success: true,
      shareCode,
      message: "工作簿创建成功",
    });
  } catch (error) {
    console.error("创建工作簿失败:", error);
    res.status(500).json({
      success: false,
      error: error.message || "创建工作簿失败",
    });
  }
});

// 根据分享码获取工作簿数据
app.get("/workbook/:shareCode", async (req, res) => {
  try {
    const { shareCode } = req.params;
    const data = await getData(shareCode);

    if (data.length === 0) {
      return res.status(404).json({
        success: false,
        error: "分享码不存在或工作簿已被删除",
      });
    }

    res.json({
      success: true,
      data,
      shareCode,
    });
  } catch (error) {
    console.error("获取工作簿数据失败:", error);
    res.status(500).json({
      success: false,
      error: error.message || "获取工作簿数据失败",
    });
  }
});

// 导入文件到指定分享码的工作簿
app.post("/import/:shareCode", async (req, res) => {
  try {
    const { shareCode } = req.params;
    const { sheets } = req.body;

    if (!sheets || !Array.isArray(sheets)) {
      return res.status(400).json({
        success: false,
        error: "无效的工作表数据",
      });
    }

    const db = client.db(dbName);
    const coll = db.collection(collectionName);

    // 检查分享码是否存在
    const existingData = await getData(shareCode);
    if (existingData.length === 0) {
      return res.status(404).json({
        success: false,
        error: "分享码不存在",
      });
    }

    // 删除现有数据
    await coll.deleteMany({ shareCode });

    // 插入新的工作表数据
    const sheetsWithShareCode = sheets.map((sheet, index) => ({
      ...sheet,
      shareCode,
      order: index,
      updatedAt: new Date(),
    }));

    await coll.insertMany(sheetsWithShareCode);

    res.json({
      success: true,
      message: "文件导入成功",
      shareCode,
    });
  } catch (error) {
    console.error("导入文件失败:", error);
    res.status(500).json({
      success: false,
      error: error.message || "导入文件失败",
    });
  }
});

// 导出指定分享码的工作簿
app.post("/export/:shareCode", async (req, res) => {
  try {
    const { shareCode } = req.params;
    const { format = "xlsx", fileName = "fortune-sheet-export" } = req.body;

    // 获取指定分享码的工作簿数据
    const sheets = await getData(shareCode);

    if (sheets.length === 0) {
      return res.status(404).json({
        success: false,
        error: "分享码不存在或工作簿为空",
      });
    }

    // 导出为文件
    const result = exportToExcel(sheets, fileName, format);

    // 返回文件下载URL
    const fileUrl = `/exports/${result.fileName}`;

    res.json({
      success: true,
      fileUrl,
      fileName: result.fileName,
      shareCode,
    });
  } catch (error) {
    console.error("导出文件失败:", error);
    res.status(500).json({
      success: false,
      error: error.message || "导出文件失败",
    });
  }
});

// get current workbook data (保持向后兼容)
app.get("/", async (req, res) => {
  res.json(await getData());
});

// drop current data and initialize a new one (保持向后兼容)
app.get("/init", async (req, res) => {
  const db = client.db(dbName);
  const coll = db.collection(collectionName);
  await coll.deleteMany();
  await db.collection(collectionName).insertOne(defaultData);
  res.json({
    ok: true,
  });
});

// 导出Excel或CSV文件 (保持向后兼容)
app.post("/export", async (req, res) => {
  try {
    const { format = "xlsx", fileName = "fortune-sheet-export" } = req.body;

    // 获取当前工作簿数据
    const sheets = await getData();

    // 导出为文件
    const result = exportToExcel(sheets, fileName, format);

    // 返回文件下载URL
    const fileUrl = `/exports/${result.fileName}`;

    res.json({
      success: true,
      fileUrl,
      fileName: result.fileName,
    });
  } catch (error) {
    console.error("导出文件失败:", error);
    res.status(500).json({
      success: false,
      error: error.message || "导出文件失败",
    });
  }
});

const server = app.listen(port, () => {
  console.info(`running on port ${port}`);
});

const connections = {};

const broadcastToOthers = (selfId, shareCode, data) => {
  Object.values(connections).forEach((ws) => {
    if (ws.id !== selfId && ws.shareCode === shareCode) {
      ws.send(data);
    }
  });
};

const wss = new SocketServer({ server, path: "/ws" });

wss.on("connection", (ws) => {
  ws.id = uuid.v4();
  connections[ws.id] = ws;

  ws.on("message", async (data) => {
    const msg = JSON.parse(data.toString());

    if (msg.req === "join") {
      // 加入指定分享码的房间
      ws.shareCode = msg.shareCode;

      // 发送工作簿数据
      const workbookData = await getData(msg.shareCode);
      ws.send(
        JSON.stringify({
          req: "getData",
          data: workbookData,
          shareCode: msg.shareCode,
        })
      );

      // 发送当前在线用户
      const presences = presencesByShareCode.get(msg.shareCode) || [];
      ws.send(JSON.stringify({ req: "addPresences", data: presences }));
    } else if (msg.req === "getData") {
      // 兼容旧版本
      ws.send(
        JSON.stringify({
          req: msg.req,
          data: await getData(),
        })
      );
      const presences = presencesByShareCode.get("default") || [];
      ws.send(JSON.stringify({ req: "addPresences", data: presences }));
    } else if (msg.req === "op") {
      const shareCode = ws.shareCode || "default";
      await applyOp(
        client.db(dbName).collection(collectionName),
        msg.data,
        shareCode
      );
      broadcastToOthers(ws.id, shareCode, data.toString());
    } else if (msg.req === "addPresences") {
      const shareCode = ws.shareCode || "default";
      ws.presences = msg.data;
      broadcastToOthers(ws.id, shareCode, data.toString());

      // 更新该分享码的在线用户列表
      let presences = presencesByShareCode.get(shareCode) || [];
      presences = _.differenceBy(presences, msg.data, (v) =>
        v.userId == null ? v.username : v.userId
      ).concat(msg.data);
      presencesByShareCode.set(shareCode, presences);
    } else if (msg.req === "removePresences") {
      const shareCode = ws.shareCode || "default";
      broadcastToOthers(ws.id, shareCode, data.toString());
    }
  });

  ws.on("close", () => {
    const shareCode = ws.shareCode || "default";

    if (ws.presences) {
      broadcastToOthers(
        ws.id,
        shareCode,
        JSON.stringify({
          req: "removePresences",
          data: ws.presences,
        })
      );

      // 从在线用户列表中移除
      let presences = presencesByShareCode.get(shareCode) || [];
      presences = _.differenceBy(presences, ws.presences, (v) =>
        v.userId == null ? v.username : v.userId
      );
      presencesByShareCode.set(shareCode, presences);
    }
    delete connections[ws.id];
  });
});
