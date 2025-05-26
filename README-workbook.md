# Fortune Sheet 分表格协同功能

这个功能实现了基于分享码的表格协同编辑，每个分享码对应一个独立的工作簿，支持导入导出和实时协同编辑。

## 功能特性

- 🔗 **分享码机制**：每个工作簿都有唯一的分享码，用户通过分享码加入协同编辑
- 📁 **文件导入导出**：支持 Excel 和 CSV 文件的导入导出
- 👥 **实时协同**：多用户可以同时编辑同一个工作簿，实时同步
- 🎨 **用户标识**：不同用户有不同的颜色标识，可以看到其他用户的光标位置
- 💾 **数据持久化**：所有数据保存在 MongoDB 中，支持断线重连

## 使用方法

### 1. 启动后端服务

```bash
cd backend-demo
node index.js
```

确保 MongoDB 正在运行。

### 2. 查看 Storybook 示例

```bash
npm run storybook
```

在浏览器中打开 Storybook，找到"协同工作簿"故事。

### 3. 使用流程

#### 创建新工作簿

1. 点击"创建新工作簿"按钮
2. 系统会自动生成一个 6 位分享码
3. 自动连接到新创建的工作簿

#### 加入现有工作簿

1. 在输入框中输入分享码
2. 点击"连接工作簿"按钮
3. 成功连接后可以看到工作簿内容

#### 导入文件

1. 连接到工作簿后，点击"导入文件"按钮
2. 选择 Excel 或 CSV 文件
3. 文件内容会替换当前工作簿的内容
4. 所有连接的用户都会看到更新

#### 导出文件

1. 点击"导出 Excel"或"导出 CSV"按钮
2. 文件会自动下载到本地

#### 协同编辑

1. 多个用户连接到同一个分享码
2. 可以看到其他用户的在线状态
3. 实时同步编辑操作
4. 不同用户有不同的光标颜色

## API 接口

### 创建工作簿

```
POST /create
Response: { success: true, shareCode: "ABC123" }
```

### 获取工作簿数据

```
GET /workbook/:shareCode
Response: { success: true, data: [...], shareCode: "ABC123" }
```

### 导入文件

```
POST /import/:shareCode
Body: { sheets: [...] }
Response: { success: true, message: "文件导入成功" }
```

### 导出文件

```
POST /export/:shareCode
Body: { format: "xlsx", fileName: "导出文件" }
Response: { success: true, fileUrl: "/exports/file.xlsx" }
```

## WebSocket 协议

### 加入房间

```json
{ "req": "join", "shareCode": "ABC123" }
```

### 操作同步

```json
{ "req": "op", "data": [...] }
```

### 用户状态

```json
{
  "req": "addPresences",
  "data": [
    {
      "userId": "...",
      "username": "...",
      "color": "...",
      "selection": { "r": 0, "c": 0 }
    }
  ]
}
```

## 数据库结构

每个工作表文档包含以下字段：

- `shareCode`: 分享码
- `name`: 工作表名称
- `celldata`: 单元格数据
- `order`: 工作表顺序
- `row`: 行数
- `column`: 列数
- `config`: 配置信息
- `createdAt`: 创建时间
- `updatedAt`: 更新时间

## 注意事项

1. 确保 MongoDB 服务正在运行
2. 分享码区分大小写
3. 导入文件会替换现有数据，请谨慎操作
4. 建议在生产环境中添加用户认证和权限控制
5. 大文件导入可能需要较长时间，请耐心等待

## 技术栈

- **前端**: React + TypeScript + Fortune Sheet
- **后端**: Node.js + Express + WebSocket
- **数据库**: MongoDB
- **文件处理**: SheetJS (xlsx)
- **实时通信**: WebSocket
