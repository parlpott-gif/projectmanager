# 硬件项目管理系统

Web 应用版本 - Flask + HTML/CSS/JS

## 快速开始

### 前置条件

```bash
pip install flask xlrd xlwt xlutils pandas python-docx
```

### 启动服务
```bash
python app_main.py
```

输出：
```
[INFO] 硬件项目管理系统已启动
[INFO] 访问地址: http://127.0.0.1:5000
[INFO] 按 Ctrl+C 停止服务器
```

### 访问应用
在浏览器中打开：**http://127.0.0.1:5000**

## 项目结构

```
project_manager/
├── app_main.py           # 入口点，启动 Flask 服务器
├── server.py             # Flask 后端，REST API
├── templates/            # HTML 模板
│   └── index.html        # 主页
├── static/               # 静态资源
│   ├── app.js            # 前端逻辑
│   ├── style.css         # 样式表
│   └── bootstrap.min.css  # UI 框架
├── data/                 # 数据文件
│   ├── projects.json     # 项目列表和商务信息
│   └── parts.json        # 器件库
└── 总表.xls              # 项目硬件设计表
```

## 功能

### 项目管理
- ✅ 查看 26 个硬件项目
- ✅ 查看项目硬件成本（从 Excel 读取）
- ✅ 编辑项目商务信息（status、price、deliverable 等）
- ✅ 查看项目文件列表

### 器件库
- ✅ 查看 79 种硬件器件
- ✅ 查看库存数量和采购记录

### 数据统计
- ✅ 项目状态分布
- ✅ 总体商务统计

## API 端点

| 方法 | 端点 | 说明 |
|------|------|------|
| GET | `/api/projects` | 获取所有项目 |
| PUT | `/api/projects/<name>` | 更新项目信息 |
| GET | `/api/parts` | 获取器件库 |
| GET | `/api/stats` | 获取统计数据 |

## 数据来源

- **项目列表**：`data/projects.json`（本地 JSON）
- **硬件信息**：`总表.xls`（Excel 文件）
- **器件库**：`data/parts.json`（本地 JSON）

项目硬件成本自动从 Excel 读取，商务信息保存在 JSON。

## 开发

修改代码后，只需重启 `app_main.py` 即可生效。

### 前端修改
- `templates/index.html` - HTML 结构
- `static/style.css` - 样式
- `static/app.js` - 交互逻辑

### 后端修改
- `server.py` - Flask 路由和数据处理

## 停止服务
在终端按 **Ctrl+C** 停止服务器。
