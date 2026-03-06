# 硬件项目管理系统 - Claude Code 配置

## 项目概述

本地 Web 应用，管理硬件外包项目的商务信息、器件库和文件导出。
Flask 后端 + Bootstrap 前端，单页应用。

**GitHub 仓库**: https://github.com/parlpott-gif/projectmanager

## 目录结构

```
project_manager/
├── server.py              ← Flask 后端（路由 + 业务逻辑）
├── app_main.py            ← 启动入口（python app_main.py）
├── start.bat              ← Windows 一键启动
├── templates/
│   └── index.html         ← 单页应用 HTML
├── static/
│   ├── app.js             ← 前端全部逻辑
│   └── style.css          ← 自定义样式
└── data/                  ← 本地数据（不提交 Git）
    ├── projects.json       ← 商务信息
    ├── parts.json          ← 器件库
    └── export_records.json ← 导出历史
```

**数据源**（在 project_manager 父目录）:
- `../总表.xlsx` — 项目元器件/功能主数据（只读为主）
- `../projects/` — 各项目文件夹（PCB、演示等文件）

## 启动方式

```bash
cd project_manager
python app_main.py
# 浏览器访问 http://127.0.0.1:5000
```

## Git 工作流

每次改完功能后提交：

```bash
git add .
git commit -m "简短描述改了什么"
git push
```

**注意**：`data/` 目录已在 `.gitignore` 中排除，不会提交到 GitHub（包含本地私有数据）。

## 技术栈

| 层次 | 技术 |
|------|------|
| 后端 | Python 3 + Flask |
| 数据持久化 | JSON 文件（projects.json / parts.json）|
| Excel 读取 | openpyxl（.xlsx）/ xlrd（.xls）|
| 文档导出 | python-docx |
| 前端 | Bootstrap 5 + 原生 JS（无框架）|

## 主要 API

| 方法 | 路由 | 说明 |
|------|------|------|
| GET | `/api/projects` | 获取所有项目（Excel + JSON 合并）|
| PUT | `/api/projects/<name>` | 更新项目商务信息 |
| GET | `/api/stats` | 统计概览 |
| POST | `/api/open_folder/<name>` | 弹出 Windows 文件夹窗口 |
| POST | `/api/projects/<name>/upload` | 上传文件到项目目录 |
| GET | `/api/parts` | 获取器件库 |
| POST | `/api/export` | 导出 PCB ZIP 或 DOCX 标签 |

## 开发规范

- **前端改动**：修改 `static/app.js` 或 `static/style.css`，刷新浏览器即可生效
- **后端改动**：修改 `server.py` 后需重启 `app_main.py`
- **不引入新框架**：前端保持原生 JS + Bootstrap，不用 React/Vue
- **HTML 版本号**：`index.html` 中引用 `app.js?v=XX` 和 `style.css?v=XX`，改动后记得更新版本号防止浏览器缓存

## 绝对禁止操作

> 以下操作任何情况下都不允许执行，无论用户是否要求：

- **禁止修改或删除** `data/` 目录下的任何文件（`projects.json`、`parts.json`、`export_records.json`）
- **禁止修改或删除** 父目录的 `总表.xlsx` / `总表.xls`
- **禁止覆盖** 上述文件（包括用写入操作替换内容）

这些是用户的真实业务数据，损坏无法恢复。如果某个任务需要操作这些文件，必须先明确告知用户并等待确认。

---

## 已知注意事项

- `os.startfile()` 打开文件夹仅在 Windows 本地运行有效，云端部署时需移除
- Excel 总表路径在父目录 `../总表.xlsx`，运行目录必须是 `project_manager/`
- `data/` 目录需手动创建，首次运行会自动生成 JSON 文件
