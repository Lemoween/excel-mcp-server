# Excel MCP Server

基于 [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) 的 Excel 文件处理服务，智能体能够通过标准化接口操作 Excel 文件。

目前仅支持**多 Excel / 多 Sheet 的统一内存容器**、**文件合并拆分**等能力。

------

## ✨ 功能特性

- **多 Excel / 多 Sheet 统一容器** —— 惰性加载 + LRU 缓存，重复访问不重复读盘
- **灵活的文件扫描** —— 支持扩展名过滤（`.xlsx` / `.xls`）、文件名通配、递归扫描
- **原子操作** —— 合并、按列拆分、按 sheet 导出、多文件导出
- **一步到位工具** —— `excel_merge_from_paths` 单次调用完成"加载 → 合并 → 导出"
- **智能命名** —— sheet 名 / 文件名自动处理重名、非法字符、超长截断（含 WARNING 日志）
- **显式 Schema** —— 每个工具都提供完整的 `inputSchema`，LLM 选工具、传参更准确
- **中文友好** —— 错误提示、日志、文档均为中文

------

## 📁 目录结构

```
excel-mcp-server/
├── excel_mcp_server/                    # 主包
│   ├── __init__.py                      # 包声明
│   ├── _shared.py                       # 共享工具（Excel 读写、JSON 序列化）
│   ├── excel_bundle.py                  # ExcelBundle：多文件/多 sheet 内存容器
│   ├── bundle_registry.py               # ExcelBundleRegistry：进程级注册表
│   ├── server.py                        # MCP 服务器入口
│   └── tools/                           # 工具模块
│       ├── __init__.py
│       └── document.py                  # DocumentTools：文件级操作工具集
├── excel_mcp.log                        # 运行日志（首次运行自动生成）
└── README.md
```



------

## 🚀 快速开始

### 1. 环境要求

- Python ≥ 3.10（用到 `match / case` 语法）

  

### 2. 启动服务器

**方式 A：模块方式（推荐，从包外目录启动）**

```
cd excel-mcp-server
python -m excel_mcp_server.server
```



**方式 B：脚本方式（从包目录里启动）**

```
cd excel-mcp-server/excel_mcp_server
python server.py
```



> ⚠️ **导入路径说明**：`server.py` 里用的是 `from bundle_registry import REGISTRY`（平级导入），需要**在 `excel_mcp_server/` 目录内运行**；`document.py` 里用的是 `from excel_mcp_server.bundle_registry import ...`（绝对导入），需要**从包外目录运行**。两者取其一，**保持一致性即可**。如果不确定，用方式 A。

启动后它会在 stdio 上等待 MCP 客户端消息，**直接运行没有输出是正常的**。

------

## 🧰 MCP 工具清单

共 **13 个工具**，分为 4 类：

### 生命周期（4）

| 工具名                    | 说明                                         |
| :------------------------ | :------------------------------------------- |
| `excel_load_bundle`       | 加载文件/目录到内存 bundle，返回 `bundle_id` |
| `excel_close_bundle`      | 释放指定 bundle                              |
| `excel_close_all_bundles` | 释放所有 bundle                              |
| `excel_list_bundles`      | 列出当前进程所有 bundle                      |

### 探查（5）

| 工具名                  | 说明                                     |
| :---------------------- | :--------------------------------------- |
| `excel_list_files`      | 列出 bundle 内的文件（含大小、sheet 数） |
| `excel_list_sheets`     | 列出 sheet（可按 file_key 过滤）         |
| `excel_get_bundle_info` | bundle 整体信息（可选带行列数）          |
| `excel_get_columns`     | 获取指定 sheet 的列名与行数              |
| `excel_preview_sheet`   | 预览指定 sheet 的前 n 行                 |

### 合并 / 拆分（3）

| 工具名                       | 说明                               |
| :--------------------------- | :--------------------------------- |
| `excel_merge_files`          | 合并 bundle 内多个 df 为一个新 df  |
| **`excel_merge_from_paths`** | **【一步到位】加载 → 合并 → 导出** |
| `excel_split_by_column`      | 按指定列的唯一值拆分为多个 df      |

### 导出（2）

| 工具名                | 说明                          |
| :-------------------- | :---------------------------- |
| `excel_export_bundle` | 导出为多 sheet 的单文件 Excel |
| `excel_export_split`  | 每个 df 单独导出一个 `.xlsx`  |