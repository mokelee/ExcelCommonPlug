# ExcelCommonTools 需求文档

## 项目概述

基于 Excel-DNA 开发的 Excel 插件，提供日常办公中的批量文本处理、工作表拆分合并、图片管理、图形批注操作等功能，并内置自动升级机制与 Django 后台服务器对接。

## 技术栈决策

| 层面 | 选型 | 理由 |
|------|------|------|
| 插件框架 | Excel-DNA (.xll) | `dotnet build` 编译，无需 VS；不依赖 VSTO Runtime |
| 目标框架 | .NET Framework 4.8 | 兼容 Office 2016+，企业环境广泛 |
| Excel 操作 | Microsoft.Office.Interop.Excel (COM) | 完整支持 Range/Shape/Chart/Comment |
| Ribbon UI | Excel-DNA CustomUI XML | 直接复用现有 CustomRibbon.xml |
| 对话框 | WinForms | 轻量、无额外依赖 |
| JSON | System.Text.Json | 版本清单解析、与 Django API 通信 |
| 安装包 | Inno Setup | 自动注册 .xll、检测 .NET 4.8 |
| 自动升级 | 内置升级引擎 + 独立 Updater.exe | 静默检测、增量下载、Excel 关闭后替换 |

---

## 功能需求

### FR-1 文本处理

| ID | 功能 | 描述 |
|----|------|------|
| FR-1.1 | 前面添加文本 | 在选中单元格内容前批量插入文本 |
| FR-1.2 | 末尾添加文本 | 在选中单元格内容末尾批量追加文本 |
| FR-1.3 | 中间插入文本 | 在指定位置（从开头/末尾第N个字符）插入文本 |

### FR-2 工作表拆分

| ID | 功能 | 描述 |
|----|------|------|
| FR-2.1 | 按列拆分为Sheet | 按指定列的唯一值将数据拆分为多个工作表 |
| FR-2.2 | 拆分为独立文件 | 将每个工作表另存为独立 .xlsx 文件 |

### FR-3 工作表合并

| ID | 功能 | 描述 |
|----|------|------|
| FR-3.1 | 合并当前工作簿所有Sheet | 将所有工作表数据追加到活动工作表 |
| FR-3.2 | 合并多个文件 | 打开多个 Excel 文件并合并到当前工作簿 |
| FR-3.3 | 跨表汇总(SUM) | 创建汇总表，用跨表SUM公式汇总数据 |

### FR-4 图片处理

| ID | 功能 | 描述 |
|----|------|------|
| FR-4.1 | 调整图片大小 | 将图片缩放适配所在单元格（保持宽高比、居中） |
| FR-4.2 | 批量插入图片 | 根据名称列匹配文件夹中的图片并插入 |
| FR-4.3 | 导出所有图片 | 将工作表中的图片按名称列导出到文件夹 |

### FR-5 工作表管理

| ID | 功能 | 描述 |
|----|------|------|
| FR-5.1 | 批量取消隐藏 | 选择并取消隐藏多个工作表 |
| FR-5.2 | 清理空行空表 | 删除空白工作表、空行、隐藏图形 |
| FR-5.3 | 解除工作表保护 | 移除工作表保护（无密码场景） |

### FR-6 图形与批注

| ID | 功能 | 描述 |
|----|------|------|
| FR-6.1 | 按类型删除图形 | 选择图形类型（图片/文本框/批注等）批量删除 |
| FR-6.2 | 调整批注位置大小 | 批量设置批注的偏移、宽高、可见性 |


### FR-7 自动升级

| ID | 功能 | 描述 |
|----|------|------|
| FR-8.1 | 启动时版本检测 | Excel启动后静默向Django服务器检查新版本 |
| FR-8.2 | 增量下载 | 仅下载变更文件（SHA256比对） |
| FR-8.3 | 全量回退 | 增量失败时自动回退到全量下载 |
| FR-8.4 | Excel关闭后应用 | 通过独立Updater.exe在Excel退出后替换文件 |
| FR-8.5 | 失败静默处理 | 升级失败不影响当前版本正常使用 |
| FR-8.6 | 强制更新提示 | 服务器标记force_update时必须升级才能使用 |

---

## 非功能需求

| ID | 类别 | 描述 |
|----|------|------|
| NFR-1 | 兼容性 | 支持 Office 2016/2019/2021/365（32位和64位） |
| NFR-2 | 编译 | `dotnet build` 一条命令编译，不依赖 Visual Studio |
| NFR-3 | 安装 | 用户双击安装包，自动检测依赖、注册插件 |
| NFR-4 | 性能 | 批量操作时关闭 ScreenUpdating/Calculation |
| NFR-5 | 稳定性 | COM 对象及时释放，避免 Excel 僵尸进程 |
| NFR-6 | 国际化 | UI 使用中文，代码注释中文 |

---

## 自动升级详细设计方案

### 架构概览

```
┌─────────────────────────────────────────────────┐
│                  Excel 进程                       │
│  ┌───────────┐     ┌──────────────────────────┐ │
│  │ AddIn主体  │────▶│  UpdateChecker (后台线程) │ │
│  └───────────┘     └──────────┬───────────────┘ │
│                               │ HTTP GET         │
└───────────────────────────────┼─────────────────┘
                                ▼
                   ┌────────────────────────┐
                   │  Django 升级服务器       │
                   │  /api/v1/products/      │
                   │    └─ {id}/manifest     │
                   │    └─ {id}/files/...    │
                   │    └─ {id}/packages/... │
                   └────────────────────────┘

┌─────────────────────────────────────────────────┐
│            Updater.exe (独立进程)                 │
│  Excel退出后启动 → 替换文件 → 重启Excel          │
└─────────────────────────────────────────────────┘
```

### 升级流程

```
Excel启动
  │
  ├─ 后台线程: GET /api/v1/products/{id}/manifest
  │   ├─ 网络失败 → 静默忽略，正常使用
  │   └─ 成功 → 比较 version
  │       ├─ 无更新 → 结束
  │       └─ 有更新 → 提示用户
  │           ├─ 用户拒绝（非强制） → 结束
  │           └─ 用户确认 / 强制更新
  │               │
  │               ├─ 下载到 .update_temp/
  │               │   ├─ 增量: 逐文件下载(SHA256校验)
  │               │   └─ 增量失败 → 全量zip下载
  │               │
  │               ├─ 下载完成 → 启动 Updater.exe
  │               └─ Excel 退出
  │
  Updater.exe:
    ├─ 等待 Excel 进程退出
    ├─ 备份当前文件到 .update_backup/
    ├─ 从 .update_temp/ 复制到安装目录
    ├─ 写入 .version 文件
    ├─ 清理临时文件
    └─ 重启 Excel
```

### Django 服务器 API 设计

```
GET /api/v1/products/{product_id}/manifest
Response:
{
  "version": "1.2.0",
  "force_update": false,
  "min_version": "1.0.0",
  "release_notes": "修复了XXX问题",
  "files": [
    {"path": "ExcelCommonTools.xll", "hash": "sha256...", "size": 102400},
    {"path": "ExcelCommonTools64.xll", "hash": "sha256...", "size": 108000}
  ],
  "package_url": "/api/v1/products/{id}/packages/1.2.0"
}

GET /api/v1/products/{product_id}/files/{version}/{file_path}
→ 返回单个文件的二进制流

GET /api/v1/products/{product_id}/packages/{version}
→ 返回全量 zip 包
```

### 关键设计决策

| 决策 | 选择 | 理由 |
|------|------|------|
| 何时替换文件 | Excel关闭后 | .xll 文件被 Excel 锁定，运行时无法覆盖 |
| 替换方式 | 独立 Updater.exe | 避免文件锁、可以处理 Excel 崩溃场景 |
| 下载策略 | 增量优先+全量回退 | 减少流量；全量兜底保证可靠性 |
| 哈希算法 | SHA256 | 安全性足够，性能可接受 |
| 失败处理 | 备份+回滚 | 替换失败时恢复上一版本 |
| 检查频率 | 仅启动时一次 | 不干扰用户正常工作 |

### Updater.exe 设计

独立的控制台程序（也用 .NET Framework 4.8 编译），职责：

1. 接收命令行参数：`--temp-dir`、`--install-dir`、`--restart-excel`
2. 等待 Excel.exe 进程退出（最多等待 60 秒）
3. 备份安装目录中的旧文件
4. 从临时目录复制新文件到安装目录
5. 校验文件哈希
6. 成功则清理备份和临时文件；失败则从备份恢复
7. 如指定 `--restart-excel`，启动 Excel

---

## 项目结构（迁移后）

```
ExcelCommonTools/
├── ExcelCommonTools.sln
├── src/
│   └── ExcelCommonTools/
│       ├── ExcelCommonTools.csproj      (.NET Framework 4.8, Excel-DNA)
│       ├── AddIn.cs                     (IExcelAddIn 入口)
│       ├── Ribbon/
│       │   ├── RibbonController.cs      (ComVisible, IRibbonExtensibility)
│       │   └── CustomRibbon.xml         (嵌入资源)
│       ├── Core/
│       │   ├── AppConstants.cs
│       │   ├── ComHelper.cs
│       │   ├── ExcelOperationScope.cs
│       │   └── ServiceLocator.cs
│       ├── Services/                    (业务逻辑，基本不变)
│       │   ├── TextService.cs
│       │   ├── SplitService.cs
│       │   ├── MergeService.cs
│       │   ├── ImageService.cs
│       │   ├── SheetService.cs
│       │   ├── ShapeService.cs
│       │   ├── CommentService.cs
│       │   ├── FormatService.cs
│       │   └── ExcelUtilService.cs
│       ├── UI/                          (WinForms对话框，不变)
│       │   ├── DialogStyleHelper.cs
│       │   ├── InsertTextDialog.cs
│       │   ├── SplitDialog.cs
│       │   └── ...
│       └── Updater/
│           ├── UpdateChecker.cs         (启动时后台检查)
│           ├── DownloadManager.cs
│           ├── HashHelper.cs
│           ├── ManifestParser.cs
│           ├── UpdateConfig.cs
│           └── UpdateState.cs
├── src/
│   └── Updater/
│       ├── Updater.csproj              (独立控制台程序)
│       └── Program.cs
├── installer/
│   └── ExcelCommonTools.iss            (Inno Setup 脚本)
└── README.md
```
