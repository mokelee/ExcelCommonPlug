; ============================================================
; 项目配置文件 - 每个 Excel-DNA 插件项目只需修改此文件
; ============================================================

; 应用信息
#define MyAppName        "Excel日常工具"
#define MyAppVersion     "0.6.4"
#define MyAppPublisher   "ExcelCommonTools"
#define MyAppURL         "https://github.com/your-repo/ExcelCommonTools"

; 唯一标识 (每个插件必须不同，用 https://www.guidgenerator.com 生成)
#define MyAppId          "B8E3F9A1-5C2D-4E7F-A6B0-1234567890AB"

; Excel-DNA 加载项基础名称 (不含 .xll 后缀)
; 编译后会生成: {AddInBaseName}.xll (32位) 和 {AddInBaseName}64.xll (64位)
#define AddInBaseName    "ExcelCommonTools-AddIn"

; 主程序集名称 (不含 .dll 后缀)
#define AssemblyName     "ExcelCommonTools"

; 安装目录名称 (在 %APPDATA% 下)
#define InstallDirName   "ExcelCommonTools"

; 编译输出目录 (相对于 installer 目录)
#define BuildOutput      "..\src\bin\Release\net48"

; UpdateClient 路径 (相对于 installer 目录)
#define UpdateClientExe  "..\..\UpdateClient\src\bin\Release\net48\UpdateClient.exe"

; 输出安装包文件名前缀
#define OutputName       "ExcelCommonTools_Setup"
