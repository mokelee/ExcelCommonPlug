; ============================================================
; Excel-DNA 插件通用安装脚本
; 
; 使用方法:
;   1. 修改 config.iss 中的项目参数
;   2. 如有额外依赖DLL，在下方 [Files] 段添加
;   3. 运行: iscc setup.iss
;
; 通用特性:
;   - 自动检测 Excel 32/64 位，注册对应 xll
;   - 安装到用户目录，无需管理员权限
;   - 集成 UpdateClient 自动升级
;   - 支持多版本 Office (2010/2013/2016/365)
;   - 中文/英文双语安装界面
; ============================================================

; 加载项目配置
#include "config.iss"
; 加载版本号（由 build.bat 从 CHANGELOG.md 提取后自动生成，请勿手动修改）
#include "version.iss"

[Setup]
AppId={{{#MyAppId}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={userappdata}\{#InstallDirName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
SetupIconFile=app.ico
UninstallDisplayIcon={app}\app.ico
OutputDir=.\output
OutputBaseFilename={#OutputName}_{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
PrivilegesRequired=lowest
RestartIfNeededByRun=no
UninstallFilesDir={app}\uninstall

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; === Excel-DNA 核心文件 (通用，由 config.iss 变量驱动) ===
Source: "{#BuildOutput}\{#AddInBaseName}64.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\{#AddInBaseName}64.dna"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\{#AddInBaseName}.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\{#AddInBaseName}.dna"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\{#AssemblyName}.dll"; DestDir: "{app}"; Flags: ignoreversion

; === 依赖 DLL (按项目需要增减) ===
Source: "{#BuildOutput}\System.Text.Json.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Buffers.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Memory.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Numerics.Vectors.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Runtime.CompilerServices.Unsafe.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Text.Encodings.Web.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.Threading.Tasks.Extensions.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\System.ValueTuple.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Microsoft.Bcl.AsyncInterfaces.dll"; DestDir: "{app}"; Flags: ignoreversion

; === UpdateClient 升级程序 ===
Source: "{#UpdateClientExe}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#UpdateClientExe}.config"; DestDir: "{app}"; DestName: "UpdateClient.exe.config"; Flags: ignoreversion

; === 注册脚本 + 配置 ===
Source: "register.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "unregister.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "addin.ini"; DestDir: "{app}"; Flags: ignoreversion

; === 清理后台 Excel 进程工具（独立 Native AOT exe，开始菜单可执行） ===
; 不依赖 powershell.exe，可在 AppLocker/SRP 锁定环境正常运行。
Source: "{#ExcelCleanerExe}"; DestDir: "{app}"; Flags: ignoreversion
; 额外释放一份到临时目录（dontcopy），供 InitializeSetup 在安装前静默清理时调用
; （此时文件尚未安装到 {app}，需先 ExtractTemporaryFile 才能运行）。
Source: "{#ExcelCleanerExe}"; DestName: "ExcelCleaner.exe"; Flags: dontcopy

; === 版本文件 ===
Source: "version.txt"; DestDir: "{app}"; DestName: ".version"; Flags: ignoreversion

; === 图标文件 ===
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单程序组：清理后台残留的 Excel 进程（策略 A：只清理无窗口的幽灵进程）
; 不指定 IconFilename，使用 ExcelCleaner.exe 内嵌的 clean.ico
Name: "{group}\清理后台 Excel 进程"; Filename: "{app}\ExcelCleaner.exe"; WorkingDir: "{app}"; Comment: "清理后台残留、无窗口的 Excel 进程，解决插件无法加载的问题"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"; IconFilename: "{app}\app.ico"

[Run]
; register.bat 作为备用，主要注册逻辑在 [Code] 的 CurStepChanged 中完成
; Filename: "{cmd}"; Parameters: "/c ""{app}\register.bat"""; Flags: runhidden

[UninstallRun]
Filename: "{cmd}"; Parameters: "/c ""{app}\unregister.bat"""; Flags: runhidden waituntilterminated; RunOnceId: "UnregXll"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\.update_temp"
Type: filesandordirs; Name: "{app}\.update_backup"
Type: filesandordirs; Name: "{app}\.update_logs"
Type: filesandordirs; Name: "{app}\logs"
Type: files; Name: "{app}\.version"
Type: files; Name: "{app}\addin.ini"
Type: filesandordirs; Name: "{app}"

[Code]
function IsAppRunning(const FileName: string): Boolean;
var
  FSWbemLocator: Variant;
  FWMIService: Variant;
  FWbemObjectSet: Variant;
begin
  try
    FSWbemLocator := CreateOleObject('WBEMScripting.SWBEMLocator');
    FWMIService := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');
    FWbemObjectSet := FWMIService.ExecQuery('SELECT Name FROM Win32_Process WHERE Name="' + FileName + '"');
    Result := (FWbemObjectSet.Count > 0);
    FWbemObjectSet := Unassigned;
    FWMIService := Unassigned;
    FSWbemLocator := Unassigned;
  except
    Result := False;
  end;
end;

// 静默清理后台残留的 Excel 幽灵进程（策略 A：只结束「没有主窗口」的 EXCEL.EXE）。
// 调用独立的 ExcelCleaner.exe --silent（Native AOT，不依赖 powershell.exe，
// 可在 AppLocker/SRP 锁定环境运行）。返回 ExcelCleaner 的退出码：
//   0  = 已清理干净，无正在使用的 Excel
//   10 = 仍有正在使用的 Excel 窗口（有主窗口），未关闭
//   11 = 有幽灵进程因权限不足无法结束（拒绝访问）
//   其它/异常 = -1（视为清理未成功）
function CleanGhostExcel(): Integer;
var
  ResultCode: Integer;
  ExePath: String;
begin
  Result := -1;
  try
    // InitializeSetup 阶段文件尚未安装到 {app}，先释放临时副本再运行。
    ExtractTemporaryFile('ExcelCleaner.exe');
    ExePath := ExpandConstant('{tmp}\ExcelCleaner.exe');
    if Exec(ExePath, '--silent', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      Result := ResultCode;
  except
    // 释放/运行失败：返回 -1，交由下方 IsAppRunning 兜底判断。
  end;
end;

function InitializeSetup(): Boolean;
var
  InstalledVersion: AnsiString;
  VersionFile: String;
  CleanCode: Integer;
begin
  Result := True;

  // 安装前先静默清理后台幽灵 Excel 进程，避免被用户看不见、关不掉的残留进程卡住。
  if IsAppRunning('EXCEL.EXE') then
  begin
    CleanCode := CleanGhostExcel();

    // 退出码 10：清理后仍有正在使用的 Excel 窗口，不能擅自结束（会丢数据），
    // 提示用户手动关闭后再安装。
    if CleanCode = 10 then
    begin
      MsgBox('检测到 Excel 正在使用中，请先保存并关闭所有 Excel 窗口后再安装。', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    // 退出码 11：有幽灵进程因权限不足无法结束，提示以管理员身份重试。
    if CleanCode = 11 then
    begin
      MsgBox('检测到后台残留的 Excel 进程需要更高权限才能清理。' + #13#10 + #13#10 +
        '请从开始菜单以管理员身份运行「清理后台 Excel 进程」后，再重新安装。', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    Sleep(300);
  end;

  // 兜底：清理工具异常未成功（返回 -1）且仍检测到 EXCEL.EXE 时，提示用户手动关闭。
  if IsAppRunning('EXCEL.EXE') then
  begin
    MsgBox('检测到 Excel 正在运行，请先保存并关闭所有 Excel 窗口后再安装。', mbError, MB_OK);
    Result := False;
    Exit;
  end;

  // 检查是否降级安装
  VersionFile := ExpandConstant('{userappdata}\{#InstallDirName}\.version');
  if FileExists(VersionFile) then
  begin
    if LoadStringFromFile(VersionFile, InstalledVersion) then
    begin
      InstalledVersion := Trim(InstalledVersion);
      if (InstalledVersion <> '') and (CompareStr(InstalledVersion, '{#MyAppVersion}') > 0) then
      begin
        if MsgBox('当前已安装版本 ' + InstalledVersion + '，即将安装的版本 {#MyAppVersion} 更低。' + #13#10 + #13#10 + '确定要降级安装吗？', mbConfirmation, MB_YESNO) = IDNO then
        begin
          Result := False;
        end;
      end;
    end;
  end;
end;

function IsOffice64Bit: Boolean;
var
  Platform: String;
  ExePath: String;
begin
  Result := False;

  // 方法1：ClickToRun 安装（Office 2016+/365，最常见）
  if IsWin64 then
  begin
    if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Platform) then
    begin
      Result := (CompareText(Platform, 'x64') = 0);
      Exit;
    end;
  end;
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Platform) then
  begin
    Result := (CompareText(Platform, 'x64') = 0);
    Exit;
  end;

  // 方法2：通过 Excel.exe 路径判断（适用于 MSI 安装）
  // 64 位 Office 装在 Program Files，32 位装在 Program Files (x86)
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', ExePath) then
  begin
    Result := (Pos('Program Files (x86)', ExePath) = 0) and (Pos('Program Files', ExePath) > 0);
    Exit;
  end;
  if IsWin64 then
  begin
    if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', ExePath) then
    begin
      Result := (Pos('Program Files (x86)', ExePath) = 0) and (Pos('Program Files', ExePath) > 0);
    end;
  end;
end;

procedure RegisterXll;
var
  XllFile: String;
  XllPath: String;
  RegKey: String;
  ValueName: String;
  ExistingValue: String;
  I: Integer;
  Registered: Boolean;
begin
  if IsOffice64Bit then
    XllFile := '{#AddInBaseName}64.xll'
  else
    XllFile := '{#AddInBaseName}.xll';

  XllPath := ExpandConstant('{app}') + '\' + XllFile;
  Registered := False;

  // 为检测到的 Office 版本注册
  if RegKeyExists(HKCU, 'Software\Microsoft\Office\16.0\Excel') then
    RegKey := 'Software\Microsoft\Office\16.0\Excel\Options'
  else if RegKeyExists(HKCU, 'Software\Microsoft\Office\15.0\Excel') then
    RegKey := 'Software\Microsoft\Office\15.0\Excel\Options'
  else if RegKeyExists(HKCU, 'Software\Microsoft\Office\14.0\Excel') then
    RegKey := 'Software\Microsoft\Office\14.0\Excel\Options'
  else
    Exit;

  // 查找已有的注册项（更新）
  for I := 0 to 5 do
  begin
    if I = 0 then ValueName := 'OPEN' else ValueName := 'OPEN' + IntToStr(I);
    if RegQueryStringValue(HKCU, RegKey, ValueName, ExistingValue) then
    begin
      if Pos('{#AddInBaseName}', ExistingValue) > 0 then
      begin
        RegWriteStringValue(HKCU, RegKey, ValueName, '/R "' + XllPath + '"');
        Registered := True;
        Break;
      end;
    end;
  end;

  // 没有已有项，找空位注册
  if not Registered then
  begin
    for I := 0 to 5 do
    begin
      if I = 0 then ValueName := 'OPEN' else ValueName := 'OPEN' + IntToStr(I);
      if not RegValueExists(HKCU, RegKey, ValueName) then
      begin
        RegWriteStringValue(HKCU, RegKey, ValueName, '/R "' + XllPath + '"');
        Break;
      end;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterXll;
end;
