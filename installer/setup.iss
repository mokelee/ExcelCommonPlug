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

; === 版本文件 ===
Source: "version.txt"; DestDir: "{app}"; DestName: ".version"; Flags: ignoreversion

; === 图标文件 ===
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

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

function InitializeSetup(): Boolean;
var
  InstalledVersion: AnsiString;
  VersionFile: String;
begin
  Result := True;
  if IsAppRunning('EXCEL.EXE') then
  begin
    MsgBox('检测到 Excel 正在运行，请先关闭 Excel 再安装。', mbError, MB_OK);
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
