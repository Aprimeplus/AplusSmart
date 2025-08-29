; Inno Setup Script for A+ Smart Solution

; --- START: กำหนดค่าคงที่ทั้งหมดไว้ที่นี่เพื่อง่ายต่อการแก้ไข ---
#define MyAppName "A+Smart"
#define MyAppVersion "1.0"
#define MyAppPublisher "APrime Plus Co., Ltd."
#define MyAppURL "https://www.aprimeplus.com/"
#define MyAppExeName "A+ Smart Solution.exe"
; --- [สำคัญ] แก้ไข Path ตรงนี้ให้เป็นที่อยู่ของโฟลเดอร์โปรแกรมของคุณ (ใช้ \ ตัวเดียวได้) ---
#define MySourcePath "C:\Users\Nitro V15\Desktop\AplusSmart\dist\A+ Smart Solution"
; --- [แนะนำ] นำไฟล์ไอคอนมาไว้ในโฟลเดอร์โปรเจกต์ ---
#define MyAppIcon "C:\Users\Nitro V15\Desktop\AplusSmart\company_logo.ico"
; --- END ---


[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; [แก้ไข] เพิ่มวงเล็บปีกกาที่ขาดไป
AppId={{1363BC07-7ED2-4BED-971A-B45285BE15ED}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
ArchitecturesInstallIn64BitMode=x64compatible
; PrivilegesRequired=lowest ; Uncomment for non-admin install mode
OutputDir=C:\Users\Nitro V15\Downloads
OutputBaseFilename=Installer A+Smart v{#MyAppVersion}
SetupIconFile={#MyAppIcon}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; --- [แก้ไข] รวม Source ให้เหลือบรรทัดเดียวก็เพียงพอ ---
Source: "{#MySourcePath}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent