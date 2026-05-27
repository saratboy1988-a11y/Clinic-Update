; Inno Setup Script for Hospital Manager
; កម្មវិីគ្រប់គ្រងគ្លីនិក (Clinic Management System)
; Author: នូរ សារ៉ាត់ (NOU SARAT)

#define Unicode
#define MyAppName "Clinic Management System"
#define MyAppVersion "2.0.18"
#define MyAppPublisher "NOU SARAT"
#define MyAppExeName "ClinicManager.exe"
#define MyAppURL "https://t.me/nousarat"

[Setup]
; ការកំណត់មូលដ្ឋាន
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=
OutputDir=Installer
OutputBaseFilename=ClinicManager_Setup_{#MyAppVersion}
SetupIconFile=healthcare.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

; ការកំណត់ភាសា
[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
; Name: "khmer"; MessagesFile: "compiler:Languages\Khmer.isl"  ; Khmer language not available, using English only

; ការកំណត់ទីតាំងដាក់ឯកសារ
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

; ឯកសារដែលត្រូវដាក់បញ្ចូល
[Files]
; កម្មវិធីគោល
Source: "dist\ClinicManager\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Icon file (សម្រាប់ Desktop Shortcut)
Source: "healthcare.ico"; DestDir: "{app}"; Flags: ignoreversion

; ឯកសារបម្រុង (បើមាន) - Comment ចោលសិន ព្រោះ clinic.db អាចនឹងមិនមាន
; Source: "dist\clinic.db"; DestDir: "{app}"; Flags: ignoreversion uninsneveruninstall

; ឯកសារផ្សេងៗ
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

; ចំណាំ: កុំដាក់ dist\* ព្រោះវានឹងយក files មិនចាំបាច់ផងដែរ

; បង្កើត Shortcuts
[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\healthcare.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\healthcare.ico"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon; IconFilename: "{app}\healthcare.ico"

; ការងារពេលចាប់ផ្តើម
[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

; ការកំណត់មុនពេលដំឡើង
[Code]
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // ពិនិត្យមើលថាតើកម្មវិធីត្រូវបានដំឡើងរួចឬនៅ
  if RegKeyExists(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1') then
  begin
    if MsgBox('កម្មវិធីនេះត្រូវបានដំឡើងរួចហើយ។ តើអ្នកចង់ដំឡើងម្តងទៀតទេ?', mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
      Exit;
    end;
  end;
  
  Result := True;
end;

// បង្ហាញសារស្វាគមន៍
procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpWelcome then
  begin
    MsgBox('សូមស្វាគមន៍មកកាន់កម្មវិធីដំឡើង ' + '{#MyAppName}' + '!' + #13#13 + 
           'កម្មវិធីនេះនឹងដំឡើងកម្មវិធីគ្រប់គ្រងគ្លីនិកក្នុងកុំព្យូទ័ររបស់អ្នក។' + #13#13 + 
           'ចុច Next ដើម្បីបន្ត។', mbInformation, MB_OK);
  end;
end;

// បង្ហាញសារពេលដំឡើងចប់
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('ការដំឡើងបានជោគជ័យ!' + #13#13 + 
           'សូមអរគុណដែលបានប្រើប្រាស់កម្មវិធីរបស់យើង។' + #13#13 + 
           'បញ្ហាឬសំណួរ សូមទាក់ទង: @nousarat', mbInformation, MB_OK);
  end;
end;

[UninstallDelete]
; លុបឯកសារពេល Uninstall
Type: filesandordirs; Name: "{app}\TelegramShares"
Type: files; Name: "{app}\clinic.db"
Type: files; Name: "{app}\*.log"
