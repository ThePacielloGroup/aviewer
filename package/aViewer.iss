#define MyAppName "Accessibility Viewer"
#define MyAppShortName "aViewer"
#define MyAppVersion "2.0"
#define MyAppPublisher "The Paciello Group"
#define MyAppURL "https://www.paciellogroup.com/"
#define MyAppExeName "aViewer.exe"

[Setup]
AppId={{21E73CD4-C351-4D38-9C85-CFCA7FA2514A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename={#MyAppShortName}{#MyAppVersion}
Compression=lzma
SolidCompression=yes
ShowUndisplayableLanguages=yes
ShowLanguageDialog=yes
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "Default"; MessagesFile: "compiler:Default.isl"
;Name: "Chinese"; MessagesFile: "compiler:Languages\Chinese.isl"
;Name: "Czech"; MessagesFile: "compiler:Languages\Czech.isl"
;Name: "Dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
;Name: "French"; MessagesFile: "compiler:Languages\French.isl"
;Name: "German"; MessagesFile: "compiler:Languages\German.isl"
;Name: "Hindi"; MessagesFile: "compiler:Languages\Hindi.islu"
;Name: "Italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "Japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
;Name: "Korean"; MessagesFile: "compiler:Languages\Korean.isl"
;Name: "Russian"; MessagesFile: "compiler:Languages\Russian.isl"
;Name: "Spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
; Remember to install 3rd party ChineseSimplified (and rename to Chinese), Hindi and Korean language files from http://www.jrsoftware.org/files/istrans/ and use Unicode InnoSetup

[Files]
Source: "./aViewer64bit.exe"; DestDir: "{pf}\{#MyAppName}"; DestName: "{#MyAppExeName}"; Check: Is64BitInstallMode; Flags: ignoreversion
Source: "./aViewer32bit.exe"; DestDir: "{pf}\{#MyAppName}"; DestName: "{#MyAppExeName}"; Check: not Is64BitInstallMode; Flags: ignoreversion
Source: "./lang\*.ini"; DestDir: "{pf}\{#MyAppName}\lang"; Flags: ignoreversion
Source: "./IAccessible2Proxy64bit.dll"; DestDir: "{pf}\{#MyAppName}"; DestName: "IAccessible2Proxy.dll"; Check: Is64BitInstallMode; Flags: ignoreversion
Source: "./IAccessible2Proxy32bit.dll"; DestDir: "{pf}\{#MyAppName}"; DestName: "IAccessible2Proxy.dll"; Check: not Is64BitInstallMode; Flags: ignoreversion

[Dirs]
Name: "{pf}\{#MyAppName}"; Permissions: users-readexec;
Name: "{pf}\{#MyAppName}\lang"; Permissions: users-readexec;

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[UninstallDelete]
Type: files; Name: "{pf}\{#MyAppName}\*.*";
Type: filesandordirs; Name: "{pf}\{#MyAppName}";

[Code]
const
  CSIDL_PERSONAL = $0005;
procedure CurStepChanged(CurStep: TSetupStep);
var
  lang, ini, path: string;
begin
  if CurStep = ssDone then
  begin
    path := GetShellFolderByCSIDL(CSIDL_PERSONAL, False);
    if path <> '' then
    begin
      lang := ActiveLanguage + '.ini';
      ini := path + '\MSAAV.ini';
      SetIniString('Settings', 'LangFile', lang, ini);
    end;
  end;
end;