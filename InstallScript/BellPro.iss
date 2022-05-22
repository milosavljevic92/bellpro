#define MyAppName "BellPro"
#define MyAppVersion "3.2"
#define MyAppPublisher "Tecomatic"
#define MyAppURL "http://tecomatic.rs"
#define MyAppExeName "BellPro.exe"
[Setup]
AppId={{53BA1E13-DC8C-4FBD-A955-802F4953DCD6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={commonpf32}\{#MyAppName}
DefaultGroupName=Bell Pro
AllowNoIcons=yes
OutputDir=.\setup\
OutputBaseFilename=BellPro_setup_x86
Compression=lzma
SolidCompression=yes  
LicenseFile=license.txt

[CustomMessages]
LaunchProgram=Start BellPro after finishing installation

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]

[Dirs]
Name: "{app}\base"

[Files]
Source: ".\vbRuntime\vbrun60.exe"; Flags: dontcopy
Source: ".\support\MDAC_TYP.EXE"; Flags: dontcopy
Source: ".\base\base.mdb"; DestDir: "{app}\base\"; Flags: ignoreversion recursesubdirs
Source: ".\dll\mscomm32.ocx"; DestDir: "{commonpf32}\{#MyAppName}"; Flags: restartreplace ignoreversion regserver 32bit
Source: ".\dll\msdatgrd.ocx"; DestDir: "{commonpf32}\{#MyAppName}"; Flags: restartreplace ignoreversion regserver 32bit
Source: ".\BellPro.exe"; DestDir: "{commonpf32}\{#MyAppName}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[RUN]
Filename: {app}\{#MyAppExeName}; Description: {cm:LaunchProgram}; Flags: nowait postinstall skipifsilent

[Code]
function IsRuntimeInstalled: Boolean;
begin
  Result := False;
end;

function PrepareToInstall(var NeedsRestart: Boolean): string;
var
  ExitCode: Integer;
begin
  // if the runtime is not already installed
  if not IsRuntimeInstalled then
  begin
    // extract the redist to the temporary folder
    ExtractTemporaryFile('vbrun60.exe');
    // run the redist from the temp folder; if that fails, return from this handler the error text
    if not Exec(ExpandConstant('{tmp}\vbrun60.exe'), '', '', SW_SHOW, ewWaitUntilTerminated, ExitCode) then
    begin
      // return the error text
      Result := 'Setup failed to install Visual Basic runtime. Exit code: ' + IntToStr(ExitCode);
      // exit this function; this makes sense only if there are further prerequisites to install; in this
      // particular example it does nothing because the function exits anyway, so it is pointless here
      Exit;
    end;
  end;
   
   if not IsRuntimeInstalled then
  begin
    // extract the redist to the temporary folder
    ExtractTemporaryFile('MDAC_TYP.EXE');
    // run the redist from the temp folder; if that fails, return from this handler the error text
    if not Exec(ExpandConstant('{tmp}\MDAC_TYP.EXE'), '', '', SW_SHOW, ewWaitUntilTerminated, ExitCode) then
    begin
      // return the error text
      Result := 'Setup failed to install MDAC_TYP.EXE. Exit code: ' + IntToStr(ExitCode);
      // exit this function; this makes sense only if there are further prerequisites to install; in this
      // particular example it does nothing because the function exits anyway, so it is pointless here
      Exit;
    end;
  end;
  end;
[UninstallRun]
//Filename: {sys}\sc.exe; Parameters: "ServiceName" ; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{commonpf32}\{#MyAppName}"
