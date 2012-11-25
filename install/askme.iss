; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{416C0CEA-20A0-4458-8718-B775D968EDF9}
AppName=Askme
AppVerName=Askme 1.0.116.942
AppPublisher=Alliance Soft
AppPublisherURL=http://www.alliance-soft.com/
AppSupportURL=http://www.alliance-soft.com/
AppUpdatesURL=http://www.alliance-soft.com/
DefaultDirName={pf}\Askme
DefaultGroupName=Askme
OutputDir=D:\datadi\programming\src\Askme\install
OutputBaseFilename=setup
SetupIconFile=D:\datadi\programming\src\Askme\Go.ico
Compression=lzma
SolidCompression=true

[Languages]
Name: english; MessagesFile: compiler:Default.isl
Name: russian; MessagesFile: compiler:Languages\Russian.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Files]
Source: ISAlliance.dll; Flags: dontcopy
Source: ISAlliance.dll; Flags: dontcopy
Source: D:\datadi\programming\src\Askme\exe\Askme.exe; DestDir: {app}; Flags: ignoreversion
Source: D:\datadi\programming\src\Askme\exe\Askme.chm; DestDir: {app}; Flags: ignoreversion
Source: D:\datadi\programming\src\Askme\exe\ICSharpCode.SharpZipLib.dll; DestDir: {app}; Flags: ignoreversion
Source: ..\exe\sysprop.dmp; DestDir: {app}; Flags: onlyifdoesntexist
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: ..\exe\office\Codes_base.xlsx; DestDir: {app}\office; Flags: onlyifdoesntexist
Source: ..\exe\convert\ChkCountChanel.xslt; DestDir: {app}\convert; Flags: ignoreversion
Source: ..\exe\convert\extractXlsxData.xslt; DestDir: {app}\convert; Flags: ignoreversion
Source: ..\exe\convert\ora_xslt.xslt; DestDir: {app}\convert; Flags: ignoreversion
Source: ..\exe\convert\xslt_result.xslt; DestDir: {app}\convert; Flags: ignoreversion
Source: ..\exe\convert\xslToRow.xslt; DestDir: {app}\convert; Flags: ignoreversion

[Icons]
Name: {group}\Askme; Filename: {app}\Askme.exe
Name: {group}\{cm:UninstallProgram,Askme}; Filename: {uninstallexe}
Name: {commondesktop}\Askme; Filename: {app}\Askme.exe; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\Askme; Filename: {app}\Askme.exe; Tasks: quicklaunchicon

[Run]
Filename: {app}\Askme.exe; Description: {cm:LaunchProgram,Askme}; Flags: nowait postinstall skipifsilent
[Dirs]
Name: {app}\resultFile; Flags: uninsalwaysuninstall
[UninstallDelete]
Name: {app}; Type: filesandordirs
[Code]
procedure KillProc(lpProcName: AnsiString); external 'KillProcess@files:ISAlliance.dll stdcall setuponly';

function InitializeSetup(): Boolean;
begin
	KillProc('Askme.exe');
	Result := true;
end;
