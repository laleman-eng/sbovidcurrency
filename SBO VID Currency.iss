; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "SBO VID Currency"
#define MyAppVersion "2.1.9"
#define MyAppPublisher "VisualD"
#define MyAppURL "http://www.visuald.cl"
#define MyAppExeName "SBO VID Currency.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{097187ED-A193-4FEA-BA26-E808E408B6AD}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf32}\\SAP\SAP Business One\AddOns\VID\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Files]
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\SBO VID Currency.exe; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\SBO VID Currency.pdb; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\SBO VID Currency.exe.config; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.Core.dll; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.Core.pdb; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.GlobalVid.dll; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.GlobalVid.pdb; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.uEncrypt.dll; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\VisualD.uEncrypt.pdb; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\Newtonsoft.Json.dll; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\Newtonsoft.Json.xml; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\SAP90\Interop.SAPbobsCOM.dll; DestDir: {app}; Flags: ignoreversion
Source: C:\VisualK\SBOVIDCurrency\bin\Debug\Microsoft.VisualBasic.PowerPacks.Vs.dll; DestDir: {app}; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: {group}\{#MyAppName}; Filename: {app}\{#MyAppExeName}
Name: {commondesktop}\{#MyAppName}; Filename: {app}\{#MyAppExeName}; Tasks: desktopicon
