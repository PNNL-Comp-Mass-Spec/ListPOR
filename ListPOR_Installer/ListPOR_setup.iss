; This is an Inno Setup configuration file
; http://www.jrsoftware.org/isinfo.php

#define ApplicationVersion GetFileVersion('..\bin\ListPOR.exe')

[CustomMessages]
AppName=ListPOR

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.
; Example with multiple lines:
; WelcomeLabel2=Welcome message%n%nAdditional sentence

[Files]
Source: ..\bin\ListPOR.exe                           ; DestDir: {app}
Source: ..\bin\ListPOR.pdb                           ; DestDir: {app}
Source: ..\bin\MathNet.Numerics.dll                  ; DestDir: {app}
Source: ..\bin\PRISM.dll                             ; DestDir: {app}
Source: ..\bin\ListPORSettings.xml                   ; DestDir: {app}
                                       
Source: ..\README.md                                 ; DestDir: {app}
Source: ..\RevisionHistory.txt                       ; DestDir: {app}
Source: ..\Docs\TestData.txt                         ; DestDir: {app}
Source: ..\Docs\TestData_filtered.txt                ; DestDir: {app}
Source: ..\Docs\TestData_NoCategories.txt            ; DestDir: {app}
Source: ..\Docs\TestData_NoCategories_filtered.txt   ; DestDir: {app}

[Dirs]
Name: {commonappdata}\ListPOR; Flags: uninsalwaysuninstall

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
; Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked

[Icons]
Name: {commondesktop}\ListPOR; Filename: {app}\ListPOR.exe; Tasks: desktopicon; Comment: ListPOR
Name: {group}\ListPOR;         Filename: {app}\ListPOR.exe; Comment: ListPOR

[Setup]
AppName=ListPOR
AppVersion={#ApplicationVersion}
;AppVerName=ProteinDigestionSimulator
AppID=ListPORId
AppPublisher=Pacific Northwest National Laboratory
AppPublisherURL=http://omics.pnl.gov/software
AppSupportURL=http://omics.pnl.gov/software
AppUpdatesURL=http://omics.pnl.gov/software
DefaultDirName={pf}\ListPOR
DefaultGroupName=PAST Toolkit
AppCopyright=© PNNL
;LicenseFile=.\License.rtf
PrivilegesRequired=poweruser
OutputBaseFilename=ListPOR_Installer
VersionInfoVersion={#ApplicationVersion}
VersionInfoCompany=PNNL
VersionInfoDescription=ListPOR
VersionInfoCopyright=PNNL
DisableFinishedPage=true
ShowLanguageDialog=no
ChangesAssociations=false
EnableDirDoesntExistWarning=false
AlwaysShowDirOnReadyPage=true
;UninstallDisplayIcon={app}\delete_16x.ico
ShowTasksTreeLines=true
OutputDir=.\Output

[Registry]
;Root: HKCR; Subkey: MyAppFile; ValueType: string; ValueName: ; ValueDataMyApp File; Flags: uninsdeletekey
;Root: HKCR; Subkey: MyAppSetting\DefaultIcon; ValueType: string; ValueData: {app}\wand.ico,0; Flags: uninsdeletevalue

[UninstallDelete]
Name: {app}; Type: filesandordirs
