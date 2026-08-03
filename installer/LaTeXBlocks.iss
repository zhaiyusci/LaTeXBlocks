#ifndef MyAppVersion
  #define MyAppVersion "0.2.14"
#endif
#ifndef SourceDir
  #error SourceDir must point to the ClickOnce publish directory
#endif
#ifndef StemTeXSourceDir
  #error StemTeXSourceDir must point to a staged StemTeX distribution
#endif
#ifndef PowerPointSourceDir
  #error PowerPointSourceDir must point to the PowerPoint ClickOnce publish directory
#endif
#ifndef CertPath
  #error CertPath must point to the public publisher certificate
#endif
#ifndef VcRedistPath
  #error VcRedistPath must point to the VC++ x64 redistributable
#endif
#ifndef VcMajor
  #error VcMajor must specify the bundled VC++ runtime version
#endif
#ifndef VcMinor
  #error VcMinor must specify the bundled VC++ runtime version
#endif
#ifndef VcBuild
  #error VcBuild must specify the bundled VC++ runtime version
#endif
#ifndef VcRevision
  #error VcRevision must specify the bundled VC++ runtime version
#endif
#ifndef OutputDir
  #define OutputDir "."
#endif

[Setup]
AppId={{B92AF36F-487D-4656-91A8-0FD09B8BD4A1}
AppName=LaTeX Blocks
AppVersion={#MyAppVersion}
AppPublisher=Y. Zhai
AppPublisherURL=https://github.com/zhaiyusci/LaTeXBlocks
AppSupportURL=https://github.com/zhaiyusci/LaTeXBlocks/issues
DefaultDirName={localappdata}\Programs\LaTeX Blocks
DisableDirPage=auto
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=LaTeXBlocks-Setup-{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes
CloseApplications=yes
CloseApplicationsFilter=WINWORD.EXE,POWERPNT.EXE
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany=Y. Zhai
VersionInfoDescription=LaTeX Blocks for Microsoft Office
VersionInfoProductName=LaTeX Blocks
VersionInfoProductVersion={#MyAppVersion}

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#PowerPointSourceDir}\*"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#StemTeXSourceDir}\runtime\*"; DestDir: "{app}\StemTeX\runtime"; Excludes: "texmf-var\fonts\cache\*,*.aux,*.log,*.xdv,*.pdf,*.synctex.gz"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#StemTeXSourceDir}\gui\profiles\*"; DestDir: "{app}\StemTeX\gui\profiles"; Excludes: "*.aux,*.log,*.xdv,*.pdf,*.synctex.gz"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#CertPath}"; DestDir: "{tmp}"; DestName: "LaTeXBlocks-publisher.cer"; Flags: ignoreversion deleteafterinstall
Source: "{#VcRedistPath}"; DestDir: "{tmp}"; DestName: "vc_redist.x64.exe"; Flags: ignoreversion deleteafterinstall

[InstallDelete]
Type: filesandordirs; Name: "{app}\StemTeX\gui\profiles\arial_lmodern_simhei"
Type: filesandordirs; Name: "{app}\StemTeX\gui\profiles\arial_lete_yahei"

[Registry]
Root: HKCU; Subkey: "Software\LaTeXBlocks"; ValueType: string; ValueName: "StemTeXHome"; ValueData: "{app}\StemTeX"; Flags: uninsdeletevalue uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "LaTeX Blocks"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn"; ValueType: string; ValueName: "Description"; ValueData: "Editable, searchable LaTeX blocks rendered by StemTeX"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn"; ValueType: string; ValueName: "FriendlyName"; ValueData: "LaTeX Blocks"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn"; ValueType: string; ValueName: "Description"; ValueData: "Editable, searchable LaTeX blocks for slides rendered by StemTeX"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetPowerPointManifestUri}"; Flags: uninsdeletekey

[Run]
Filename: "{sys}\certutil.exe"; Parameters: "-user -f -addstore Root ""{tmp}\LaTeXBlocks-publisher.cer"""; StatusMsg: "Trusting the LaTeX Blocks development publisher..."; Flags: runhidden waituntilterminated
Filename: "{sys}\certutil.exe"; Parameters: "-user -f -addstore TrustedPublisher ""{tmp}\LaTeXBlocks-publisher.cer"""; StatusMsg: "Registering the LaTeX Blocks publisher..."; Flags: runhidden waituntilterminated
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/install /quiet /norestart"; Verb: "runas"; StatusMsg: "Installing the Microsoft Visual C++ x64 runtime..."; Check: NeedsVCRuntime; Flags: shellexec waituntilterminated
Filename: "{app}\setup.exe"; Parameters: "/q /norestart"; WorkingDir: "{app}"; StatusMsg: "Installing required Microsoft components..."; Check: NeedsBootstrapper; Flags: waituntilterminated
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Uninstall ""{code:GetManifestUri}"" /Silent"; StatusMsg: "Removing an earlier LaTeX Blocks registration..."; Flags: runhidden waituntilterminated
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Install ""{code:GetManifestUri}"" /Silent"; StatusMsg: "Registering LaTeX Blocks with Microsoft Word..."; Flags: runhidden waituntilterminated; AfterInstall: RegisterInstalledWordManifest
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Uninstall ""{code:GetPowerPointManifestUri}"" /Silent"; StatusMsg: "Removing an earlier PowerPoint registration..."; Flags: runhidden waituntilterminated
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Install ""{code:GetPowerPointManifestUri}"" /Silent"; StatusMsg: "Registering LaTeX Blocks with Microsoft PowerPoint..."; Flags: runhidden waituntilterminated; AfterInstall: RegisterInstalledPowerPointManifest

[UninstallRun]
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Uninstall ""{code:GetManifestUri}"" /Silent"; RunOnceId: "UninstallVstoSolution"; Flags: runhidden waituntilterminated skipifdoesntexist
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/Uninstall ""{code:GetPowerPointManifestUri}"" /Silent"; RunOnceId: "UninstallPowerPointVstoSolution"; Flags: runhidden waituntilterminated skipifdoesntexist

[Code]
function InitializeSetup: Boolean;
var
  OfficePlatform: String;
begin
  Result := True;
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
       'Platform', OfficePlatform) and (CompareText(OfficePlatform, 'x86') = 0) then begin
    MsgBox('LaTeX Blocks requires 64-bit Microsoft Office because StemTeX is an x64 native runtime.',
      mbError, MB_OK);
    Result := False;
  end;
end;

function GetManifestUri(Param: String): String;
var
  ManifestPath: String;
begin
  ManifestPath := ExpandConstant('{app}\LaTeXBlocks.Word.AddIn.vsto');
  StringChangeEx(ManifestPath, '\', '/', True);
  StringChangeEx(ManifestPath, ' ', '%20', True);
  Result := 'file:///' + ManifestPath;
end;

function GetPowerPointManifestUri(Param: String): String;
var
  ManifestPath: String;
begin
  ManifestPath := ExpandConstant('{app}\PowerPoint\LaTeXBlocks.PowerPoint.AddIn.vsto');
  StringChangeEx(ManifestPath, '\', '/', True);
  StringChangeEx(ManifestPath, ' ', '%20', True);
  Result := 'file:///' + ManifestPath;
end;

procedure RegisterInstalledWordManifest;
var
  AddInKey: String;
begin
  { VSTOInstaller can retain a Visual Studio |vstolocal registration when the
    same solution identity was previously run from a development build. The
    installer owns the final registration and must always point Word at its
    self-contained installed files. }
  AddInKey := 'Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn';
  if not RegWriteStringValue(HKCU, AddInKey, 'Manifest', GetManifestUri('')) then
    RaiseException('Could not register the installed LaTeX Blocks manifest.');
  if not RegWriteDWordValue(HKCU, AddInKey, 'LoadBehavior', 3) then
    RaiseException('Could not enable the installed LaTeX Blocks add-in.');
end;

procedure RegisterInstalledPowerPointManifest;
var
  AddInKey: String;
begin
  AddInKey := 'Software\Microsoft\Office\PowerPoint\Addins\LaTeXBlocks.PowerPoint.AddIn';
  if not RegWriteStringValue(HKCU, AddInKey, 'Manifest', GetPowerPointManifestUri('')) then
    RaiseException('Could not register the installed PowerPoint LaTeX Blocks manifest.');
  if not RegWriteDWordValue(HKCU, AddInKey, 'LoadBehavior', 3) then
    RaiseException('Could not enable the installed PowerPoint LaTeX Blocks add-in.');
end;

function GetVstoInstallerPath(Param: String): String;
var
  Candidate: String;
begin
  if IsWin64 then begin
    Candidate := ExpandConstant('{commoncf64}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
    if FileExists(Candidate) then begin
      Result := Candidate;
      exit;
    end;
  end;
  Result := ExpandConstant('{commoncf32}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
end;

function NeedsBootstrapper: Boolean;
var
  DotNetRelease: Cardinal;
  HasDotNet48: Boolean;
  HasVstoRuntime: Boolean;
begin
  if IsWin64 then
    HasDotNet48 := RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', DotNetRelease)
  else
    HasDotNet48 := RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', DotNetRelease);
  HasDotNet48 := HasDotNet48 and (DotNetRelease >= 528040);
  HasVstoRuntime := FileExists(GetVstoInstallerPath(''));
  Result := (not HasDotNet48) or (not HasVstoRuntime);
end;

function NeedsVCRuntime: Boolean;
var
  Installed: Cardinal;
  Major: Cardinal;
  Minor: Cardinal;
  Build: Cardinal;
  Revision: Cardinal;
begin
  if not RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
       'Installed', Installed) or (Installed <> 1) or
     not RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
       'Major', Major) or
     not RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
       'Minor', Minor) or
     not RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
       'Bld', Build) or
     not RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
       'Rbld', Revision) then begin
    Result := True;
    exit;
  end;
  Result := (Major < {#VcMajor}) or
    ((Major = {#VcMajor}) and (Minor < {#VcMinor})) or
    ((Major = {#VcMajor}) and (Minor = {#VcMinor}) and (Build < {#VcBuild})) or
    ((Major = {#VcMajor}) and (Minor = {#VcMinor}) and (Build = {#VcBuild}) and
      (Revision < {#VcRevision}));
end;
