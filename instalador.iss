; =============================================================================
;  Servidor de Pedidos — Instalador (Inno Setup 6)
;  Visual moderno + términos obligatorios + conserva de licencia.lic
;  Workflow: .github/workflows/instalador.yml
; =============================================================================

#define MyAppName "Servidor de Pedidos"
#define MyAppVersion "1.0"
#define MyAppPublisher "José Carlos Fernández"
#define MyAppExeName "ServidorPedidos.exe"

[Setup]
AppId={{A8F3C2D1-9E4B-4F6A-8C1D-7B2E5A9F0D33}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright=Copyright © 2026 José Carlos Fernández
; Misma carpeta que instalaciones anteriores (no perder licencia.lic)
DefaultDirName={autopf}\Toto Tools\Servidor de Pedidos
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=LICENCIA_INSTALACION.txt
InfoBeforeFile=
OutputDir=instalador
OutputBaseFilename=Instalar-ServidorPedidos
SetupIconFile=icono.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
LZMAUseSeparateProcess=yes
WizardStyle=modern
WizardSizePercent=110
WizardImageFile=wizard-grande.bmp
WizardSmallImageFile=wizard-peque.bmp
WizardImageStretch=no
DisableWelcomePage=no
ShowLanguageDialog=no
VersionInfoCompany={#MyAppPublisher}
VersionInfoCopyright=Desarrollado por José Carlos Fernández
VersionInfoDescription={#MyAppName} — Instalador
VersionInfoProductName={#MyAppName}
VersionInfoVersion={#MyAppVersion}.0.0
CloseApplications=yes
RestartApplications=no
SetupLogging=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[CustomMessages]
spanish.WelcomeLabel1=Bienvenido a {#MyAppName}
spanish.WelcomeLabel2=Este asistente instalará {#MyAppName} en su equipo.%n%nDesarrollado por José Carlos Fernández.%n%n• Pedidos desde el móvil%n• Catálogo e inventario%n• Impresión térmica%n%nSi ya tenía el programa, se conservarán licencia y datos.%n%nDebe aceptar los términos para continuar.
spanish.FinishedLabel={#MyAppName} se instaló correctamente.%n%nPuede abrir el programa ahora. En el primer uso configure su negocio con el asistente.
spanish.ClickNext=Haga clic en Siguiente para continuar.
spanish.ButtonInstall=Instalar ahora

[Messages]
spanish.BeveledLabel=Desarrollado por José Carlos Fernández
spanish.SetupWindowTitle=Instalación — {#MyAppName}
spanish.WizardWelcome=Bienvenida
spanish.WizardLicense=Términos y privacidad
spanish.WizardSelectDir=Carpeta de instalación
spanish.WizardReady=Listo para instalar
spanish.WizardInstalling=Instalando…
spanish.WizardFinished=Instalación completada
spanish.LicenseLabel=Lea los Términos y condiciones y la Política de privacidad. Debe aceptarlos para instalar el programa.
spanish.LicenseAccepted=Acepto los Términos y la Política de privacidad
spanish.LicenseNotAccepted=No acepto
spanish.SelectDirDesc=Elija dónde instalar {#MyAppName}
spanish.SelectDirLabel3=Se instalará en la siguiente carpeta. Se recomienda dejar la ruta por defecto si ya tenía una versión anterior (así se conserva la licencia).
spanish.ReadyLabel1=El asistente está listo para instalar {#MyAppName} en su equipo.
spanish.FinishedHeadingLabel=¡Todo listo!
spanish.StatusExtractFiles=Copiando archivos…
spanish.StatusCreateIcons=Creando accesos directos…

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Opciones adicionales:"; Flags: checkedonce

[Dirs]
Name: "{app}"; Permissions: users-modify

[Files]
Source: "dist\ServidorPedidos.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "wizard-grande.bmp"; DestDir: "{tmp}"; Flags: dontcopy noencryption
Source: "wizard-peque.bmp"; DestDir: "{tmp}"; Flags: dontcopy noencryption

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Comment: "Abrir el servidor de pedidos"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Comment: "Servidor de Pedidos"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName} ahora"; Flags: nowait postinstall skipifsilent unchecked

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

// Tras instalar: recuperar licencia.lic de carpetas anteriores si hace falta
procedure CurStepChanged(CurStep: TSetupStep);
var
  OldDirs: array of String;
  I: Integer;
  Src, Dst: String;
begin
  if CurStep = ssPostInstall then
  begin
    Dst := ExpandConstant('{app}\licencia.lic');
    if not FileExists(Dst) then
    begin
      SetArrayLength(OldDirs, 4);
      OldDirs[0] := ExpandConstant('{autopf}\Toto Tools\Servidor de Pedidos\licencia.lic');
      OldDirs[1] := ExpandConstant('{autopf}\ServidorPedidos\licencia.lic');
      OldDirs[2] := ExpandConstant('{localappdata}\Toto Tools\Servidor de Pedidos\licencia.lic');
      OldDirs[3] := ExpandConstant('{commonappdata}\TotoTools\licencia.lic');
      for I := 0 to GetArrayLength(OldDirs) - 1 do
      begin
        Src := OldDirs[I];
        if (Src <> Dst) and FileExists(Src) then
        begin
          try
            FileCopy(Src, Dst, False);
          except
          end;
          if FileExists(Dst) then Break;
        end;
      end;
    end;
  end;
end;
