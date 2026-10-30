; Instalador - Servidor de Pedidos (Inno Setup 6)
; Workflow: .github/workflows/instalador.yml
; Muestra Términos + Política de privacidad; el usuario DEBE aceptar para continuar.

#define MyAppName "Servidor de Pedidos"
#define MyAppVersion "1.0"
#define MyAppPublisher "José Carlos Fernández"

[Setup]
AppId={{A8F3C2D1-9E4B-4F6A-8C1D-7B2E5A9F0D33}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\ServidorPedidos
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Texto legal obligatorio (Términos + Privacidad)
LicenseFile=LICENCIA_INSTALACION.txt
OutputDir=instalador
OutputBaseFilename=Instalar-ServidorPedidos
SetupIconFile=icono.ico
UninstallDisplayIcon={app}\ServidorPedidos.exe
UninstallDisplayName={#MyAppName}
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
VersionInfoCompany={#MyAppPublisher}
VersionInfoCopyright=Desarrollado por José Carlos Fernández
VersionInfoDescription={#MyAppName}
VersionInfoProductName={#MyAppName}
VersionInfoVersion={#MyAppVersion}.0.0

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[CustomMessages]
spanish.WelcomeLabel2=Este asistente instalará [name/ver] en su equipo.%n%nDesarrollado por José Carlos Fernández.%n%nA continuación deberá aceptar los Términos y condiciones y la Política de privacidad para continuar.%n%nSe recomienda cerrar otras aplicaciones antes de seguir.

[Messages]
spanish.LicenseLabel=Lea los Términos y condiciones y la Política de privacidad. Debe aceptarlos para continuar la instalación.
spanish.LicenseAccepted=Acepto los Términos y la Política de privacidad
spanish.LicenseNotAccepted=No acepto los Términos ni la Política de privacidad

[Tasks]
Name: "desktopicon"; Description: "Crear un icono en el escritorio"; GroupDescription: "Accesos directos:"

[Dirs]
; El servidor guarda pedidos, catálogo y fotos junto al .exe: la carpeta debe poder escribirse.
Name: "{app}"; Permissions: users-modify

[Files]
Source: "dist\ServidorPedidos.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Servidor de Pedidos"; Filename: "{app}\ServidorPedidos.exe"; WorkingDir: "{app}"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Servidor de Pedidos"; Filename: "{app}\ServidorPedidos.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\ServidorPedidos.exe"; Description: "Abrir el Servidor de Pedidos ahora"; Flags: nowait postinstall skipifsilent
