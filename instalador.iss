; Instalador de Toto Tools - Servidor de Pedidos (Inno Setup 6)
; Se construye con el workflow "Construir instalador PC" (instalador.yml).

[Setup]
AppName=Toto Tools - Servidor de Pedidos
AppVersion=1.0
AppPublisher=Toto Tools
DefaultDirName={autopf}\Toto Tools\Servidor de Pedidos
DefaultGroupName=Toto Tools
DisableProgramGroupPage=yes
OutputDir=instalador
OutputBaseFilename=Instalar-ServidorPedidos
SetupIconFile=icono.ico
UninstallDisplayIcon={app}\ServidorPedidos.exe
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un icono en el escritorio"; GroupDescription: "Accesos directos:"

[Dirs]
; El servidor guarda pedidos, catalogo y fotos junto al .exe: la carpeta debe poder escribirse.
Name: "{app}"; Permissions: users-modify

[Files]
Source: "dist\ServidorPedidos.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Servidor de Pedidos"; Filename: "{app}\ServidorPedidos.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\Servidor de Pedidos"; Filename: "{app}\ServidorPedidos.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\ServidorPedidos.exe"; Description: "Abrir el Servidor de Pedidos ahora"; Flags: nowait postinstall skipifsilent
