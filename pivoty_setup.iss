; Script de instalación para Pivoty
; Desarrollado para automatización profesional de Excel

[Setup]
AppId={{shohan-pivoty-bot-2026}}
AppName=Pivoty
AppVersion=1.0.0
AppPublisher=Shohan
DefaultDirName={autopf}\Pivoty
DefaultGroupName=Pivoty
AllowNoIcons=yes
; El icono del instalador
SetupIconFile=assets\LogoIconoDino.ico
; El ejecutable final generado por PyInstaller
OutputBaseFilename=Pivoty_Installer_v1.0.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Imágenes personalizadas del Dinosaurio (Ahora en PNG con transparencia)
WizardImageFile=assets\WelcomeDinoSinFondo.png
WizardSmallImageFile=assets\ThanksDinoSinFondo.png

; --- SEGURIDAD DEL INSTALADOR ---
; 1. Contraseña Global para abrir el instalador
Password=pivoprotect2026
; 2. Activa la página de Información de Usuario (Nombre y Licencia)
UserInfoPage=yes


[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Dirs]
Name: "{app}\logs"; Permissions: users-full
Name: "{app}\config"; Permissions: users-full

[Files]
; Archivo principal
Source: "dist\Pivoty.exe"; DestDir: "{app}"; Flags: ignoreversion
; Carpetas de configuración esenciales (vacías o por defecto)
Source: "config\*"; DestDir: "{app}\config"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "assets\LogoIconoDino.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: ".env"; DestDir: "{app}"; Flags: ignoreversion hidden; Permissions: users-full
; Crear carpeta de logs vacía
Source: "logs\*"; DestDir: "{app}\logs"; Flags: ignoreversion recursesubdirs createallsubdirs



[Icons]
Name: "{group}\Pivoty"; Filename: "{app}\Pivoty.exe"; IconFilename: "{app}\LogoIconoDino.ico"
Name: "{autodesktop}\Pivoty"; Filename: "{app}\Pivoty.exe"; Tasks: desktopicon; IconFilename: "{app}\LogoIconoDino.ico"

[Run]
Filename: "{app}\Pivoty.exe"; Description: "{cm:LaunchProgram,Pivoty}"; Flags: nowait postinstall skipifsilent

[Code]
// Función para validar la Licencia (Serial) durante la instalación
function CheckSerial(Serial: String): Boolean;
begin
  // --- LÓGICA DE MULTI-LICENCIA ---
  
  // 1. Clave Maestra (Para Shohan - El Creador)
  if Serial = 'SHOHAN-MASTER-2026' then
  begin
    Result := True;
    Exit;
  end;

  // 2. Claves para Clientes/Usuarios (Puedes añadir más aquí)
  if (Serial = 'PIVOTY-USER-001-X7') or 
     (Serial = 'PIVOTY-USER-002-Y9') or
     (Serial = 'PIVOTY-V1-DINO-2026') then
  begin
    Result := True;
  end
  else
  begin
    MsgBox('Licencia Inválida. Por favor, contacta a Shohan para obtener una clave válida.', mbError, MB_OK);
    Result := False;
  end;
end;
