; Inno Setup 6/7 - genera dist\PortafolioQA-Setup.exe
; Se compila solo desde Construir-Instalador.ps1 (o: ISCC.exe Instalador\PortafolioQA.iss tras construir el paquete).
; Este archivo debe guardarse como UTF-8 con BOM (tiene tildes).
; El instalador NO toca SQL Server: la conexión a la base de datos se configura después, desde la aplicación (/Setup).
#define Version "2.0.0"

[Setup]
AppName=Portafolio QA
AppVersion={#Version}
CreateAppDir=no
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=admin
; Modo 64 bits: {sys} apunta al System32 real (PowerShell/appcmd/IIS de 64 bits, sin redirección a SysWOW64).
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=PortafolioQA-Setup
Compression=lzma2
SolidCompression=yes
Uninstallable=no
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Files]
Source: "..\dist\PortafolioQA-Instalador\app\*"; DestDir: "{tmp}\pqa\app"; Flags: recursesubdirs createallsubdirs ignoreversion deleteafterinstall
Source: "..\dist\PortafolioQA-Instalador\Instalar.ps1"; DestDir: "{tmp}\pqa"; Flags: ignoreversion deleteafterinstall
; Se extrae antes de mostrar el asistente (lista los certificados del equipo).
Source: "..\dist\PortafolioQA-Instalador\Utilidades.ps1"; DestDir: "{tmp}"; Flags: dontcopy

[Code]
var
  pgModo: TInputOptionWizardPage;
  pgSitio: TInputQueryWizardPage;
  pgCert: TInputOptionWizardPage;
  pgAdmin: TInputQueryWizardPage;
  CertThumbs: TArrayOfString;
  UtilPath: String;
  Fallo: Boolean;
  Advertencia: Boolean;
  LogPath: String;

function PsExe: String;
begin
  Result := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
end;

function EsActualizacion: Boolean;
begin
  Result := pgModo.SelectedValueIndex = 1;
end;

procedure Agregar(var L: TArrayOfString; Clave, Valor: String);
var N: Integer;
begin
  N := GetArrayLength(L);
  SetArrayLength(L, N + 1);
  L[N] := Clave + '=' + Valor;
end;

function ListarCertificados(Salida: String): Boolean;
var
  Rc: Integer;
  Args: String;
begin
  Args := '-NoProfile -ExecutionPolicy Bypass -File "' + UtilPath + '" -Accion Certificados -Salida "' + Salida + '"';
  Result := Exec(PsExe, Args, '', SW_HIDE, ewWaitUntilTerminated, Rc) and (Rc = 0);
end;

procedure CargarCertificados;
var
  Lineas: TArrayOfString;
  Salida: String;
  I, P: Integer;
begin
  pgCert.Add('Crear un certificado autofirmado (solo PRUEBAS: los navegadores mostrarán una advertencia)');
  SetArrayLength(CertThumbs, 0);
  Salida := ExpandConstant('{tmp}\certs.txt');
  DeleteFile(Salida);
  if ListarCertificados(Salida) and LoadStringsFromFile(Salida, Lineas) then
  begin
    for I := 0 to GetArrayLength(Lineas) - 1 do
    begin
      P := Pos('|', Lineas[I]);
      if P > 1 then
      begin
        SetArrayLength(CertThumbs, GetArrayLength(CertThumbs) + 1);
        CertThumbs[GetArrayLength(CertThumbs) - 1] := Copy(Lineas[I], 1, P - 1);
        pgCert.Add(Copy(Lineas[I], P + 1, Length(Lineas[I])));
      end;
    end;
  end;
  if GetArrayLength(CertThumbs) > 0 then pgCert.SelectedValueIndex := 1
  else pgCert.SelectedValueIndex := 0;
end;

procedure InitializeWizard;
begin
  ExtractTemporaryFile('Utilidades.ps1');
  UtilPath := ExpandConstant('{tmp}\Utilidades.ps1');

  pgModo := CreateInputOptionPage(wpWelcome, 'Tipo de instalación',
    '¿Qué desea hacer?', 'Elija una opción y presione Siguiente. La base de datos NO se configura aquí: se hace después, desde la propia aplicación.', True, False);
  pgModo.Add('Instalación nueva (crea el sitio IIS y el application pool; la aplicación queda pendiente de conectarse a la base de datos)');
  pgModo.Add('Actualizar una instalación existente (reemplaza archivos; no toca la configuración ni la base de datos)');
  pgModo.SelectedValueIndex := 0;

  pgSitio := CreateInputQueryPage(pgModo.ID, 'Sitio web (IIS)',
    'Datos del sitio', 'Indique cómo se publicará la aplicación.');
  pgSitio.Add('Nombre del sitio IIS:', False);
  pgSitio.Add('Carpeta del sitio:', False);
  pgSitio.Add('Host del sitio (vacío = cualquier host):', False);
  pgSitio.Add('Puerto HTTPS:', False);
  pgSitio.Add('Carpeta persistente de claves de protección de datos (se incluye en el respaldo):', False);
  pgSitio.Values[0] := 'PortafolioQA';
  pgSitio.Values[1] := 'C:\inetpub\portafolio-qa';
  pgSitio.Values[3] := '443';
  pgSitio.Values[4] := 'D:\portafolio-qa-keys';

  pgCert := CreateInputOptionPage(pgSitio.ID, 'Certificado HTTPS',
    'Certificado del sitio', 'Elija el certificado con el que se publicará el sitio.', True, False);
  CargarCertificados;

  pgAdmin := CreateInputQueryPage(pgCert.ID, 'Administrador inicial y Squash TM',
    'Primer administrador de la aplicación',
    'Se crea al activar la aplicación, solo si la base de datos no tiene usuarios. Cambie la contraseña en el primer ingreso.');
  pgAdmin.Add('Correo del administrador inicial:', False);
  pgAdmin.Add('Contraseña inicial (mínimo 10 caracteres, con mayúscula, minúscula y número):', True);
  pgAdmin.Add('Repita la contraseña:', True);
  pgAdmin.Add('URL de Squash TM (vacío = sin integración con Squash):', False);
  pgAdmin.Add('Token de API de Squash TM (opcional; permiso de solo lectura):', True);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if EsActualizacion and ((PageID = pgCert.ID) or (PageID = pgAdmin.ID)) then
    Result := True;
end;

function UrlAsistente: String;
var Host, Puerto: String;
begin
  Host := Trim(pgSitio.Values[2]);
  if Host = '' then Host := 'localhost';
  Puerto := Trim(pgSitio.Values[3]);
  if (Puerto = '') or (Puerto = '443') then Result := 'https://' + Host + '/Setup'
  else Result := 'https://' + Host + ':' + Puerto + '/Setup';
end;

procedure CurPageChanged(CurPageID: Integer);
var I: Integer;
begin
  if CurPageID = pgSitio.ID then
  begin
    // En una actualización solo hace falta el nombre del sitio (la carpeta y el pool se leen de IIS).
    for I := 1 to 4 do
    begin
      pgSitio.PromptLabels[I].Visible := not EsActualizacion;
      pgSitio.Edits[I].Visible := not EsActualizacion;
    end;
    if EsActualizacion then
      pgSitio.Description := 'Indique el nombre del sitio IIS que desea actualizar.'
    else
      pgSitio.Description := 'Indique cómo se publicará la aplicación.';
  end;
  if CurPageID = wpFinished then
  begin
    if Fallo then
      WizardForm.FinishedLabel.Caption := 'La instalación NO se completó.' + #13#10#13#10 + 'Revise el detalle en:' + #13#10 + LogPath
    else if EsActualizacion then
      WizardForm.FinishedLabel.Caption := 'Actualización terminada. La conexión a la base de datos no se modificó.' + #13#10#13#10 +
        'Si la versión nueva trae cambios en la base de datos, un Administrador verá el aviso y debe aplicarlos en ' +
        'Administración > Conexión a la base de datos.' + #13#10#13#10 + 'Registro: ' + LogPath
    else
      WizardForm.FinishedLabel.Caption := 'Instalación terminada. La aplicación quedó SIN ACTIVAR: falta configurar la conexión a la base de datos.' + #13#10#13#10 +
        '1. Abra ' + UrlAsistente + #13#10 +
        '2. Escriba el código de configuración que está en:' + #13#10 + '   ' + ExpandConstant('{commonappdata}\PortafolioQA\codigo-configuracion.txt') + #13#10 +
        '3. Indique servidor, base y autenticación, pruebe la conexión y active la aplicación.' + #13#10 +
        '4. Ya activa, ingrese con el administrador inicial, cambie su contraseña y elimine las variables ' +
        'Seguridad__AdminInicial__* del application pool (IIS > Application Pools > Configuration Editor).' + #13#10#13#10 +
        'Registro: ' + LogPath;
    if Advertencia and (not Fallo) then
      WizardForm.FinishedLabel.Caption := WizardForm.FinishedLabel.Caption + #13#10#13#10 +
        'ATENCIÓN: el sitio no respondió en /healthz. Revise el registro y la carpeta logs del sitio.';
  end;
end;

function ClaveValida(const S: String): Boolean;
var I: Integer; C: Char; Dig, Mi, Ma: Boolean;
begin
  Dig := False; Mi := False; Ma := False;
  for I := 1 to Length(S) do
  begin
    C := S[I];
    if (C >= '0') and (C <= '9') then Dig := True;
    if (C >= 'a') and (C <= 'z') then Mi := True;
    if (C >= 'A') and (C <= 'Z') then Ma := True;
  end;
  Result := (Length(S) >= 10) and Dig and Mi and Ma;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var Puerto: Integer;
begin
  Result := True;

  if CurPageID = pgSitio.ID then
  begin
    if Trim(pgSitio.Values[0]) = '' then
    begin
      MsgBox('Escriba el nombre del sitio IIS.', mbError, MB_OK); Result := False; Exit;
    end;
    if not EsActualizacion then
    begin
      if Trim(pgSitio.Values[1]) = '' then
      begin
        MsgBox('Escriba la carpeta del sitio.', mbError, MB_OK); Result := False; Exit;
      end;
      Puerto := StrToIntDef(Trim(pgSitio.Values[3]), 0);
      if (Puerto < 1) or (Puerto > 65535) then
      begin
        MsgBox('El puerto HTTPS debe ser un número entre 1 y 65535.', mbError, MB_OK); Result := False; Exit;
      end;
      if Trim(pgSitio.Values[4]) = '' then
      begin
        MsgBox('Escriba la carpeta de claves de protección de datos.', mbError, MB_OK); Result := False; Exit;
      end;
    end;
  end;

  if CurPageID = pgAdmin.ID then
  begin
    if Pos('@', pgAdmin.Values[0]) < 2 then
    begin
      MsgBox('Escriba un correo válido para el administrador inicial.', mbError, MB_OK); Result := False; Exit;
    end;
    if not ClaveValida(pgAdmin.Values[1]) then
    begin
      MsgBox('La contraseña debe tener al menos 10 caracteres, con mayúscula, minúscula y número.', mbError, MB_OK); Result := False; Exit;
    end;
    if pgAdmin.Values[1] <> pgAdmin.Values[2] then
    begin
      MsgBox('Las contraseñas no coinciden.', mbError, MB_OK); Result := False; Exit;
    end;
    if (pgAdmin.Values[4] <> '') and (Trim(pgAdmin.Values[3]) = '') then
    begin
      MsgBox('Para usar un token de Squash TM indique también su URL.', mbError, MB_OK); Result := False; Exit;
    end;
    if (Trim(pgAdmin.Values[3]) <> '') and (Pos('http', LowerCase(Trim(pgAdmin.Values[3]))) <> 1) then
    begin
      MsgBox('La URL de Squash TM debe empezar con http:// o https://', mbError, MB_OK); Result := False; Exit;
    end;
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var S, Cert: String;
begin
  if EsActualizacion then
  begin
    S := 'Actualización del sitio IIS: ' + Trim(pgSitio.Values[0]) + NewLine + NewLine +
         'Se reemplazan los archivos de la aplicación. No se modifica la configuración ni la base de datos.';
  end else
  begin
    if pgCert.SelectedValueIndex = 0 then Cert := 'Autofirmado (PRUEBAS)'
    else Cert := pgCert.CheckListBox.Items[pgCert.SelectedValueIndex];
    S := 'Instalación nueva' + NewLine + NewLine +
         'Sitio IIS: ' + Trim(pgSitio.Values[0]) + NewLine +
         Space + 'Carpeta: ' + Trim(pgSitio.Values[1]) + NewLine +
         Space + 'Host: ' + Trim(pgSitio.Values[2]) + NewLine +
         Space + 'Puerto HTTPS: ' + Trim(pgSitio.Values[3]) + NewLine +
         Space + 'Certificado: ' + Cert + NewLine +
         Space + 'Claves de protección de datos: ' + Trim(pgSitio.Values[4]) + NewLine + NewLine +
         'Administrador inicial: ' + Trim(pgAdmin.Values[0]) + NewLine;
    if Trim(pgAdmin.Values[3]) <> '' then S := S + 'Squash TM: ' + Trim(pgAdmin.Values[3]) + NewLine;
    if pgAdmin.Values[4] <> '' then S := S + 'Token de Squash TM: (se guardará)' + NewLine;
    S := S + NewLine + 'Base de datos: no se toca. Al terminar, la aplicación queda sin activar y la conexión a SQL Server se configura desde ' + UrlAsistente + '.';
  end;
  Result := S;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  L, Lineas: TArrayOfString;
  Params, Args, Cola: String;
  Rc, I, Desde: Integer;
begin
  if CurStep <> ssPostInstall then Exit;

  SetArrayLength(L, 0);
  if EsActualizacion then Agregar(L, 'Modo', 'actualizar') else Agregar(L, 'Modo', 'nuevo');
  Agregar(L, 'Sitio', Trim(pgSitio.Values[0]));
  if not EsActualizacion then
  begin
    Agregar(L, 'Carpeta', Trim(pgSitio.Values[1]));
    Agregar(L, 'Host', Trim(pgSitio.Values[2]));
    Agregar(L, 'Puerto', Trim(pgSitio.Values[3]));
    Agregar(L, 'CarpetaClaves', Trim(pgSitio.Values[4]));
    if pgCert.SelectedValueIndex = 0 then Agregar(L, 'CertThumb', 'AUTOFIRMADO')
    else Agregar(L, 'CertThumb', CertThumbs[pgCert.SelectedValueIndex - 1]);
    Agregar(L, 'AdminCorreo', Trim(pgAdmin.Values[0]));
    Agregar(L, 'AdminClave', pgAdmin.Values[1]);
    Agregar(L, 'SquashUrl', Trim(pgAdmin.Values[3]));
    Agregar(L, 'SquashToken', pgAdmin.Values[4]);
  end;

  Params := ExpandConstant('{tmp}\pqa\parametros.ini');
  SaveStringsToUTF8File(Params, L, False);
  ForceDirectories(ExpandConstant('{commonappdata}\PortafolioQA'));
  LogPath := ExpandConstant('{commonappdata}\PortafolioQA\instalacion.log');

  WizardForm.StatusLabel.Caption := 'Instalando Portafolio QA (puede tardar unos minutos)...';
  WizardForm.FilenameLabel.Caption := 'Detalle en: ' + LogPath;
  Args := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{tmp}\pqa\Instalar.ps1') + '"' +
          ' -RutaApp "' + ExpandConstant('{tmp}\pqa\app') + '"' +
          ' -Parametros "' + Params + '" -Log "' + LogPath + '"';
  if not Exec(PsExe, Args, ExpandConstant('{tmp}\pqa'), SW_HIDE, ewWaitUntilTerminated, Rc) then Rc := -1;
  DeleteFile(Params);

  if Rc = 2 then Advertencia := True
  else if Rc <> 0 then
  begin
    Fallo := True;
    Cola := '';
    if LoadStringsFromFile(LogPath, Lineas) then
    begin
      Desde := GetArrayLength(Lineas) - 12;
      if Desde < 0 then Desde := 0;
      for I := Desde to GetArrayLength(Lineas) - 1 do Cola := Cola + Lineas[I] + #13#10;
    end;
    MsgBox('La instalación no se completó (código ' + IntToStr(Rc) + ').' + #13#10#13#10 + Cola + #13#10 + 'Registro completo: ' + LogPath, mbError, MB_OK);
  end;
end;

function InitializeSetup: Boolean;
var Faltan: String;
begin
  Faltan := '';
  if not RegKeyExists(HKLM, 'SYSTEM\CurrentControlSet\Services\W3SVC') then
    Faltan := Faltan + #13#10 + ' - IIS (rol Web Server) no está instalado.';
  if not FileExists(ExpandConstant('{commonpf64}\IIS\Asp.Net Core Module\V2\aspnetcorev2.dll')) then
    Faltan := Faltan + #13#10 + ' - .NET 8 Hosting Bundle (https://dotnet.microsoft.com/download/dotnet/8.0).';
  if Faltan <> '' then
  begin
    MsgBox('Faltan requisitos en este servidor:' + Faltan + #13#10#13#10 +
           'Instálelos (y ejecute "iisreset" tras el Hosting Bundle) y vuelva a ejecutar este instalador.', mbError, MB_OK);
    Result := False;
  end else
    Result := True;
end;
