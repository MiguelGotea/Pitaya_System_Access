<#
.SYNOPSIS
    Importa cambios de archivos .bas / .cls de vuelta al archivo Pitaya_System.accdb.
    Operacion INVERSA a Export-AccessVBA.ps1.

.DESCRIPCION
    Lee los archivos VBA modificados en:
      vba\formularios\   Form_*.bas    -> actualiza solo el codigo del formulario
      vba\informes\      Report_*.bas  -> actualiza solo el codigo del informe
      vba\modulos\       *.bas         -> reemplaza el modulo completo
      vba\clases\        *.cls         -> reemplaza la clase completa

    NO modifica los controles, propiedades ni diseño de formularios/informes.
    Solo el codigo VBA cambia.

.PARAMETROS
    -AccdbTarget   Ruta del .accdb destino. Por defecto: el original (no la copia).
    -OnlyChanged   Si se especifica, solo importa archivos modificados segun git diff.
    -DryRun        Muestra que se importaria sin hacer cambios reales.

.USO
    # Importar todos los archivos VBA al .accdb original
    .\scripts\Import-AccessVBA.ps1

    # Solo los archivos que cambiaron desde el ultimo commit de git
    .\scripts\Import-AccessVBA.ps1 -OnlyChanged

    # Ver que se importaria sin hacer nada
    .\scripts\Import-AccessVBA.ps1 -DryRun

    # Importar a un .accdb especifico
    .\scripts\Import-AccessVBA.ps1 -AccdbTarget "C:\ruta\otro.accdb"
#>

param(
    # Por defecto trabaja sobre la COPIA en db\ — nunca toca el original automaticamente
    # Cuando estes listo, copia db\Pitaya_System.accdb manualmente a tu carpeta Sistema\
    [string]$AccdbTarget = (Join-Path (Split-Path $PSScriptRoot -Parent) "db\Pitaya_System.accdb"),
    [switch]$OnlyChanged,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── Rutas ───────────────────────────────────────────────────────────────────
$Root      = Split-Path $PSScriptRoot -Parent
$VbaDir    = Join-Path $Root "vba"
$FolmsDir  = Join-Path $VbaDir "formularios"
$InformDir = Join-Path $VbaDir "informes"
$ModDir    = Join-Path $VbaDir "modulos"
$ClsDir    = Join-Path $VbaDir "clases"
$LogFile   = Join-Path $Root "import_log.txt"
$TempDir   = Join-Path $env:TEMP "PitayaVBAImport"

$AccessExe = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"
if (-not (Test-Path $AccessExe)) {
    $AccessExe = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
}

# ─── Helpers ─────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Color = "Cyan")
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line -ForegroundColor $Color
    try {
        $fs = [System.IO.File]::Open($LogFile, [System.IO.FileMode]::Append,
              [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($line); $sw.Close(); $fs.Close()
    } catch { }
}

# Quita el bloque de cabecera que agrega Export-AccessVBA.ps1
# La cabecera va desde ' ===... hasta la segunda linea ' ===... + linea en blanco
function Strip-Header ([string]$Content) {
    $lines    = $Content -split "`r`n|`n"
    $sepCount = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^' ={10,}") {
            $sepCount++
            if ($sepCount -eq 2) {
                # Saltar la linea separadora y la linea en blanco que le sigue
                $startIdx = $i + 1
                if ($startIdx -lt $lines.Count -and $lines[$startIdx].Trim() -eq "") {
                    $startIdx++
                }
                return ($lines[$startIdx..($lines.Count - 1)] -join "`r`n")
            }
        }
    }
    return $Content  # si no hay cabecera, devolver tal cual
}

# Obtiene el nombre del componente desde el nombre del archivo (.bas/.cls)
function Get-CompName ([System.IO.FileInfo]$File) {
    return [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
}

# ─── Inicio ──────────────────────────────────────────────────────────────────
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

Write-Log "==========================================================" "Yellow"
Write-Log "  PITAYA SYSTEM ACCESS -- Importador VBA (.bas -> .accdb)" "Yellow"
if ($DryRun)     { Write-Log "  MODO: DRY RUN (sin cambios reales)" "Yellow" }
if ($OnlyChanged){ Write-Log "  MODO: Solo archivos modificados (git diff)" "Yellow" }
Write-Log "==========================================================" "Yellow"

# Resolver ruta del .accdb destino
$AccdbResolved = $null
try { $AccdbResolved = (Resolve-Path $AccdbTarget -ErrorAction Stop).Path } catch {}
if (-not $AccdbResolved) {
    Write-Log "ERROR: No se encontro el .accdb destino: $AccdbTarget" "Red"
    Write-Log "  Usa -AccdbTarget para especificar la ruta." "Red"
    exit 1
}
Write-Log "Destino : $AccdbResolved" "Green"
Write-Log "Fuente  : $Root\vba\" "Green"

# ─── Obtener lista de archivos a importar ─────────────────────────────────────
Write-Log "" "White"
Write-Log "-- Recopilando archivos VBA a importar --" "Magenta"

# Mapa: ruta_archivo -> tipo (Form/Report/Module/Class)
$allFiles = [System.Collections.Generic.List[hashtable]]::new()

Get-ChildItem $FolmsDir -Filter "*.bas" -ErrorAction SilentlyContinue |
    ForEach-Object { $allFiles.Add(@{ File=$_; Kind="Form" }) }
Get-ChildItem $InformDir -Filter "*.bas" -ErrorAction SilentlyContinue |
    ForEach-Object { $allFiles.Add(@{ File=$_; Kind="Report" }) }
Get-ChildItem $ModDir -Filter "*.bas" -ErrorAction SilentlyContinue |
    ForEach-Object { $allFiles.Add(@{ File=$_; Kind="Module" }) }
Get-ChildItem $ClsDir -Filter "*.cls" -ErrorAction SilentlyContinue |
    ForEach-Object { $allFiles.Add(@{ File=$_; Kind="Class" }) }

Write-Log "  Total archivos VBA encontrados: $($allFiles.Count)" "Cyan"

# Filtrar solo los cambiados si se pidio -OnlyChanged
if ($OnlyChanged) {
    Write-Log "  Detectando cambios con git diff..." "DarkCyan"
    $gitChanged = git -C $Root diff --name-only HEAD 2>$null
    $gitChangedSet = @{}
    $gitChanged | ForEach-Object { $gitChangedSet[$_.Replace("/","\")] = $true }

    # Tambien incluir staged y untracked
    $gitStaged   = git -C $Root diff --name-only --cached HEAD 2>$null
    $gitStaged   | ForEach-Object { $gitChangedSet[$_.Replace("/","\")] = $true }

    $allFiles = $allFiles | Where-Object {
        $rel = $_.File.FullName.Substring($Root.Length + 1)
        $gitChangedSet.ContainsKey($rel)
    } | ForEach-Object { $_ }   # reconvertir a lista

    Write-Log "  Archivos modificados a importar: $($allFiles.Count)" "Cyan"
    if ($allFiles.Count -eq 0) {
        Write-Log "  No hay cambios pendientes. Nada que importar." "Green"
        exit 0
    }
}

if ($DryRun) {
    Write-Log "" "White"
    Write-Log "-- DRY RUN: Archivos que se importarian --" "Magenta"
    foreach ($entry in $allFiles) {
        Write-Log "  [$($entry.Kind)]  $($entry.File.Name)" "DarkGray"
    }
    Write-Log "" "White"
    Write-Log "Ejecuta sin -DryRun para aplicar los cambios." "Yellow"
    exit 0
}

# ─── Lanzar Access y conectar ─────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "-- Lanzando Microsoft Access --" "Magenta"

# Limpiar instancias anteriores
Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$accessProc = $null
$access     = $null
$counters   = @{ Updated=0; Skipped=0; Errors=0 }

try {
    if (-not (Test-Path $AccessExe)) {
        Write-Log "ERROR: No se encontro MSACCESS.EXE" "Red"; exit 1
    }
    $accessProc = Start-Process -FilePath $AccessExe -PassThru

    $maxWait = 25; $waited = 0
    while ($waited -lt $maxWait) {
        Start-Sleep -Seconds 2; $waited += 2
        try {
            $access = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")
            Write-Log "  Access listo en ${waited}s (v$($access.Version))" "Green"
            break
        } catch { }
    }
    if ($null -eq $access) { Write-Log "ERROR: No pudo conectar Access." "Red"; exit 1 }

    $access.AutomationSecurity = 3
    Start-Sleep -Seconds 2

    Write-Log "  Abriendo: $AccdbResolved" "DarkCyan"
    $access.OpenCurrentDatabase($AccdbResolved, $false)
    Start-Sleep -Seconds 3

    $vbProject = $access.VBE.ActiveVBProject
    if ($null -eq $vbProject) {
        Write-Log "ERROR: No se encontro proyecto VBA en la BD." "Red"; exit 1
    }
    Write-Log "  Proyecto: '$($vbProject.Name)'" "DarkCyan"

    # Preparar directorio temporal para archivos sin cabecera
    if (-not (Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }

    # ─── Importar cada archivo ────────────────────────────────────────────────
    Write-Log "" "White"
    Write-Log "-- Importando modulos VBA --" "Magenta"

    foreach ($entry in $allFiles) {
        $file     = $entry.File
        $kind     = $entry.Kind
        $compName = Get-CompName $file

        try {
            # Leer contenido y quitar cabecera de exportacion
            $rawContent  = Get-Content $file.FullName -Raw -Encoding UTF8
            $cleanCode   = Strip-Header $rawContent

            # Buscar el componente en el proyecto VBA
            $comp = $null
            try { $comp = $vbProject.VBComponents.Item($compName) } catch { }

            if ($null -eq $comp) {
                Write-Log "  WARN '$compName' no existe en el proyecto. Saltando." "Yellow"
                $counters.Skipped++
                continue
            }

            $compType = $comp.Type   # 1=StdModule, 2=Class, 100=Document(Form/Report)

            if ($kind -eq "Form" -or $kind -eq "Report") {
                # FORMULARIOS / INFORMES: solo reemplazar el codigo (no el diseño)
                $mod = $comp.CodeModule
                if ($mod.CountOfLines -gt 0) {
                    $mod.DeleteLines(1, $mod.CountOfLines)
                }
                if ($cleanCode.Trim().Length -gt 0) {
                    $mod.InsertLines(1, $cleanCode.TrimEnd())
                }
                Write-Log "  UPD [$kind]  $($file.Name)" "Green"
                $counters.Updated++

            } else {
                # MODULOS Y CLASES: eliminar el componente y reimportar desde archivo limpio
                # Crear archivo temporal sin cabecera
                $tempFile = Join-Path $TempDir $file.Name
                Set-Content -Path $tempFile -Value $cleanCode -Encoding UTF8 -NoNewline

                # Para modulos estandar el archivo .bas debe tener la linea
                # "Attribute VB_Name = ..." al inicio para que Access lo reconozca
                # Si no la tiene, la agregamos
                if ($cleanCode -notmatch "^Attribute VB_Name") {
                    $attr = "Attribute VB_Name = `"$compName`"`r`n"
                    Set-Content -Path $tempFile -Value ($attr + $cleanCode) -Encoding UTF8 -NoNewline
                }

                $vbProject.VBComponents.Remove($comp)
                Start-Sleep -Milliseconds 200
                $vbProject.VBComponents.Import($tempFile) | Out-Null
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                Write-Log "  IMP [$kind]  $($file.Name)" "Green"
                $counters.Updated++
            }
        } catch {
            Write-Log "  ERROR '$compName': $_" "Red"
            $counters.Errors++
        }
    }

    # Guardar cambios
    Write-Log "" "White"
    Write-Log "  Guardando cambios en el .accdb..." "DarkCyan"
    $access.CurrentDb().Close()
    Start-Sleep -Seconds 1

} catch {
    Write-Log "ERROR CRITICO: $_" "Red"
    Write-Log $_.ScriptStackTrace "Red"
} finally {
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch { }
        if ($null -ne $accessProc) { try { $access.Quit() } catch { } }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null } catch { }
        $access = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Write-Log "  COM liberado." "DarkGray"
    }
    if ($null -ne $accessProc) {
        Start-Sleep -Seconds 1
        Stop-Process -Id $accessProc.Id -Force -ErrorAction SilentlyContinue
        Write-Log "  Access cerrado." "DarkGray"
    }
    if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

# ─── Resumen ──────────────────────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "==========================================================" "Yellow"
Write-Log "  IMPORTACION COMPLETADA" "Yellow"
Write-Log "  Actualizados : $($counters.Updated)" "Green"
Write-Log "  Saltados     : $($counters.Skipped)" "DarkGray"
if ($counters.Errors -gt 0) {
    Write-Log "  Errores      : $($counters.Errors)  (ver import_log.txt)" "Red"
}
Write-Log "  Log: import_log.txt" "Cyan"
Write-Log "==========================================================" "Yellow"
Write-Log "" "White"
Write-Log "  SIGUIENTE PASO RECOMENDADO:" "Yellow"
Write-Log "  1. Abre y prueba:  db\Pitaya_System.accdb" "White"
Write-Log "  2. Si todo OK, copia manualmente:" "White"
Write-Log "     db\Pitaya_System.accdb  -->  C:\...\Sistema\Pitaya_System.accdb" "White"
Write-Log "  3. Ejecuta Export-AccessVBA.ps1 para sincronizar vba\ y sql\" "White"
Write-Log "  4. Ejecuta gitpush.ps1 para documentar en GitHub" "White"
Write-Log "==========================================================" "Yellow"
