<#
.SYNOPSIS
    Exporta todos los modulos VBA y SQL de Pitaya_System.accdb hacia
    carpetas organizadas dentro de Pitaya_System_Access.

.DESCRIPCION
    Estructura generada:
      db\                   copia de Pitaya_System.accdb
      vba\formularios\      Form_*.bas
      vba\informes\         Report_*.bas
      vba\modulos\          modulos estandar .bas
      vba\clases\           clases .cls
      sql\queries\          SQL de cada consulta .sql
      export_log.txt        log del proceso

.USO
    .\Export-AccessVBA.ps1
    .\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\otro.accdb"

.NOTAS
    Compatible con Office 365/2024 Click-to-Run.
    Lanza Access como proceso, luego se conecta via GetActiveObject.
    Requiere que no haya otra BD de Access abierta al ejecutarlo.
#>

param(
    [string]$AccdbPath = (Join-Path $PSScriptRoot "..\..\..\..\Pitaya_System.accdb")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── Rutas ───────────────────────────────────────────────────────────────────
# El script vive en scripts\ pero la salida va al directorio padre del proyecto
$Root      = Split-Path $PSScriptRoot -Parent
$DbDir     = Join-Path $Root "db"
$VbaDir    = Join-Path $Root "vba"
$FolmsDir  = Join-Path $VbaDir "formularios"
$InformDir = Join-Path $VbaDir "informes"
$ModDir    = Join-Path $VbaDir "modulos"
$ClsDir    = Join-Path $VbaDir "clases"
$SqlDir    = Join-Path $Root "sql"
$QryDir    = Join-Path $SqlDir "queries"
$LogFile   = Join-Path $Root "export_log.txt"

$AccessExe = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"
if (-not (Test-Path $AccessExe)) {
    # Fallback a instalacion MSI 32-bit
    $AccessExe = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
}

# ─── Helpers ─────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Color = "Cyan")
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line -ForegroundColor $Color
    # Usar StreamWriter con FileShare.ReadWrite para evitar conflictos con VS Code u otros editores
    try {
        $fs = [System.IO.File]::Open($LogFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($line)
        $sw.Close()
        $fs.Close()
    } catch { <# ignorar errores de log para no romper el flujo principal #> }
}

function Ensure-Dir ([string]$Path) {
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Log "  Carpeta creada: $(Split-Path $Path -Leaf)" "DarkGray"
    }
}
function Safe-FileName ([string]$Name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $Name = $Name.Replace([string]$c, '_') }
    return $Name
}

# ─── Inicio ──────────────────────────────────────────────────────────────────
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

Write-Log "==========================================================" "Yellow"
Write-Log "  PITAYA SYSTEM ACCESS -- Exportador VBA + SQL" "Yellow"
Write-Log "==========================================================" "Yellow"

# Resolver ruta del .accdb fuente
$AccdbResolved = $null
try { $AccdbResolved = (Resolve-Path $AccdbPath -ErrorAction Stop).Path } catch {}
if (-not $AccdbResolved) {
    Write-Log "ERROR: No se encontro el archivo: $AccdbPath" "Red"
    exit 1
}
Write-Log "Origen  : $AccdbResolved" "Green"
Write-Log "Destino : $Root" "Green"

# ── 1. Crear carpetas ────────────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "-- [1/4] Creando estructura de carpetas --" "Magenta"
@($DbDir, $FolmsDir, $InformDir, $ModDir, $ClsDir, $QryDir) | ForEach-Object { Ensure-Dir $_ }

# ── 2. Copiar .accdb ─────────────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "-- [2/4] Copiando base de datos --" "Magenta"
$DbDest = Join-Path $DbDir "Pitaya_System.accdb"
Copy-Item -Path $AccdbResolved -Destination $DbDest -Force
$sizeMb = [math]::Round((Get-Item $DbDest).Length / 1MB, 1)
Write-Log "  Copiado: db\Pitaya_System.accdb  ($sizeMb MB)" "Green"

# ── 3. Extraer VBA via COM (Access lanzado como proceso) ─────────────────────
Write-Log "" "White"
Write-Log "-- [3/4] Extrayendo modulos VBA --" "Magenta"

$accessProc = $null
$access      = $null
$counters    = @{ Forms=0; Reports=0; Modules=0; Classes=0; Skipped=0; Errors=0 }

try {
    # Verificar si Access ya esta corriendo (puede fallar si no esta abierto)
    $existingAccess = $null
    try {
        $existingAccess = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")
        Write-Log "  Usando instancia de Access ya abierta (version $($existingAccess.Version))" "DarkCyan"
    } catch { }

    if ($null -ne $existingAccess) {
        $access = $existingAccess
    } else {
        # Lanzar Access en segundo plano
        if (-not (Test-Path $AccessExe)) {
            Write-Log "  ERROR: No se encontro MSACCESS.EXE en: $AccessExe" "Red"
            exit 1
        }
        Write-Log "  Lanzando Access como proceso..." "DarkCyan"
        $accessProc = Start-Process -FilePath $AccessExe -PassThru

        # Esperar a que Access este listo para responder al COM
        $maxWait = 20
        $waited  = 0
        $access  = $null
        while ($waited -lt $maxWait) {
            Start-Sleep -Seconds 2
            $waited += 2
            try {
                $access = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")
                Write-Log "  Access listo en ${waited}s (version $($access.Version))" "Green"
                break
            } catch { }
        }
        if ($null -eq $access) {
            Write-Log "  ERROR: No se pudo conectar a Access tras ${maxWait}s" "Red"
            exit 1
        }
    }

    # Configurar Access
    $access.AutomationSecurity = 3

    # Dar tiempo a que Access inicialice completamente
    Start-Sleep -Seconds 3

    # Abrir la BD (desde la copia en db\)
    Write-Log "  Abriendo: $DbDest" "DarkCyan"
    $access.OpenCurrentDatabase($DbDest, $false)
    Start-Sleep -Seconds 3

    $vbProject = $access.VBE.ActiveVBProject
    if ($null -eq $vbProject) {
        Write-Log "  ADVERTENCIA: No se encontro proyecto VBA en esta BD." "Yellow"
    } else {
        $totalComps = $vbProject.VBComponents.Count
        Write-Log "  Proyecto: '$($vbProject.Name)'  |  Componentes: $totalComps" "DarkCyan"

        foreach ($comp in $vbProject.VBComponents) {
            $compName = $comp.Name
            $compType = $comp.Type   # 1=StdModule 2=ClassModule 100=Document(Form/Report)

            try {
                $lines = $comp.CodeModule.CountOfLines
                if ($lines -eq 0) {
                    Write-Log "  SKIP (vacio): $compName" "DarkGray"
                    $counters.Skipped++
                    continue
                }
                $code = $comp.CodeModule.Lines(1, $lines)

                # Clasificar por nombre y tipo
                if ($compName -like "Form_*") {
                    $destDir = $FolmsDir; $ext = ".bas"; $counters.Forms++
                } elseif ($compName -like "Report_*") {
                    $destDir = $InformDir; $ext = ".bas"; $counters.Reports++
                } elseif ($compType -eq 2) {
                    $destDir = $ClsDir; $ext = ".cls"; $counters.Classes++
                } else {
                    $destDir = $ModDir; $ext = ".bas"; $counters.Modules++
                }

                $safeFile = (Safe-FileName $compName) + $ext
                $outPath  = Join-Path $destDir $safeFile

                $header = @"
' ==========================================================
' Modulo  : $compName
' Tipo    : $compType
' Lineas  : $lines
' Proyecto: $($vbProject.Name)
' Exportado: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
' ==========================================================
"@
                Set-Content -Path $outPath -Value ($header + "`r`n" + $code) -Encoding UTF8
                Write-Log "  OK  $safeFile  ($lines lineas)" "DarkGray"
            } catch {
                Write-Log "  WARN no exportado '$compName': $_" "Yellow"
                $counters.Errors++
            }
        }
    }

    # ── 4. Extraer SQL de consultas ──────────────────────────────────────────
    Write-Log "" "White"
    Write-Log "-- [4/4] Extrayendo SQL de consultas --" "Magenta"

    $db = $access.CurrentDb()
    $qryCount = 0

    for ($i = 0; $i -lt $db.QueryDefs.Count; $i++) {
        $qdef  = $db.QueryDefs($i)
        $qName = $qdef.Name
        if ($qName.StartsWith("~")) { continue }   # queries internas de Access

        try {
            $sql     = $qdef.SQL
            $safeQry = (Safe-FileName $qName) + ".sql"
            $outQry  = Join-Path $QryDir $safeQry
            $sqlHeader = @"
-- ==========================================================
-- Consulta : $qName
-- Exportado: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
-- ==========================================================

"@
            Set-Content -Path $outQry -Value ($sqlHeader + $sql) -Encoding UTF8
            $qryCount++
            Write-Log "  SQL  $safeQry" "DarkGray"
        } catch {
            Write-Log "  WARN query '$qName': $_" "Yellow"
        }
    }
    Write-Log "  Total consultas exportadas: $qryCount" "Cyan"

} catch {
    Write-Log "ERROR CRITICO: $_" "Red"
    Write-Log $_.ScriptStackTrace "Red"
} finally {
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch { }
        # Solo cerrar Access si nosotros lo abrimos
        if ($null -ne $accessProc) {
            try { $access.Quit() } catch { }
        }
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        } catch { }
        $access = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-Log "  COM liberado." "DarkGray"
    }
    if ($null -ne $accessProc) {
        Start-Sleep -Seconds 1
        Stop-Process -Id $accessProc.Id -Force -ErrorAction SilentlyContinue
        Write-Log "  Proceso Access cerrado (PID $($accessProc.Id))." "DarkGray"
    }
}

# ── Resumen ──────────────────────────────────────────────────────────────────
$qryFiles = Get-ChildItem $QryDir -Filter "*.sql" -ErrorAction SilentlyContinue
$qryFinal = if ($qryFiles) { $qryFiles.Count } else { 0 }
Write-Log "" "White"
Write-Log "==========================================================" "Yellow"
Write-Log "  EXPORTACION COMPLETADA" "Yellow"
Write-Log "  Formularios : $($counters.Forms)    vba\formularios\" "Green"
Write-Log "  Informes    : $($counters.Reports)   vba\informes\" "Green"
Write-Log "  Modulos     : $($counters.Modules)   vba\modulos\" "Green"
Write-Log "  Clases      : $($counters.Classes)   vba\clases\" "Green"
Write-Log "  Vacios omit.: $($counters.Skipped)" "DarkGray"
Write-Log "  SQL queries : $qryFinal   sql\queries\" "Green"
if ($counters.Errors -gt 0) {
    Write-Log "  Errores     : $($counters.Errors)   (ver export_log.txt)" "Yellow"
}
Write-Log "  Log: export_log.txt" "Cyan"
Write-Log "==========================================================" "Yellow"
