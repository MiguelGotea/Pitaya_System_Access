<#
.SYNOPSIS
    Extrae TODO de Pitaya_System.accdb en una sola ejecucion:
      VBA (formularios, informes, modulos, clases)
      SQL de consultas guardadas
      DDL de estructura de tablas (CREATE TABLE)
      Datos de tablas (INSERT INTO) -- opcional, puede ser grande

.USO
    # Ejecucion completa (VBA + consultas + DDL de tablas)
    .\scripts\Export-AccessVBA.ps1

    # Incluir tambien datos de tablas (INSERT INTO)
    .\scripts\Export-AccessVBA.ps1 -ExportTableData

    # Con ruta personalizada al .accdb
    .\scripts\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\Pitaya_System.accdb"

.SALIDA
    db\                   copia de trabajo del .accdb (ignorada por git)
    vba\formularios\      Form_*.bas
    vba\informes\         Report_*.bas
    vba\modulos\          modulos estandar .bas
    vba\clases\           clases .cls
    sql\queries\          consultas guardadas .sql
    sql\tables\           CREATE TABLE por cada tabla .sql
    sql\data\             INSERT INTO por cada tabla .sql  (solo con -ExportTableData)
    export_log.txt        log detallado del proceso
#>

param(
    [string]$AccdbPath      = (Join-Path $PSScriptRoot "..\..\..\..\Pitaya_System.accdb"),
    [switch]$ExportTableData,          # false por defecto (puede generar GBs de datos)
    [int]   $MaxRowsExport  = 50000    # limite de filas por tabla en modo -ExportTableData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── Rutas (el script vive en scripts\, la salida va al directorio padre) ────
$Root      = Split-Path $PSScriptRoot -Parent
$DbDir     = Join-Path $Root "db"
$VbaDir    = Join-Path $Root "vba"
$FolmsDir  = Join-Path $VbaDir "formularios"
$InformDir = Join-Path $VbaDir "informes"
$ModDir    = Join-Path $VbaDir "modulos"
$ClsDir    = Join-Path $VbaDir "clases"
$SqlDir    = Join-Path $Root "sql"
$QryDir    = Join-Path $SqlDir "queries"
$TblDir    = Join-Path $SqlDir "tables"
$DataDir   = Join-Path $SqlDir "data"
$LogFile   = Join-Path $Root "export_log.txt"

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
        $fs = [System.IO.File]::Open($LogFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($line)
        $sw.Close(); $fs.Close()
    } catch { }
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
function Count-Files ([string]$Dir, [string]$Filter = "*.sql") {
    $f = Get-ChildItem $Dir -Filter $Filter -ErrorAction SilentlyContinue
    if ($f) { return $f.Count } else { return 0 }
}

# ─── Mapa tipos DAO → SQL ───────────────────────────────────────────────────
$typeMap = @{
    1  = 'BIT'             # dbBoolean      (Sí/No)
    2  = 'BYTE'            # dbByte         (Número/Byte)
    3  = 'SMALLINT'        # dbInteger      (Número/Entero)
    4  = 'LONG'            # dbLong         (Número/Entero largo)
    5  = 'CURRENCY'        # dbCurrency     (Moneda)
    6  = 'SINGLE'          # dbSingle       (Número/Simple)
    7  = 'DOUBLE'          # dbDouble       (Número/Doble)
    8  = 'DATETIME'        # dbDate         (Fecha/Hora)
    9  = 'BINARY'          # dbBinary
    10 = 'VARCHAR'         # dbText         (Texto corto) -- se agrega tamaño abajo
    11 = 'LONGBINARY'      # dbLongBinary   (Objeto OLE)
    12 = 'MEMO'            # dbMemo         (Texto largo)
    15 = 'UNIQUEIDENTIFIER'# dbGUID
    16 = 'BIGINT'          # dbBigInt
    101= 'AUTOINCREMENT'   # Autonumérico
}

# ─────────────────────────────────────────────────────────────────────────────
#  INICIO
# ─────────────────────────────────────────────────────────────────────────────
if (Test-Path $LogFile) { Remove-Item $LogFile -Force }

Write-Log "==========================================================" "Yellow"
Write-Log "  PITAYA SYSTEM ACCESS -- Exportador Completo" "Yellow"
Write-Log "  VBA  |  Consultas SQL  |  DDL Tablas  |  $(if ($ExportTableData) {'Datos'} else {'(Datos: usar -ExportTableData)'})" "Yellow"
Write-Log "==========================================================" "Yellow"

# Resolver ruta del .accdb fuente
$AccdbResolved = $null
try { $AccdbResolved = (Resolve-Path $AccdbPath -ErrorAction Stop).Path } catch {}
if (-not $AccdbResolved) {
    Write-Log "ERROR: No se encontro: $AccdbPath" "Red"
    exit 1
}
Write-Log "Origen  : $AccdbResolved" "Green"
Write-Log "Destino : $Root" "Green"

# ── 1. LIMPIAR INSTANCIAS ANTERIORES DE ACCESS ───────────────────────────────
Write-Log "" "White"
Write-Log "-- [1/7] Limpiando procesos Access anteriores --" "Magenta"
$prevProcs = Get-Process -Name MSACCESS -ErrorAction SilentlyContinue
if ($prevProcs) {
    $prevProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Log "  $($prevProcs.Count) proceso(s) Access detenidos." "DarkGray"
} else {
    Write-Log "  No habia procesos Access activos." "DarkGray"
}

# ── 2. CREAR CARPETAS ─────────────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "-- [2/7] Creando estructura de carpetas --" "Magenta"
$dirsToCreate = @($DbDir, $FolmsDir, $InformDir, $ModDir, $ClsDir, $QryDir, $TblDir)
if ($ExportTableData) { $dirsToCreate += $DataDir }
$dirsToCreate | ForEach-Object { Ensure-Dir $_ }

# ── 3. COPIAR .ACCDB ──────────────────────────────────────────────────────────
Write-Log "" "White"
Write-Log "-- [3/7] Copiando base de datos --" "Magenta"
$DbDest = Join-Path $DbDir "Pitaya_System.accdb"
Copy-Item -Path $AccdbResolved -Destination $DbDest -Force
$sizeMb = [math]::Round((Get-Item $DbDest).Length / 1MB, 1)
Write-Log "  Copiado: db\Pitaya_System.accdb  ($sizeMb MB)" "Green"

# ── 4. LANZAR ACCESS Y CONECTAR VIA COM ──────────────────────────────────────
Write-Log "" "White"
Write-Log "-- [4/7] Conectando a Microsoft Access --" "Magenta"

$accessProc = $null
$access      = $null
$counters    = @{ Forms=0; Reports=0; Modules=0; Classes=0; Skipped=0; Errors=0 }

try {
    # Lanzar Access sin abrir ninguna BD
    if (-not (Test-Path $AccessExe)) {
        Write-Log "  ERROR: No se encontro MSACCESS.EXE en: $AccessExe" "Red"
        exit 1
    }
    Write-Log "  Lanzando Access..." "DarkCyan"
    $accessProc = Start-Process -FilePath $AccessExe -PassThru

    # Esperar hasta que Access responda al COM
    $maxWait = 25; $waited = 0
    while ($waited -lt $maxWait) {
        Start-Sleep -Seconds 2; $waited += 2
        try {
            $access = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Access.Application")
            Write-Log "  Access listo en ${waited}s (v$($access.Version))" "Green"
            break
        } catch { }
    }
    if ($null -eq $access) {
        Write-Log "  ERROR: No se pudo conectar a Access tras ${maxWait}s." "Red"
        exit 1
    }

    $access.AutomationSecurity = 3
    Start-Sleep -Seconds 2

    Write-Log "  Abriendo BD de trabajo..." "DarkCyan"
    $access.OpenCurrentDatabase($DbDest, $false)
    Start-Sleep -Seconds 3

    $vbProject = $access.VBE.ActiveVBProject
    $db        = $access.CurrentDb()

    # ── 5. EXTRAER VBA ───────────────────────────────────────────────────────
    Write-Log "" "White"
    Write-Log "-- [5/7] Extrayendo modulos VBA --" "Magenta"

    if ($null -eq $vbProject) {
        Write-Log "  ADVERTENCIA: No se encontro proyecto VBA." "Yellow"
    } else {
        Write-Log "  Proyecto: '$($vbProject.Name)'  |  Componentes: $($vbProject.VBComponents.Count)" "DarkCyan"
        foreach ($comp in $vbProject.VBComponents) {
            $compName = $comp.Name
            $compType = $comp.Type
            try {
                $lines = $comp.CodeModule.CountOfLines
                if ($lines -eq 0) { $counters.Skipped++; continue }
                $code = $comp.CodeModule.Lines(1, $lines)

                if      ($compName -like "Form_*")   { $destDir=$FolmsDir; $ext=".bas"; $counters.Forms++ }
                elseif  ($compName -like "Report_*") { $destDir=$InformDir; $ext=".bas"; $counters.Reports++ }
                elseif  ($compType -eq 2)            { $destDir=$ClsDir; $ext=".cls"; $counters.Classes++ }
                else                                 { $destDir=$ModDir; $ext=".bas"; $counters.Modules++ }

                $header = "' ==========================================================`r`n" +
                          "' Modulo  : $compName`r`n" +
                          "' Tipo    : $compType  |  Lineas: $lines`r`n" +
                          "' Proyecto: $($vbProject.Name)`r`n" +
                          "' Exportado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n" +
                          "' ==========================================================`r`n"

                Set-Content -Path (Join-Path $destDir ((Safe-FileName $compName) + $ext)) `
                            -Value ($header + "`r`n" + $code) -Encoding UTF8
                Write-Log "  VBA  $(Safe-FileName $compName)$ext  ($lines ln)" "DarkGray"
            } catch {
                Write-Log "  WARN VBA '$compName': $_" "Yellow"
                $counters.Errors++
            }
        }
    }

    # ── 6a. EXTRAER SQL DE CONSULTAS ─────────────────────────────────────────
    Write-Log "" "White"
    Write-Log "-- [6/7] Extrayendo SQL de consultas y estructura de tablas --" "Magenta"
    $qryCount = 0
    for ($i = 0; $i -lt $db.QueryDefs.Count; $i++) {
        $qdef = $db.QueryDefs($i)
        if ($qdef.Name.StartsWith("~")) { continue }
        try {
            $hdr = "-- ==========================================================`r`n" +
                   "-- Consulta : $($qdef.Name)`r`n" +
                   "-- Exportado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n" +
                   "-- ==========================================================`r`n`r`n"
            Set-Content -Path (Join-Path $QryDir ((Safe-FileName $qdef.Name) + ".sql")) `
                        -Value ($hdr + $qdef.SQL) -Encoding UTF8
            $qryCount++
            Write-Log "  QRY  $(Safe-FileName $qdef.Name).sql" "DarkGray"
        } catch { Write-Log "  WARN QRY '$($qdef.Name)': $_" "Yellow" }
    }
    Write-Log "  Consultas exportadas: $qryCount" "Cyan"

    # ── 6b. EXTRAER DDL DE TABLAS (CREATE TABLE) ──────────────────────────────
    $tblCount = 0
    foreach ($tdef in $db.TableDefs) {
        $tName = $tdef.Name
        if ($tName.StartsWith("MSys") -or $tName.StartsWith("~")) { continue }
        try {
            $out = [System.Text.StringBuilder]::new()
            $out.AppendLine("-- ==========================================================") | Out-Null
            $out.AppendLine("-- Tabla    : $tName") | Out-Null
            $out.AppendLine("-- Campos   : $($tdef.Fields.Count)") | Out-Null
            $out.AppendLine("-- Exportado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
            $out.AppendLine("-- ==========================================================") | Out-Null
            $out.AppendLine("") | Out-Null
            $out.AppendLine("CREATE TABLE [$tName] (") | Out-Null

            $fldLines = @()
            foreach ($fld in $tdef.Fields) {
                $ft   = $fld.Type
                $tn   = if ($typeMap.ContainsKey($ft)) { $typeMap[$ft] } else { "TYPE_$ft" }
                if ($ft -eq 10) { $tn = "VARCHAR($($fld.Size))" }
                $nn   = if ($fld.Required) { " NOT NULL" } else { "" }
                $fldLines += "    [$($fld.Name)] $tn$nn"
            }
            $out.AppendLine(($fldLines -join ",`r`n")) | Out-Null
            $out.AppendLine(");") | Out-Null
            $out.AppendLine("") | Out-Null

            foreach ($idx in $tdef.Indexes) {
                if ($idx.Primary) {
                    $pk = ($idx.Fields | ForEach-Object { "[$($_.Name)]" }) -join ", "
                    $out.AppendLine("ALTER TABLE [$tName] ADD PRIMARY KEY ($pk);") | Out-Null
                    $out.AppendLine("") | Out-Null
                }
            }

            Set-Content -Path (Join-Path $TblDir ((Safe-FileName $tName) + ".sql")) `
                        -Value $out.ToString() -Encoding UTF8
            $tblCount++
            Write-Log "  DDL  $(Safe-FileName $tName).sql" "DarkGray"
        } catch { Write-Log "  WARN DDL '$tName': $_" "Yellow" }
    }
    Write-Log "  Tablas DDL exportadas: $tblCount" "Cyan"

    # ── 7. DATOS DE TABLAS VIA PYODBC (opcional) ────────────────────────────
    if ($ExportTableData) {
        Write-Log "" "White"
        Write-Log "-- [7/7] Extrayendo datos de tablas (INSERT INTO) --" "Magenta"
        Write-Log "  Limite: $MaxRowsExport filas por tabla." "Yellow"
        Ensure-Dir $DataDir

        $pyScript = Join-Path $env:TEMP "pitaya_export_data.py"
        $dbEsc    = $DbDest.Replace('\', '\\')
        $dataEsc  = $DataDir.Replace('\', '\\')

        $pyCode = @"
import pyodbc, os, re
from datetime import datetime

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$dbEsc;')
cursor = conn.cursor()
tables = [t.table_name for t in cursor.tables(tableType='TABLE')
          if not t.table_name.startswith('MSys')]
print(f'Tablas encontradas: {len(tables)}')

MAX_ROWS = $MaxRowsExport
for tname in tables:
    try:
        cursor.execute(f'SELECT COUNT(*) FROM [{tname}]')
        total = cursor.fetchone()[0]
        if total > MAX_ROWS:
            print(f'  SKIP {tname}  ({total} filas > limite {MAX_ROWS})')
            continue
        cursor.execute(f'SELECT * FROM [{tname}]')
        rows = cursor.fetchall()
        cols = [d[0] for d in cursor.description]
        safe = re.sub(r'[<>:"/\\|?*]', '_', tname)
        with open(os.path.join(r'$dataEsc', safe + '.sql'), 'w', encoding='utf-8') as f:
            f.write(f'-- Tabla: {tname}  |  Filas: {len(rows)}\n')
            f.write(f'-- Exportado: {datetime.now():%Y-%m-%d %H:%M:%S}\n')
            f.write('-- ===========================================================\n\n')
            if rows:
                cols_str = ', '.join(f'[{c}]' for c in cols)
                for row in rows:
                    vals = []
                    for v in row:
                        if v is None: vals.append('NULL')
                        elif isinstance(v, str): vals.append("'" + v.replace("'","''") + "'")
                        elif isinstance(v, (int,float)): vals.append(str(v))
                        else: vals.append("'" + str(v) + "'")
                    f.write(f'INSERT INTO [{tname}] ({cols_str}) VALUES ({", ".join(vals)});\n')
        print(f'  OK  {safe}.sql  ({len(rows)} filas)')
    except Exception as e:
        print(f'  WARN {tname}: {e}')

conn.close()
print('DONE')
"@
        Set-Content -Path $pyScript -Value $pyCode -Encoding UTF8
        Write-Log "  Ejecutando Python+pyodbc..." "DarkCyan"
        python $pyScript 2>&1 | ForEach-Object { Write-Log "  $_" "DarkGray" }
        Remove-Item $pyScript -Force -ErrorAction SilentlyContinue
    }

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
        Write-Log "  Access cerrado (PID $($accessProc.Id))." "DarkGray"
    }
}

# ─── RESUMEN FINAL ────────────────────────────────────────────────────────────
$qForms = Count-Files $FolmsDir "*.bas"
$qReps  = Count-Files $InformDir "*.bas"
$qMods  = Count-Files $ModDir "*.bas"
$qCls   = Count-Files $ClsDir "*.cls"
$qQry   = Count-Files $QryDir "*.sql"
$qTbl   = Count-Files $TblDir "*.sql"
$qDat   = Count-Files $DataDir "*.sql"

Write-Log "" "White"
Write-Log "==========================================================" "Yellow"
Write-Log "  EXPORTACION COMPLETADA" "Yellow"
Write-Log "  VBA Formularios : $qForms    vba\formularios\" "Green"
Write-Log "  VBA Informes    : $qReps     vba\informes\" "Green"
Write-Log "  VBA Modulos     : $qMods     vba\modulos\" "Green"
Write-Log "  VBA Clases      : $qCls      vba\clases\" "Green"
Write-Log "  SQL Consultas   : $qQry      sql\queries\" "Green"
Write-Log "  SQL Tablas DDL  : $qTbl      sql\tables\   (CREATE TABLE)" "Green"
if ($ExportTableData) {
    Write-Log "  SQL Datos       : $qDat      sql\data\     (INSERT INTO)" "Green"
} else {
    Write-Log "  SQL Datos       : (usar -ExportTableData para exportar)" "DarkGray"
}
if ($counters.Errors -gt 0) {
    Write-Log "  Advertencias    : $($counters.Errors)   (ver export_log.txt)" "Yellow"
}
Write-Log "  Log: export_log.txt" "Cyan"
Write-Log "==========================================================" "Yellow"
