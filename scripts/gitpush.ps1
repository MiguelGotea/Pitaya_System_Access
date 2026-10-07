# PowerShell - Push rapido a GitHub (Pitaya_System_Access)
# Uso: .\scripts\gitpush.ps1 ["mensaje opcional"]

# Forzar codificación UTF8 para emojis
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Auto-navegar a la raíz del repo (scripts\ -> Pitaya_System_Access\)
Set-Location $PSScriptRoot
Set-Location ..

# Script Tanque v7 (Anti-Choque)
git add .
$msg = $args[0]
if (-not $msg) { $msg = "Human Push $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" }

git commit -m "$msg" 2>$null

Write-Host "🚀 Intentando sincronizar y subir cambios a GitHub..." -ForegroundColor Cyan
git pull origin main --rebase

if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️ Conflicto detectado. Aplicando reparación de Hierro..." -ForegroundColor Yellow
    git rebase --abort 2>$null
    git pull origin main --no-rebase -X ours
    git add .
    git commit -m "$msg (Conflict Resolved)" 2>$null
}

git push origin main

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ ¡Subida completada con éxito! (Repositorio actualizado en GitHub)" -ForegroundColor Green
} else {
    Write-Host "❌ Error en push." -ForegroundColor Red
}
