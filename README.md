# Pitaya System Access — Control de versiones VBA + SQL

Control de versiones del código VBA, consultas SQL y estructura de tablas  
extraídos de `Pitaya_System.accdb` mediante PowerShell y Git.

> **El archivo `.accdb` NO está en este repositorio** (~216 MB).
> Vive en `C:\Users\<usuario>\Desktop\Sistema\Pitaya_System.accdb`.
> Los scripts trabajan sobre una **copia local** — el original nunca se
> toca automáticamente.

📖 **Para el flujo de trabajo completo, escenarios y referencia detallada, ver [MANUAL.md](MANUAL.md).**

---

## Estructura del repositorio

```
Pitaya_System_Access/
│
├── scripts/
│   ├── Export-AccessVBA.ps1    ← Desglosar .accdb → archivos de código
│   ├── Import-AccessVBA.ps1    ← Reensamblar archivos de código → .accdb
│   └── gitpush.ps1             ← Subir cambios a GitHub
│
├── vba/
│   ├── formularios/            ← Form_*.bas   (código VBA de formularios)
│   ├── informes/               ← Report_*.bas (código VBA de informes)
│   ├── modulos/                ← módulos estándar .bas
│   └── clases/                 ← módulos de clase .cls
│
├── sql/
│   ├── queries/                ← SQL de consultas guardadas en Access
│   └── tables/                 ← DDL CREATE TABLE de cada tabla
│
├── db/                         ← ⚠️ IGNORADO POR GIT — copia de trabajo
├── MANUAL.md                   ← Guía completa de uso y flujo de trabajo
├── README.md                   ← Este archivo
└── .gitignore
```

---

## Requisitos

| Requisito | Detalle |
|---|---|
| **Windows** | 10 / 11 (64-bit) |
| **Microsoft Access** | 2016 / 2019 / 2021 / 365 (Office 16.0) |
| **PowerShell** | 5.1 o superior (incluido en Windows) |
| **Python + pyodbc** | Solo si usas `-ExportTableData` |
| **Pitaya_System.accdb** | En `C:\...\Desktop\Sistema\` |

> **Office Click-to-Run (Microsoft 365 / 2024):** el script lanza Access
> como proceso y se conecta automáticamente. No requiere configuración adicional.

---

## Estructura de carpetas esperada en tu PC

El repositorio debe clonarse dentro de esta estructura para que las rutas
automáticas funcionen:

```
C:\Users\<usuario>\Desktop\Sistema\
├── Pitaya_System.accdb              ← archivo oficial
└── Pitaya Web\
    └── VisualCode\
        └── Pitaya_System_Access\    ← este repositorio
            └── scripts\
                └── Export-AccessVBA.ps1
```

Con esta estructura, los scripts encuentran el `.accdb` automáticamente.
Si está en otra ubicación, usa `-AccdbPath` al llamar el script.

---

## Primera instalación (PC nueva)

```powershell
# 1. Ir a la carpeta correcta ANTES de clonar
cd "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya Web\VisualCode"

# 2. Clonar
git clone https://github.com/MiguelGotea/Pitaya_System_Access.git
cd Pitaya_System_Access

# 3. Verificar que el .accdb existe
Test-Path "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya_System.accdb"

# 4. Exportar para sincronizar todo
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1
```

---

## Los tres scripts

### Export — Desglosar el .accdb

```powershell
# Exportar VBA + consultas + DDL de tablas (uso normal)
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1

# También exportar datos de tablas como INSERT INTO
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1 -ExportTableData
```

**Qué exporta:**

| Tipo | Carpeta | Bidireccional |
|---|---|:---:|
| Código VBA formularios | `vba/formularios/` | ✅ |
| Código VBA informes | `vba/informes/` | ✅ |
| Módulos estándar | `vba/modulos/` | ✅ |
| Módulos de clase | `vba/clases/` | ✅ |
| SQL de consultas | `sql/queries/` | ⚠️ Solo ida |
| DDL estructura tablas | `sql/tables/` | ❌ Solo docs |

> ⚠️ **Nota importante:** el diseño visual de formularios e informes
> (controles, layouts, propiedades) **no se exporta**. Solo el código VBA.

---

### Import — Reensamblar en la copia de trabajo

```powershell
# Importar todos los archivos .bas/.cls
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1

# Solo los archivos modificados según git (más rápido)
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1 -OnlyChanged

# Ver qué se importaría sin hacer cambios
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1 -DryRun
```

> **El Import escribe en `db\Pitaya_System.accdb` (la copia), no en el original.**
> Cuando hayas verificado que funciona, copia manualmente
> `db\Pitaya_System.accdb` → `C:\...\Sistema\Pitaya_System.accdb`.

---

### Push — Subir a GitHub

```powershell
# Con mensaje automático
powershell -ExecutionPolicy Bypass -File .\scripts\gitpush.ps1

# Con mensaje descriptivo (recomendado)
powershell -ExecutionPolicy Bypass -File .\scripts\gitpush.ps1 "feat: descripción del cambio"
```

---

## Flujo de trabajo (resumen)

```
1. Export   → desglosar .accdb actual en vba\ y sql\
2. Editar   → modificar .bas en VS Code (con IA o manualmente)
3. Import   → reensamblar cambios en db\Pitaya_System.accdb
4. Verificar→ abrir db\ en Access y probar
5. Promover → copiar db\ → Sistema\ (manual, cuando estés seguro)
6. Export   → re-sincronizar vba\ con el estado final del .accdb
7. Push     → documentar cambios en GitHub con historial
```

> ⚠️ **El paso 6 (re-export antes del push) es importante.**
> Access puede hacer ajustes automáticos al importar. El re-export asegura
> que git refleja el estado real del `.accdb`.

📖 Ver [MANUAL.md](MANUAL.md) para el flujo detallado, escenarios y referencia completa.

---

## Qué incluye y qué excluye Git

| ✅ Incluido en Git | ❌ Excluido (.gitignore) |
|---|---|
| Código VBA `.bas` / `.cls` | `db\` — el .accdb de trabajo (~216 MB) |
| SQL de consultas `.sql` | `*.log` — logs locales |
| DDL de tablas `.sql` | `.DS_Store`, `Thumbs.db` |
| Los tres scripts `.ps1` | `.vscode/`, `.idea/` |
| `README.md`, `MANUAL.md`, `.gitignore` | |

---

## Solución de problemas

### Error: "No se encontró el archivo .accdb"
```
Verifica que Pitaya_System.accdb existe en:
C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya_System.accdb

O pasa la ruta manualmente:
.\scripts\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\Pitaya_System.accdb"
```

### Access no se conecta / timeout
```
1. Cierra cualquier instancia de Access abierta manualmente
2. Vuelve a ejecutar el script
Los scripts gestionan su propia instancia de Access automáticamente.
```

### "La directiva de ejecución no permite..."
```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1
```

### "La base de datos ya está abierta"
```
Ocurre si una ejecución anterior quedó colgada.
Los scripts limpian instancias anteriores automáticamente al inicio.
Si persiste, cierra Access desde el Administrador de tareas y reintenta.
```

---

## Estadísticas del proyecto

| Categoría | Cantidad |
|---|---|
| Formularios (`Form_*.bas`) | 377 |
| Informes (`Report_*.bas`) | 36 |
| Módulos estándar (`.bas`) | 40 |
| SQL Consultas (`.sql`) | 58 |
| SQL Tablas DDL (`.sql`) | 104 |
| **Total archivos versionados** | **615** |

---

## Repositorio

- **GitHub:** [MiguelGotea/Pitaya_System_Access](https://github.com/MiguelGotea/Pitaya_System_Access)
- **Base de datos:** `Pitaya_System.accdb` — sistema principal Pitaya (Access 2024)
- **Manual completo:** [MANUAL.md](MANUAL.md)
