# Pitaya System Access — Repositorio VBA + SQL

Control de versiones del código VBA y las consultas SQL extraídos de `Pitaya_System.accdb`.

> **El archivo `.accdb` NO está en este repositorio** (pesa ~216 MB).  
> Vive en `C:\Users\<usuario>\Desktop\Sistema\Pitaya_System.accdb` y se copia localmente al ejecutar el script de exportación.

---

## Estructura del repositorio

```
Pitaya_System_Access/
├── scripts/
│   └── Export-AccessVBA.ps1   ← script de exportación (el que genera todo)
├── vba/
│   ├── formularios/           ← Form_*.bas   (código detrás de cada formulario)
│   ├── informes/              ← Report_*.bas (código detrás de cada informe)
│   ├── modulos/               ← módulos estándar .bas (AccessHostinger, Delivery, etc.)
│   └── clases/                ← módulos de clase .cls (si existen)
├── sql/
│   └── queries/               ← SQL de cada consulta definida en Access (.sql)
├── db/                        ← IGNORADO por git — copia local del .accdb
├── export_log.txt             ← IGNORADO por git — log del último proceso
├── .gitignore
└── README.md
```

---

## Requisitos

| Requisito | Detalle |
|---|---|
| **Windows** | 10 / 11 (64-bit) |
| **Microsoft Access** | 2016 / 2019 / 2021 / 365 (cualquier versión con Office 16.0) |
| **PowerShell** | 5.1 o superior (incluido en Windows) |
| **Pitaya_System.accdb** | Debe existir en la ruta configurada (ver abajo) |

> **Office Click-to-Run (Microsoft 365 / 2024):** el script lanza Access como proceso y se conecta a él automáticamente. No requiere configuración adicional.

---

## Configuración de ruta del .accdb

### Ruta por defecto (estructura estándar de Pitaya)

El script asume que el repositorio se clona dentro de la siguiente estructura:

```
C:\Users\<usuario>\Desktop\Sistema\
├── Pitaya_System.accdb             ← archivo original
└── Pitaya Web\
    └── VisualCode\
        └── Pitaya_System_Access\   ← este repositorio
            └── scripts\
                └── Export-AccessVBA.ps1
```

Con esta estructura, el script encuentra `Pitaya_System.accdb` automáticamente subiendo **4 niveles** desde `scripts\`.

### Si el .accdb está en otra ubicación

Pasa la ruta como parámetro al ejecutar:

```powershell
.\scripts\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\personalizada\Pitaya_System.accdb"
```

---

## Cómo usar — Primera vez (o nueva PC)

### 1. Clonar el repositorio

```powershell
# Navega a la carpeta correcta ANTES de clonar
cd "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya Web\VisualCode"

git clone https://github.com/MiguelGotea/Pitaya_System_Access.git
```

### 2. Verificar que el .accdb existe

```powershell
# Debe mostrar el archivo (no error)
Test-Path "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya_System.accdb"
```

### 3. Ejecutar el script de exportación

```powershell
cd Pitaya_System_Access

# Ejecución normal (ruta automática)
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1

# O si el .accdb está en otra ruta
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\Pitaya_System.accdb"
```

El script:
1. Copia `Pitaya_System.accdb` → `db\Pitaya_System.accdb` (ignorado por git)
2. Lanza Microsoft Access en segundo plano
3. Extrae **todo el código VBA** a `vba\formularios\`, `vba\informes\`, `vba\modulos\`, `vba\clases\`
4. Extrae **el SQL de todas las consultas** a `sql\queries\`
5. Cierra Access automáticamente
6. Genera `export_log.txt` con el resumen detallado

### 4. Subir los cambios a GitHub

```powershell
git add .
git commit -m "feat: exportar VBA y SQL - $(Get-Date -Format 'yyyy-MM-dd')"
git push origin main
```

---

## Flujo de trabajo diario

```
Modificas código en Access
        ↓
Guardas los cambios en Access
        ↓
Ejecutas Export-AccessVBA.ps1
        ↓
git add . → git commit → git push
```

---

## Qué incluye y qué excluye Git

| ✅ Incluido en Git | ❌ Excluido (en .gitignore) |
|---|---|
| Todo el código VBA `.bas` / `.cls` | `db\` — el archivo .accdb (~216 MB) |
| SQL de consultas `.sql` | `export_log.txt` — log local |
| `scripts\Export-AccessVBA.ps1` | `*.log` — cualquier log |
| `README.md` / `.gitignore` | `.DS_Store`, `Thumbs.db` |

---

## Descripción de componentes VBA exportados

### `vba\formularios\` — Formularios (Form_*)
Código VBA detrás de cada formulario de Access. Equivale a lo que ves en el VBA Editor (Alt+F11) bajo cada `Form_NombreFormulario`.

Ejemplos: `Form_Menu Principal.bas`, `Form_PanelDescargaExcelConsumoDirecto.bas`

### `vba\informes\` — Informes (Report_*)
Código VBA detrás de cada informe de Access.

Ejemplos: `Report_Boleta.bas`, `Report_VentasCliente.bas`

### `vba\modulos\` — Módulos estándar
Módulos de código reutilizable (funciones y subrutinas globales).

Ejemplos: `AccessHostinger.bas`, `Delivery.bas`, `modulo_sync_ventas.bas`

### `sql\queries\` — Consultas SQL
SQL de cada consulta guardada en Access (equivalente a las Queries en el panel de navegación).

Ejemplos: `InformeDiario.sql`, `ResumenVentasMesExcel.sql`

---

## Solución de problemas

### Error: "No se encontró el archivo .accdb"
```
Verifica que Pitaya_System.accdb existe en:
C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya_System.accdb

O pasa la ruta manualmente con -AccdbPath
```

### Access no se conecta / timeout
```
1. Cierra Access si está abierto manualmente
2. Vuelve a ejecutar el script
El script abre su propia instancia de Access automáticamente.
```

### "La directiva de ejecución no permite..."
```powershell
# Solución: usar el flag -ExecutionPolicy Bypass al llamar el script
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1
```

### Módulos duplicados entre ejecuciones
El script sobreescribe los archivos en cada ejecución. No hay duplicados — siempre refleja el estado actual del .accdb.

---

## Estadísticas del proyecto (última exportación)

| Categoría | Cantidad |
|---|---|
| Formularios (`Form_*.bas`) | 377 |
| Informes (`Report_*.bas`) | 36 |
| Módulos estándar (`.bas`) | 40 |
| Consultas SQL (`.sql`) | 58 |
| **Total archivos** | **511** |

---

## Repositorio relacionado

- **GitHub:** [MiguelGotea/Pitaya_System_Access](https://github.com/MiguelGotea/Pitaya_System_Access)
- **Base de datos:** `Pitaya_System.accdb` — sistema principal Pitaya (Access 2024)
