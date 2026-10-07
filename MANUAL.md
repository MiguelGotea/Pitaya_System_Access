# Manual de Uso — Pitaya System Access

> Guía de flujo de trabajo para el control de versiones del código VBA y SQL  
> de `Pitaya_System.accdb` mediante este repositorio.

---

## Índice

1. [Conceptos clave](#conceptos-clave)
2. [Estructura del repositorio](#estructura-del-repositorio)
3. [Los tres scripts](#los-tres-scripts)
4. [Qué se exporta y qué se importa](#qué-se-exporta-y-qué-se-importa)
5. [Flujo de trabajo principal](#flujo-de-trabajo-principal)
6. [Escenarios de uso](#escenarios-de-uso)
7. [Versionado con Git](#versionado-con-git)
8. [Advertencias importantes](#advertencias-importantes)

---

## Conceptos clave

### El archivo original nunca se toca automáticamente

Todos los scripts trabajan sobre una **copia de trabajo** en `db\`. El archivo oficial
`C:\...\Sistema\Pitaya_System.accdb` solo se modifica **manualmente por ti** cuando
has verificado que los cambios funcionan correctamente.

```
Sistema\Pitaya_System.accdb   ← OFICIAL. Solo tú lo modificas manualmente.
Pitaya_System_Access\db\      ← COPIA DE TRABAJO. Los scripts trabajan aquí.
                                 Ignorada por git (no sube a GitHub).
```

### Git trackea código, no el .accdb

El archivo `.accdb` (binario, ~216 MB) nunca sube a GitHub. Lo que sube son
los archivos de texto extraídos de él:

- `.bas` — código VBA de formularios, informes y módulos
- `.cls` — módulos de clase VBA
- `.sql` — SQL de consultas y estructura de tablas

Git compara el **contenido real** de cada archivo. Si un módulo no cambió
entre exportaciones, git no registra nada para ese archivo. Solo documenta
los cambios reales.

---

## Estructura del repositorio

```
Pitaya_System_Access/
│
├── scripts/
│   ├── Export-AccessVBA.ps1    ← Desglosar .accdb en archivos de código
│   ├── Import-AccessVBA.ps1    ← Reensamblar archivos de código en .accdb
│   └── gitpush.ps1             ← Subir cambios a GitHub
│
├── vba/
│   ├── formularios/            ← Form_*.bas  (código detrás de formularios)
│   ├── informes/               ← Report_*.bas (código detrás de informes)
│   ├── modulos/                ← módulos estándar .bas
│   └── clases/                 ← módulos de clase .cls
│
├── sql/
│   ├── queries/                ← SQL de consultas guardadas en Access
│   └── tables/                 ← DDL CREATE TABLE de cada tabla
│
├── db/                         ← ⚠️ IGNORADO POR GIT — copia de trabajo
│   └── Pitaya_System.accdb
│
├── README.md                   ← Instalación y requisitos
├── MANUAL.md                   ← Este archivo
└── .gitignore
```

---

## Los tres scripts

### `Export-AccessVBA.ps1` — Desglosar
```powershell
# Uso normal (extrae VBA + consultas SQL + estructura de tablas)
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1

# Incluir también datos de tablas como INSERT INTO (puede ser grande)
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1 -ExportTableData

# Con ruta personalizada al .accdb fuente
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1 -AccdbPath "C:\ruta\Pitaya_System.accdb"
```

**Qué hace:**
1. Copia `Sistema\Pitaya_System.accdb` → `db\Pitaya_System.accdb`
2. Abre Access en segundo plano (sin ventana visible)
3. Extrae todo el código VBA a `vba\`
4. Extrae el SQL de todas las consultas a `sql\queries\`
5. Genera el DDL (CREATE TABLE) de cada tabla a `sql\tables\`
6. Cierra Access automáticamente

---

### `Import-AccessVBA.ps1` — Reensamblar
```powershell
# Importar todos los archivos .bas/.cls a la copia de trabajo
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1

# Solo los archivos modificados según git (más rápido)
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1 -OnlyChanged

# Simular sin hacer cambios (para revisar qué se importaría)
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1 -DryRun

# Importar a un .accdb específico (no la copia de trabajo)
powershell -ExecutionPolicy Bypass -File .\scripts\Import-AccessVBA.ps1 -AccdbTarget "C:\ruta\destino.accdb"
```

**Qué hace:**
1. Abre `db\Pitaya_System.accdb` (la copia de trabajo, NO el original)
2. Para cada `.bas` / `.cls` modificado, inyecta el código en el proyecto VBA
3. Cierra Access

> ⚠️ El resultado queda en `db\Pitaya_System.accdb`. Tú decides cuándo
> copiarlo manualmente al Sistema oficial.

---

### `gitpush.ps1` — Subir a GitHub
```powershell
# Push con mensaje automático (fecha y hora)
powershell -ExecutionPolicy Bypass -File .\scripts\gitpush.ps1

# Push con mensaje descriptivo (recomendado)
powershell -ExecutionPolicy Bypass -File .\scripts\gitpush.ps1 "feat: nuevo módulo de sincronización"
```

**Qué hace:**
1. `git add .` — incluye todos los cambios
2. `git commit` — crea el snapshot con el mensaje indicado
3. `git pull --rebase` — sincroniza con el remoto primero
4. `git push` — sube a GitHub
5. Si hay conflictos, los resuelve automáticamente favoreciendo tu versión local

---

## Qué se exporta y qué se importa

Esta es la tabla más importante para entender el alcance del sistema:

| Tipo | Carpeta | ¿Se exporta? | ¿Se importa? | Notas |
|---|---|:---:|:---:|---|
| Código VBA de formularios | `vba/formularios/` | ✅ | ✅ | Solo el código. El diseño (controles, propiedades) permanece intacto. |
| Código VBA de informes | `vba/informes/` | ✅ | ✅ | Solo el código. El diseño del informe permanece intacto. |
| Módulos estándar | `vba/modulos/` | ✅ | ✅ | Reemplazo completo del módulo. |
| Módulos de clase | `vba/clases/` | ✅ | ✅ | Reemplazo completo de la clase. |
| SQL de consultas guardadas | `sql/queries/` | ✅ | ❌ | Solo documentación/versionado. No se reimportan a Access. |
| DDL de estructura de tablas | `sql/tables/` | ✅ | ❌ | Solo documentación. Las tablas no se recrean desde aquí. |
| Datos de tablas | `sql/data/` | ⚙️ Opcional | ❌ | Solo con `-ExportTableData`. Solo documentación. |
| Diseño de formularios/informes | — | ❌ | ❌ | Vive en la estructura binaria del .accdb. No se puede extraer como texto. |
| Relaciones entre tablas | — | ❌ | ❌ | Solo en el .accdb. |
| Macros de Access | — | ❌ | ❌ | Solo en el .accdb. |

### En resumen: ¿qué abarca el reensamblado?

El Import reconstruye **todo el código VBA** del proyecto. Lo que **no** abarca:
- El diseño visual de formularios e informes (controles, layouts, propiedades)
- Las consultas SQL (se documentan pero no se reimportan)
- Las tablas y sus datos
- Las relaciones entre tablas
- Las macros

Para cambiar esas cosas, se trabaja directamente en Access.

---

## Flujo de trabajo principal

### Flujo estándar (el más común)

```
┌─────────────────────────────────────────────────────────────────┐
│  1. EXPORT — Desglosar el .accdb actual                         │
│                                                                 │
│     .\scripts\Export-AccessVBA.ps1                              │
│                                                                 │
│     Sistema\Pitaya_System.accdb  ──copia──▶  db\               │
│                                       extrae▶  vba\ y sql\     │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  2. EDITAR — Modificar el código en VS Code                     │
│                                                                 │
│     Editas vba\modulos\MiModulo.bas, vba\formularios\Form_X.bas │
│     Puedes usar IA (Antigravity) para hacer cambios grandes.    │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  3. IMPORT — Reensamblar en la copia de trabajo                 │
│                                                                 │
│     .\scripts\Import-AccessVBA.ps1 -OnlyChanged                 │
│                                                                 │
│     vba\ ──inyecta código──▶  db\Pitaya_System.accdb            │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  4. VERIFICAR — Probar el .accdb reensamblado                   │
│                                                                 │
│     Abre  db\Pitaya_System.accdb  en Access                     │
│     Prueba los formularios y módulos modificados                │
│     Verifica que no hay errores en el Editor VBA (Alt+F11)      │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  5. PROMOVER — Copiar al Sistema oficial (manual)               │
│                                                                 │
│     Copia:  db\Pitaya_System.accdb                              │
│       →  C:\Users\...\Desktop\Sistema\Pitaya_System.accdb       │
│                                                                 │
│     (Asegúrate que el original no esté abierto antes de copiar) │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  6. RE-EXPORT — Sincronizar vba\ con el estado final            │
│                                                                 │
│     .\scripts\Export-AccessVBA.ps1                              │
│                                                                 │
│     (Asegura que git refleja exactamente lo que quedó en        │
│      el .accdb final, incluyendo ajustes automáticos de Access) │
└─────────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────────┐
│  7. PUSH — Documentar en GitHub                                 │
│                                                                 │
│     .\scripts\gitpush.ps1 "feat: descripción del cambio"        │
│                                                                 │
│     GitHub guarda el historial de qué cambió, cuándo y cómo.   │
└─────────────────────────────────────────────────────────────────┘
```

> **¿Por qué el paso 6 (re-export)?**  
> Access puede ajustar automáticamente el código al importar (por ejemplo,
> agregar líneas `Attribute`, reordenar declaraciones). El re-export captura
> esos ajustes para que git refleje el estado real del .accdb.

---

## Escenarios de uso

### Escenario A: Alguien actualizó el .accdb y me lo pasó

Alguien hizo cambios en Access, te envió el archivo actualizado y lo colocas en `Sistema\`.

```powershell
# 1. Desglosar el nuevo .accdb
.\scripts\Export-AccessVBA.ps1

# 2. git status mostrará solo los módulos que realmente cambiaron
git status

# 3. Revisar qué cambió (opcional pero recomendado)
git diff vba/

# 4. Subir la nueva versión a GitHub
.\scripts\gitpush.ps1 "chore: actualizar desde nueva versión de Pitaya_System"
```

---

### Escenario B: Quiero modificar código VBA con IA

```powershell
# 1. Exportar estado actual
.\scripts\Export-AccessVBA.ps1

# 2. Pídele a Antigravity/IA que modifique el .bas correspondiente
#    Ejemplo: "modifica vba\modulos\AccessHostinger.bas para..."

# 3. Revisar el cambio en VS Code antes de importar
# 4. Ver qué se importaría sin hacer nada
.\scripts\Import-AccessVBA.ps1 -DryRun

# 5. Importar solo los archivos modificados
.\scripts\Import-AccessVBA.ps1 -OnlyChanged

# 6. Probar db\Pitaya_System.accdb en Access

# 7. Si funciona, promover y documentar
#    (copiar db\ → Sistema\, re-exportar, push)
```

---

### Escenario C: Necesito revertir un cambio

```powershell
# Ver historial de un módulo específico
git log --oneline vba/modulos/AccessHostinger.bas

# Recuperar una versión anterior (reemplaza abc1234 con el commit deseado)
git checkout abc1234 -- vba/modulos/AccessHostinger.bas

# Importar solo ese archivo de vuelta
.\scripts\Import-AccessVBA.ps1 -OnlyChanged

# Probar, promover, re-exportar y push
```

---

### Escenario D: Primera instalación en una PC nueva

```powershell
# 1. Navegar a la carpeta correcta
cd "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya Web\VisualCode"

# 2. Clonar el repositorio
git clone https://github.com/MiguelGotea/Pitaya_System_Access.git

# 3. Verificar que el .accdb existe en Sistema\
Test-Path "C:\Users\TU_USUARIO\Desktop\Sistema\Pitaya_System.accdb"

# 4. Exportar para generar la copia de trabajo y sincronizar
cd Pitaya_System_Access
powershell -ExecutionPolicy Bypass -File .\scripts\Export-AccessVBA.ps1
```

---

## Versionado con Git

### Cómo git detecta cambios al re-exportar

Git **no** trata los archivos sobreescritos como archivos nuevos. Compara
el contenido real (hash SHA) del archivo anterior con el nuevo:

| Situación | Lo que ve Git |
|---|---|
| `ClientesClub.bas` con mismo contenido | Sin cambios — no aparece en `git status` |
| `ClientesClub.bas` con código diferente | Modificado — muestra las líneas exactas que cambiaron |
| Módulo nuevo en el .accdb | Archivo nuevo — aparece como `A` (added) en `git status` |
| Módulo eliminado del .accdb | Archivo borrado — aparece como `D` (deleted) en `git status` |

### Comandos de git útiles para este proyecto

```powershell
# Ver qué cambió desde el último commit
git status

# Ver las líneas exactas que cambiaron en un archivo
git diff vba/modulos/AccessHostinger.bas

# Ver el historial de cambios de un módulo
git log --oneline vba/modulos/AccessHostinger.bas

# Comparar dos versiones de un módulo
git diff abc1234..def5678 -- vba/modulos/AccessHostinger.bas

# Ver todos los archivos que cambiaron en un commit específico
git show --name-only abc1234

# Buscar en qué commit se introdujo una función
git log -S "NombreDeLaFuncion" vba/
```

---

## Advertencias importantes

> [!IMPORTANT]
> **Nunca ejecutes Import mientras el .accdb de `db\` está abierto en Access.**
> El script lanza su propia instancia de Access. Cierra cualquier instancia
> antes de ejecutar Import o Export.

> [!WARNING]
> **El re-export después del Import es obligatorio antes del push.**
> Si no lo haces, git podría documentar el estado de los `.bas` antes de que
> Access hiciera sus ajustes automáticos, generando inconsistencias.

> [!NOTE]
> **El diseño de formularios no se puede versionar aquí.**
> Si cambias la posición de un botón, agregas un campo de texto o modificas
> colores en un formulario, ese cambio vive solo en el `.accdb`. Solo el
> código VBA detrás del formulario está bajo control de versiones.

> [!TIP]
> **Usa `-OnlyChanged` en el Import para iteraciones rápidas.**
> Si solo modificaste 2-3 módulos, `-OnlyChanged` detecta cuáles cambiaron
> con `git diff` y solo importa esos, ahorrando varios minutos.
