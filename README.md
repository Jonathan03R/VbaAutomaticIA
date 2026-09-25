# VBA Dev

Edita macros y Ribbon Custom UI desde VS Code. Guarda archivo; herramienta sincroniza Excel.

## Antes de empezar

Necesitas Windows y Excel escritorio. En Excel activa una vez:

```text
Archivo → Opciones → Centro de confianza → Configuración de macros
→ Confiar en acceso al modelo de objetos de proyectos de VBA
```

## Elige tu camino

| Quiero... | Ejecuta |
|---|---|
| Editar macros de Excel existente | `vba dev .` |
| Crear Excel nuevo | `vba new .` |
| Editar pestañas, botones o iconos Ribbon | `vba ui .` |
| Ejecutar una vez y terminar | agrega `-Once` |

`.` significa carpeta actual.

## Macros VBA

Con libro `.xlsm` o `.xlsb` cerrado:

```powershell
vba dev .
```

Proyecto queda ordenado:

```text
MiProyecto/
├─ excel/     libro Excel
├─ src/       código que editas
│  ├─ modules/    .bas
│  ├─ classes/    .cls
│  ├─ forms/      .frm y .frx
│  └─ documents/  hojas y ThisWorkbook
└─ .vba/      configuración y respaldos
```

```text
vba dev . → editas src → Ctrl+S → Excel recibe cambio
Excel guarda código → src recibe cambio
```

Si ambos cambian mismo archivo, `src` gana.

```powershell
vba dev .          # abre y vigila Excel existente
vba dev . -Once    # sincroniza una vez y termina
vba new .          # crea estructura, Ribbon inicial y abre Excel para dev
vba new . -Once    # crea libro nuevo y termina
```

`vba new .` se ejecuta en una carpeta vacía. Prepara `excel/`, `src/` con sus carpetas VBA y `src/custom-ui/` con XML inicial, `_rels/` e `images/`; agrega la pestaña **Excel negocios**, con grupos de texto VBA, Custom UI y Proyecto, y después inicia `dev` con Excel abierto.

`Ctrl+C` guarda y cierra libro manejado. Si herramienta abrió Excel, también cierra Excel.

## Ribbon: botones e iconos

Excel debe estar **cerrado**.

```powershell
vba ui .
```

Guarda XML con `vba ui` activo; el libro se actualiza. Vuelve a abrirlo en Excel para ver el Ribbon.

Primera vez, extrae Ribbon existente a:

```text
src/custom-ui/
├─ customUI.xml o customUI14.xml
├─ _rels/                 relaciones de imágenes
└─ images/                iconos y subcarpetas
```

Si libro no tiene Ribbon, `vba ui` crea pestaña inicial **Excel negocios** con grupos de texto **VBA**, **Custom UI** y **Proyecto**, sin botones ni macros. También prepara `_rels/` e `images/` para ampliar el Ribbon.

```text
vba ui . abierto
editas XML, relaciones o imágenes
guardas
libro Excel se actualiza
```

También vigila cambios del libro:

```text
Excel → src/custom-ui
```

Crear, editar o eliminar XML, relaciones, imágenes y subcarpetas se sincroniza. Si ambos lados cambian mismo archivo, `src` gana.

```powershell
vba ui .          # sincroniza Ribbon y queda vigilando
vba ui . -Once    # extrae o aplica Ribbon una vez y termina
```

Sin carpeta `src/custom-ui`, extrae Ribbon existente o crea plantilla inicial cuando libro no tenga Ribbon. Con carpeta existente, aplica `src` a Excel.

## Reglas importantes

```text
vba dev .  = Excel abierto, macros
vba ui .   = Excel cerrado, Ribbon e iconos
```

No ejecutes ambos a la vez. Detén primero con `Ctrl+C`.

## Errores rápidos

| Mensaje | Qué hacer |
|---|---|
| Más de un `.xlsm` / `.xlsb` | Indica archivo: `vba dev ".\Ventas.xlsm"` |
| Excel abierto al usar `ui` | Cierra Excel completo |
| No contiene Custom UI Ribbon | Libro aún no tiene Ribbon personalizado |
| Proyecto VBA bloqueado | Escribe contraseña en diálogo Excel |
| Acceso VBA denegado | Activa permiso de Centro de confianza |

## Respaldos

Antes de cambiar VBA o Ribbon, copia queda en:

```text
.vba/backups/
```

Guía técnica: [GUIA-AVANZADA.md](docs/GUIA-AVANZADA.md).
