# VBA Dev

Herramienta de consola para editar VBA fuera de Excel. Recibe un libro existente, extrae sus componentes a archivos y mantiene el libro abierto mientras sincroniza los cambios guardados.

## Instalación en otra computadora

1. Instala los requisitos indicados más abajo: Windows y Excel de escritorio.
2. Clona este repositorio privado con una cuenta autorizada, o descarga su ZIP desde GitHub y extráelo, por ejemplo en `C:\Herramientas\VbaAutomaticIA`.
3. En **Variables de entorno > Variables de usuario > Path > Editar > Nuevo**, agrega `C:\Herramientas\VbaAutomaticIA`. Agrega la carpeta, no el archivo `vba.cmd`.
4. Cierra y vuelve a abrir VS Code o PowerShell.
5. Desde la carpeta que contiene tu Excel, ejecuta:

```powershell
vba dev .
```

Para crear un proyecto en una carpeta vacía, ejecuta `vba new .`. No necesitas instalar Python ni Node.js. Conserva juntos `vba.cmd`, `vba.ps1` y la carpeta `tools`.

Este repositorio distribuye código y documentación. Los libros Excel, recursos binarios de formularios, respaldos y el código extraído localmente en `src` quedan fuera del repositorio.

## Usar un Excel existente

Pon tu `.xlsm` o `.xlsb` en una carpeta propia. Desde la carpeta de esta herramienta ejecuta:

```powershell
.\vba.cmd dev "C:\MisProyectos\Ventas\Ventas.xlsm"
```

También puedes indicar la carpeta si contiene un único libro:

```powershell
.\vba.cmd dev "C:\MisProyectos\Ventas"
```

Si el libro está en la raíz del proyecto, `dev` lo mueve a `excel` y actualiza la configuración, también en proyectos existentes. Cierra el libro antes de este traslado. Si ya existe otro archivo con ese nombre en el destino, no se reemplaza. La primera ejecución extrae el código. Las siguientes conservan tus archivos y sincronizan los cambios pendientes, incluso los realizados con la herramienta apagada.

```text
Ventas/
├── excel/
│   └── Ventas.xlsm
├── src/
│   ├── modules/       # .bas: módulos estándar
│   ├── classes/       # .cls: clases
│   ├── forms/         # .frm y .frx: formularios y sus recursos
│   └── documents/     # .cls: hojas y ThisWorkbook
└── .vba/
    ├── project.json
    ├── state.json
    └── backups/       # copia del libro antes de importar cambios
```

Abre `src` con tu editor o IA. Al guardar un archivo, su componente se actualiza en Excel, normalmente en alrededor de un segundo cuando Excel está disponible. Con autoguardado en el editor, basta editar. Mantén el comando ejecutándose. `Ctrl+C` detiene la vigilancia y deja Excel abierto. También puedes cerrar el libro desde Excel: el comando detecta el cierre y termina. Si Excel muestra una pregunta de guardado, respóndela normalmente; la herramienta espera mientras Excel está ocupado.

`src` tiene prioridad para el código: sus cambios pendientes se importan al reiniciar, aunque Excel también haya cambiado. Mientras el proceso está encendido, los cambios de código hechos solamente en Excel se exportan a `src` al guardar el libro. Guardar archivos en `src` actualiza Excel. Si el mismo módulo cambia en ambos lados antes de sincronizar, gana `src`. Al iniciar el proceso se mantiene la prioridad de `src` para el código existente. Los componentes nuevos creados en Excel se extraen a `src`. Si eliminas un componente en Excel y guardas el libro, se elimina su archivo; si ese archivo tiene cambios locales pendientes, se conserva y vuelve a importarse a Excel. Los cambios en celdas se conservan; guardar una sincronización también guarda el resto del libro abierto.

## Crear un proyecto nuevo

Si ya estás dentro de una carpeta vacía, crea el proyecto allí mismo:

```powershell
& "C:\Users\USER\Documents\VbaAutomaticIA\vba.cmd" new .
```

Se crean `excel`, `src` y `.vba` directamente dentro de la carpeta actual. También puedes indicar una carpeta nueva o una carpeta existente vacía:

```powershell
.\vba.cmd new "C:\MisProyectos\MiMacro"
```

Crea `MiMacro\excel\MiMacro.xlsm`, un módulo `Main` de ejemplo, las carpetas de código y empieza a vigilar. Para continuar otro día:

```powershell
.\vba.cmd dev "C:\MisProyectos\MiMacro"
```

## Requisitos

- Windows con Excel de escritorio instalado.
- Windows PowerShell 5.1. El lanzador lo utiliza automáticamente, incluso desde PowerShell 7.
- En Excel: **Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Configuración de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA**.
- Conocer la contraseña, si el libro o proyecto VBA está protegido, y tener permiso de escritura.

## Contraseñas

Usa el mismo comando `vba.cmd dev`. Si el archivo requiere contraseña de apertura o escritura, Excel la solicita en su ventana antes de abrirlo. Si el proyecto VBA está bloqueado, la herramienta muestra el editor y abre el diálogo nativo de propiedades para que introduzcas su contraseña. Si después aparecen las propiedades, cierra ese cuadro: la sincronización continúa automáticamente al verificar el desbloqueo.

Son contraseñas distintas: un archivo puede pedir ambas. Escríbelas directamente en Excel; la herramienta no las recibe por consola ni las almacena en configuración, historial o registros. Después de abrir el diálogo VBA, la herramienta espera hasta cinco minutos a que Excel permita leer los módulos, incluso si Excel rechaza temporalmente las llamadas mientras escribes. Si cancelas el diálogo, pulsa `Ctrl+C` en consola para detener la espera. Si se agota el tiempo, Excel queda abierto para terminar el desbloqueo y reintentar; no se inicia la sincronización mientras siga bloqueado.

Si está desactivado el acceso al modelo de objetos VBA, primero debes habilitar el permiso descrito arriba. La contraseña no sustituye ese permiso. Si Office no permite mostrar su diálogo de propiedades, el mensaje indica cómo desbloquear manualmente el proyecto y reintentar.

La herramienta no cambia la configuración de seguridad de Office. Respeta la configuración de macros del usuario al abrir el libro y desactiva temporalmente eventos durante apertura y guardado para evitar ejecutar manejadores de eventos por la sincronización.

## Editar archivos

Los archivos de texto de `src` usan **UTF-8**. Conserva `Attribute VB_Name`, el nombre del archivo y los `.frx` asociados a formularios. Excel utiliza la página de códigos ANSI del sistema para importar módulos y formularios; si un carácter no puede representarse, la operación falla y conserva un respaldo.

Puedes organizar `modules`, `classes` y `forms` con subcarpetas propias, por ejemplo `modules\\ventas\\Facturas.bas`. La sincronización conserva esa ubicación aunque edites el componente desde Excel. Puedes usar un nombre descriptivo de archivo distinto al componente interno: `ClienteDllSunat.bas` puede conservar `Attribute VB_Name = "OK_FROMDLL"`. El nombre interno debe seguir siendo único en todo `src`; Excel no muestra estas carpetas ni el nombre físico del archivo.

Para agregar un módulo, crea por ejemplo `src\modules\Calculos.bas`:

```vb
Attribute VB_Name = "Calculos"
Option Explicit

Public Function Doble(ByVal valor As Double) As Double
    Doble = valor * 2
End Function
```

Eliminar el archivo de un módulo, clase o formulario elimina ese componente de Excel en la siguiente sincronización. Las hojas y `ThisWorkbook` no se eliminan: vacía su código conservando la cabecera. Los controles y el diseño visual de formularios siguen almacenados en el `.frx`; el código de sus eventos se edita en el `.frm`.

Al editar solamente código de un formulario, se conserva su diseño actual de Excel. Los cambios locales en `.frx` también se importan. Para extraer cambios hechos solamente en el diseñador visual de Excel, utiliza `Pull` explícito: los binarios exportados por Excel contienen bytes variables y no se usan para detectar conflictos del lado de Excel. Si editas el diseño en ambos lados, el `.frx` local modificado tiene prioridad.

El sincronizador importa código; no sustituye al compilador ni al depurador de VBA. Ejecuta y prueba tus macros en Excel. Si Excel está ocupado, muestra el error y reintenta cuando vuelve a estar disponible.

## Custom UI Ribbon

Custom UI controla pestañas y botones Ribbon desde XML. No es código `.bas`, `.cls` ni `.frm`.

Solo admite libros `.xlsm` que ya tengan Custom UI. Excel debe estar cerrado durante todo este modo:

```powershell
vba ui . -Once
```

La primera ejecución extrae las partes Ribbon existentes a `src\custom-ui\customUI.xml` y, si existe, `src\custom-ui\customUI14.xml`. Edita esos archivos UTF-8 en VS Code. Ejecuta el mismo comando nuevamente para aplicar XML local al libro. Antes de cambiarlo, se guarda copia completa en `.vba\backups`.

Para vigilar XML mientras editas, usa:

```powershell
vba ui .
```

El libro sigue cerrado; cada guardado XML válido actualiza el `.xlsm`. Después abre Excel normalmente para ver Ribbon. La herramienta no abre Excel ni genera capturas.

`vba dev .` y `vba ui .` son excluyentes. Si uno está activo, el otro se detiene sin modificar nada e indica en español que debes detener el primer comando con `Ctrl+C`.

Errores habituales:

- Libro abierto: ciérralo antes de ejecutar `vba ui`.
- XML inválido o raíz distinta de `customUI`: corrige XML; libro no se reemplaza.
- Libro sin Custom UI: primera versión no crea Ribbon desde cero.
- `.xlsb`: usa `.xlsm`; formato binario no está soportado en este modo.

## Una sola sincronización

```powershell
.\vba.cmd dev "C:\MisProyectos\Ventas" -Once
```

Extrae o sincroniza sin quedarse vigilando. Cierra solamente la instancia de Excel que haya creado para esa ejecución.

También puedes usar el script directamente:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\vba.ps1 dev "C:\MisProyectos\Ventas"
```

## Prioridad del código y recuperación

Puedes cerrar Excel, editar, crear o eliminar archivos de `src` y volver a ejecutar `dev .`: los cambios pendientes se aplican al libro. Conserva `.vba`, que registra el estado para detectar eliminaciones. Si cambió el código en ambos lados, gana `src`. Antes de modificar el libro se crea un respaldo.

Para conservar intencionalmente cambios de código hechos en Excel, detén el comando y usa `Pull` explícito. Los comandos siguientes reemplazan la otra versión:

```powershell
# Conservar archivos y llevarlos a Excel.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-Vba.ps1 -Workbook "C:\MisProyectos\Ventas\Ventas.xlsm" -SourceRoot "C:\MisProyectos\Ventas\src" -Direction Push

# Conservar Excel y reemplazar archivos.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-Vba.ps1 -Workbook "C:\MisProyectos\Ventas\Ventas.xlsm" -SourceRoot "C:\MisProyectos\Ventas\src" -Direction Pull
```

Después reinicia `vba.cmd dev`. Los respaldos anteriores a cada importación quedan en `.vba\backups`; no se borran automáticamente. Si una importación falla, se intenta restaurar el código anterior en memoria y se informa la ubicación del respaldo.

Cada proyecto usa un libro y una carpeta `src`. La extracción inicial no sobrescribe una carpeta de código que ya tenga archivos y carezca de estado de sincronización. Usa una carpeta nueva para tu libro existente, o elige explícitamente `Push`/`Pull` con el script anterior.

## Comprobar la herramienta

Pruebas de lógica y pruebas de integración con libros temporales de Excel:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-VbaDev.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-VbaDev.ps1 -Integration
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-VbaWatch.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-VbaAccess.ps1 -Integration
```

La prueba `Test-VbaWatch.ps1` abre Excel temporalmente para verificar el flujo real de vigilancia. Requiere un libro local `test.xlsm` con un módulo estándar. `Test-VbaDev.ps1 -Integration` requiere además los ejemplos locales `src\UserForm1.frm` y `src\UserForm1.frx`, con un control y código que contenga `"hola"`. Estos archivos de prueba no se distribuyen en GitHub. Las pruebas trabajan sobre copias; las pruebas de lógica sin `-Integration` no necesitan esos archivos ni abrir Excel.

La prueba de acceso simula contraseña aceptada, espera con diálogo abierto, errores COM transitorios, agotamiento del tiempo y permiso VBA denegado; con `-Integration` verifica además que Excel exponga el comando nativo de propiedades. No introduce contraseñas en diálogos reales. Para comprobarlo manualmente, ejecuta `dev` sobre una copia de un libro protegido, espera unos segundos antes de escribir la contraseña en Excel, cierra las propiedades si aparecen y comprueba que se extraiga `src`. Repite cancelando y pulsa `Ctrl+C`: no debe iniciarse la sincronización.

El antiguo `tools\Clean-VbaTest.ps1` es una utilidad destructiva de limpieza de ejemplos y no participa en este flujo.
